using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace XrmToolBox.DocumentTemplateDeploymentManager
{
    internal sealed class DocumentTemplateDeploymentService
    {
        private static readonly Regex NamespaceRegex = new Regex(
            @"urn:microsoft-crm/document-template/([a-zA-Z][a-zA-Z0-9_]*)/([0-9]+)/",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public IReadOnlyList<TemplateItem> GetWordTemplates(IOrganizationService source)
        {
            var query = new QueryExpression("documenttemplate")
            {
                ColumnSet = new ColumnSet(
                    "documenttemplateid", "name", "associatedentitytypecode", "documenttype",
                    "languagecode", "content", "clientdata", "description", "createdon", "modifiedon"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };
            query.Criteria.AddCondition("documenttype", ConditionOperator.Equal, 2);
            query.AddOrder("name", OrderType.Ascending);
            return source.RetrieveMultiple(query).Entities.Select(e => new TemplateItem { Record = e }).ToList();
        }

        public MoveResult Deploy(IOrganizationService target, string targetName, Entity sourceTemplate)
        {
            var name = sourceTemplate.GetAttributeValue<string>("name");
            var content = sourceTemplate.GetAttributeValue<string>("content");
            if (string.IsNullOrWhiteSpace(content)) throw new InvalidOperationException("Source content is null or empty.");

            byte[] sourceBytes;
            try { sourceBytes = Convert.FromBase64String(content); }
            catch (FormatException ex) { throw new InvalidOperationException("Source content is not valid Base64.", ex); }

            var associatedTable = ReadAssociatedTable(sourceBytes);
            var recordTable = sourceTemplate.GetAttributeValue<string>("associatedentitytypecode");
            if (!string.Equals(associatedTable, recordTable, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"DOCX table '{associatedTable}' does not match record table '{recordTable}'.");

            var targetTypeCode = GetObjectTypeCode(target, associatedTable);
            var reboundBytes = Rebind(sourceBytes, associatedTable, targetTypeCode);
            var validation = Validate(reboundBytes, associatedTable, targetTypeCode);
            if (!validation.Valid) throw new InvalidOperationException(validation.Message);

            var matches = FindMatches(target, name, associatedTable);
            if (matches.Count > 1)
                throw new InvalidOperationException("Multiple matching Target templates exist. No record was changed. IDs: " + string.Join(", ", matches.Select(e => e.Id)));

            var payload = new Entity("documenttemplate");
            payload["name"] = name;
            payload["associatedentitytypecode"] = associatedTable;
            payload["documenttype"] = 2;
            payload["languagecode"] = sourceTemplate.GetAttributeValue<int?>("languagecode") ?? 1033;
            payload["content"] = Convert.ToBase64String(reboundBytes);
            CopyOptional(sourceTemplate, payload, "clientdata");
            CopyOptional(sourceTemplate, payload, "description");

            string action;
            Guid targetId;
            if (matches.Count == 1)
            {
                targetId = matches[0].Id;
                payload.Id = targetId;
                target.Update(payload);
                action = "Updated";
            }
            else
            {
                targetId = target.Create(payload);
                action = "Created";
            }

            var stored = target.Retrieve("documenttemplate", targetId, new ColumnSet("content", "name", "associatedentitytypecode", "documenttype"));
            var storedContent = stored.GetAttributeValue<string>("content");
            if (string.IsNullOrWhiteSpace(storedContent)) throw new InvalidOperationException("Target content was null after deployment.");
            var storedBytes = Convert.FromBase64String(storedContent);
            var storedValidation = Validate(storedBytes, associatedTable, targetTypeCode);
            if (!storedValidation.Valid) throw new InvalidOperationException("Post-upload validation failed: " + storedValidation.Message);
            if (!reboundBytes.SequenceEqual(storedBytes)) throw new InvalidOperationException("Target content does not exactly match the uploaded DOCX package.");

            return new MoveResult
            {
                Target = targetName,
                Template = name,
                Action = action,
                Success = true,
                Bindings = storedValidation.Bindings,
                Message = $"{action} {targetId}; target type code {targetTypeCode}; validated {storedValidation.Bindings} bindings."
            };
        }

        private static void CopyOptional(Entity source, Entity target, string attribute)
        {
            if (source.Attributes.Contains(attribute) && source[attribute] != null) target[attribute] = source[attribute];
        }

        private static List<Entity> FindMatches(IOrganizationService service, string name, string associatedTable)
        {
            var query = new QueryExpression("documenttemplate")
            {
                ColumnSet = new ColumnSet("documenttemplateid"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, name);
            query.Criteria.AddCondition("associatedentitytypecode", ConditionOperator.Equal, associatedTable);
            query.Criteria.AddCondition("documenttype", ConditionOperator.Equal, 2);
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        private static int GetObjectTypeCode(IOrganizationService service, string logicalName)
        {
            var response = (RetrieveEntityResponse)service.Execute(new RetrieveEntityRequest
            {
                LogicalName = logicalName,
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = true
            });
            if (!response.EntityMetadata.ObjectTypeCode.HasValue)
                throw new InvalidOperationException($"Target table '{logicalName}' has no ObjectTypeCode.");
            return response.EntityMetadata.ObjectTypeCode.Value;
        }

        private static string ReadAssociatedTable(byte[] docx)
        {
            var xml = ReadEntry(docx, "customXml/item1.xml");
            var matches = NamespaceRegex.Matches(xml).Cast<Match>().Select(m => m.Groups[1].Value.ToLowerInvariant()).Distinct().ToList();
            if (matches.Count != 1) throw new InvalidOperationException($"Expected one associated table in the DOCX; found {matches.Count}.");
            return matches[0];
        }

        private static byte[] Rebind(byte[] docx, string table, int targetTypeCode)
        {
            using (var stream = new MemoryStream())
            {
                stream.Write(docx, 0, docx.Length);
                stream.Position = 0;
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true))
                {
                    var itemEntry = RequireEntry(archive, "customXml/item1.xml");
                    var documentEntry = RequireEntry(archive, "word/document.xml");
                    var itemXml = ReadEntry(itemEntry);
                    var sourceNamespaces = NamespaceRegex.Matches(itemXml).Cast<Match>()
                        .Where(m => string.Equals(m.Groups[1].Value, table, StringComparison.OrdinalIgnoreCase))
                        .Select(m => m.Value).Distinct().ToList();
                    if (sourceNamespaces.Count != 1) throw new InvalidOperationException($"Expected one CRM namespace for '{table}'; found {sourceNamespaces.Count}.");
                    var sourceNamespace = sourceNamespaces[0];
                    var targetNamespace = $"urn:microsoft-crm/document-template/{table}/{targetTypeCode}/";
                    ReplaceEntry(itemEntry, itemXml, sourceNamespace, targetNamespace);
                    ReplaceEntry(documentEntry, ReadEntry(documentEntry), sourceNamespace, targetNamespace);
                }
                return stream.ToArray();
            }
        }

        private static ValidationResult Validate(byte[] docx, string table, int targetTypeCode)
        {
            var expected = $"urn:microsoft-crm/document-template/{table}/{targetTypeCode}/";
            var itemXml = ReadEntry(docx, "customXml/item1.xml");
            var documentXml = ReadEntry(docx, "word/document.xml");
            var combined = itemXml + "\n" + documentXml;
            var namespaces = NamespaceRegex.Matches(combined).Cast<Match>()
                .Where(m => string.Equals(m.Groups[1].Value, table, StringComparison.OrdinalIgnoreCase))
                .Select(m => m.Value).Distinct().ToList();
            var bindings = Regex.Matches(documentXml, @"<w:dataBinding\b").Count;
            var occurrences = Regex.Matches(combined, Regex.Escape(expected)).Count;
            var valid = namespaces.Count == 1 && namespaces[0] == expected && bindings > 0 && occurrences == bindings + 1;
            return new ValidationResult
            {
                Valid = valid,
                Bindings = bindings,
                Message = valid ? "Valid" : $"Expected namespace '{expected}' in one custom XML root plus all bindings; found {occurrences} occurrences, {bindings} bindings, and {namespaces.Count} distinct namespaces."
            };
        }

        private static string ReadEntry(byte[] docx, string name)
        {
            using (var stream = new MemoryStream(docx, false))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                return ReadEntry(RequireEntry(archive, name));
        }

        private static ZipArchiveEntry RequireEntry(ZipArchive archive, string name)
        {
            var entry = archive.GetEntry(name);
            if (entry == null) throw new InvalidOperationException($"DOCX part '{name}' is missing.");
            return entry;
        }

        private static string ReadEntry(ZipArchiveEntry entry)
        {
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true, 1024, false)) return reader.ReadToEnd();
        }

        private static void ReplaceEntry(ZipArchiveEntry entry, string text, string oldValue, string newValue)
        {
            var count = Regex.Matches(text, Regex.Escape(oldValue)).Count;
            if (count == 0) throw new InvalidOperationException($"DOCX part '{entry.FullName}' does not contain the source namespace.");
            using (var target = entry.Open())
            {
                target.SetLength(0);
                using (var writer = new StreamWriter(target, new UTF8Encoding(false), 1024, true)) writer.Write(text.Replace(oldValue, newValue));
            }
        }

        private sealed class ValidationResult
        {
            public bool Valid { get; set; }
            public int Bindings { get; set; }
            public string Message { get; set; }
        }
    }
}
