import { associatedTable, rebind, validate } from "./docx";
import type { DeploymentResult, TemplateRecord } from "./types";

const api = () => window.dataverseAPI;
const esc = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

export async function loadTemplates(): Promise<TemplateRecord[]> {
  const fetch = `<fetch><entity name="documenttemplate"><attribute name="documenttemplateid"/><attribute name="name"/><attribute name="associatedentitytypecode"/><attribute name="documenttype"/><attribute name="languagecode"/><attribute name="content"/><attribute name="clientdata"/><attribute name="description"/><attribute name="createdon"/><attribute name="modifiedon"/><filter><condition attribute="documenttype" operator="eq" value="2"/></filter><order attribute="name"/></entity></fetch>`;
  return (await api().fetchXmlQuery(fetch)).value as unknown as TemplateRecord[];
}
async function findMatches(name: string, targetTypeCode: number): Promise<TemplateRecord[]> {
  const fetch = `<fetch><entity name="documenttemplate"><attribute name="documenttemplateid"/><filter><condition attribute="name" operator="eq" value="${esc(name)}"/><condition attribute="associatedentitytypecode" operator="eq" value="${targetTypeCode}"/><condition attribute="documenttype" operator="eq" value="2"/></filter></entity></fetch>`;
  return (await api().fetchXmlQuery(fetch, "secondary")).value as unknown as TemplateRecord[];
}
export async function deployTemplate(source: TemplateRecord, targetName: string): Promise<DeploymentResult> {
  try {
    if (!source.content) throw new Error("Source content is null or empty.");
    const table = await associatedTable(source.content);
    if (table.toLowerCase() !== source.associatedentitytypecode.toLowerCase()) throw new Error(`DOCX table '${table}' does not match record table '${source.associatedentitytypecode}'.`);
    const metadata = await api().getEntityMetadata(table, true, ["ObjectTypeCode"], "secondary");
    const typeCode = Number(metadata.ObjectTypeCode); if (!Number.isInteger(typeCode)) throw new Error(`Target table '${table}' has no ObjectTypeCode.`);
    const rebound = await rebind(source.content, table, typeCode); const matches = await findMatches(source.name, typeCode);
    if (matches.length > 1) throw new Error(`Multiple matching target templates exist. No record was changed. IDs: ${matches.map(x => x.documenttemplateid).join(", ")}`);
    const payload: Record<string, unknown> = { name: source.name, associatedentitytypecode: table, documenttype: 2, languagecode: source.languagecode ?? 1033, content: rebound.content };
    if (source.clientdata != null) payload.clientdata = source.clientdata; if (source.description != null) payload.description = source.description;
    let id: string; let action: string;
    if (matches.length === 1) { id = matches[0].documenttemplateid; await api().update("documenttemplate", id, payload, "secondary"); action = "Updated"; }
    else { const created = await api().create("documenttemplate", payload, "secondary"); id = created.id; action = "Created"; }
    const stored = await api().retrieve("documenttemplate", id, ["content", "name", "associatedentitytypecode", "documenttype"], "secondary") as unknown as TemplateRecord;
    if (!stored.content) throw new Error("Target content was null after deployment.");
    const checked = await validate(stored.content, table, typeCode); if (!checked.valid) throw new Error(`Post-upload validation failed: ${checked.message}`);
    if (stored.content !== rebound.content) throw new Error("Target content does not exactly match the uploaded DOCX package.");
    return { target: targetName, template: source.name, action, success: true, bindings: checked.bindings, message: `${action} ${id}; target type code ${typeCode}; validated ${checked.bindings} bindings.` };
  } catch (error) { return { target: targetName, template: source.name, action: "Failed", success: false, bindings: null, message: error instanceof Error ? error.message : String(error) }; }
}
