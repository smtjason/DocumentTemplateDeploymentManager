using System;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;

namespace XrmToolBox.DocumentTemplateDeploymentManager
{
    internal sealed class TemplateItem
    {
        public Entity Record { get; set; }
        public string Name => Record.GetAttributeValue<string>("name");
        public string AssociatedTable => Record.GetAttributeValue<string>("associatedentitytypecode");
        public DateTime? CreatedOn => Record.GetAttributeValue<DateTime?>("createdon");
        public DateTime? ModifiedOn => Record.GetAttributeValue<DateTime?>("modifiedon");
    }

    internal sealed class TargetItem
    {
        public ConnectionDetail Detail { get; set; }
        public override string ToString() => Detail?.ConnectionName ?? "Unnamed connection";
    }

    internal sealed class MoveResult
    {
        public string Target { get; set; }
        public string Template { get; set; }
        public string Action { get; set; }
        public bool Success { get; set; }
        public int Bindings { get; set; }
        public string Message { get; set; }
    }
}
