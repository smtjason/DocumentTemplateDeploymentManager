# Document Template Deployment Manager

Document Template Deployment Manager safely deploys Dataverse Word document templates from one environment to another. It repairs the environment-specific table type code embedded in the DOCX XML bindings before upload, preventing generated documents from displaying field names instead of merged values.

## Features

- Uses the active connection as Source and the secondary connection as Target.
- Lists Word templates with name, created date, modified date, and associated table.
- Supports sortable columns and multi-select deployment.
- Updates exactly one matching Target record by name, associated table, and Word document type.
- Creates a Target record when no match exists.
- Stops safely if duplicate Target records make the operation ambiguous.
- Rebinds both `customXml/item1.xml` and all `word/document.xml` data bindings to the Target table ObjectTypeCode.
- Downloads the stored Target content and validates its package, namespace, and binding count.
- Copies readable deployment results for email, documents, and chat.
- Never deletes document templates.

## Use

1. Open the tool with two different Dataverse connections.
2. Choose the active/primary connection as Source.
3. Choose the secondary connection as Target.
4. Select one or more Word templates.
5. Select **Deploy selected**.
6. Review or copy the result rows.

The tool refuses to run when Source and Target refer to the same environment.

## Matching behavior

- **Updated**: exactly one matching Target record exists.
- **Created**: no matching Target record exists.
- **Failed**: validation, metadata lookup, duplicate matching, upload, or post-upload verification fails.

## Local development

```powershell
npm ci
npm test
npm run build
npm run validate
```

Enable **Show Debug Menu** in Power Platform ToolBox, open **Debug**, and load this directory under **Load Local Tool**.

## Privacy

The tool communicates only through the Power Platform ToolBox Dataverse API. Do not commit exported customer templates, connection information, credentials, or access tokens to this repository.
