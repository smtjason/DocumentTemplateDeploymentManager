# Contributing

Contributions and issue reports are welcome.

## Development rules

- Never include tenant data, exported Dataverse records, environment URLs, credentials, tokens, or customer document templates.
- Use synthetic test fixtures that contain no organization-specific names or schema.
- Keep the XrmToolBox and Power Platform ToolBox implementations behaviorally consistent.
- Build and test the affected platform implementation before opening a pull request.
- Document user-visible changes in the pull request and release notes.

## XrmToolBox build

```powershell
dotnet build .\src\XrmToolBox\DocumentTemplateDeploymentManager.csproj --configuration Release
```
