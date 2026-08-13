# Document Template Deployment Manager

[![Build XrmToolBox tool](https://github.com/smtjason/DocumentTemplateDeploymentManager/actions/workflows/build-xrmtoolbox.yml/badge.svg)](https://github.com/smtjason/DocumentTemplateDeploymentManager/actions/workflows/build-xrmtoolbox.yml)
[![NuGet](https://img.shields.io/nuget/v/XrmToolBox.DocumentTemplateDeploymentManager.svg)](https://www.nuget.org/packages/XrmToolBox.DocumentTemplateDeploymentManager/)

An XrmToolBox plugin for moving Microsoft Word document templates between Dataverse environments while repairing the environment-specific `ObjectTypeCode` embedded in DOCX XML bindings.

This repository is organized as one product with platform-specific implementations:

- `src/XrmToolBox` contains the production XrmToolBox implementation.
- `src/PowerPlatformToolBox` is reserved for the planned Power Platform ToolBox implementation.
- `assets` contains shared product artwork.

## Features

- Uses the main XrmToolBox connection as Source.
- Lists Source Word templates with name, created date, modified date, and associated table.
- Sorts the Source template list by any displayed column, with ascending/descending header indicators.
- Supports selecting one or more templates.
- Supports adding and selecting one or more Target connections.
- Updates an exact Target match by name, associated table, and Word document type.
- Creates a Target record only when no match exists.
- Stops safely when duplicate Target matches make the operation ambiguous.
- Rebinds `customXml/item1.xml` and every `word/document.xml` data binding to the Target table's `ObjectTypeCode`.
- Downloads and validates the Target content after each create/update.
- Shows a result row for every template/Target combination.
- Copies the complete deployment-results table from its right-click menu (or Ctrl+C), using a formatted HTML table for Outlook/Word and Markdown for chat or plain text.

## Build

Requirements:

- Visual Studio 2022 or the .NET SDK with .NET Framework 4.8 targeting support
- NuGet access

```powershell
dotnet build .\src\XrmToolBox\DocumentTemplateDeploymentManager.csproj --configuration Release
```

Copy `XrmToolBox.DocumentTemplateDeploymentManager.dll` from `src\XrmToolBox\bin\Release\net48` into the XrmToolBox `Plugins` folder for local testing.

Or run:

```powershell
& .\src\XrmToolBox\Install-Local.ps1 -XrmToolBoxDirectory 'C:\Path\To\XrmToolBox'
```

## Safety behavior

The plugin never deletes templates. A Target operation is either:

- `Updated`: exactly one matching record existed.
- `Created`: no matching record existed.
- `Failed`: source content, metadata, matching, upload, or post-upload validation failed.

## Privacy and test data

Do not commit exported Dataverse records, downloaded customer templates, environment URLs, connection details, credentials, or access tokens. Use synthetic fixtures for public automated tests.
