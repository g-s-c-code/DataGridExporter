# MudBlazor DataGrid Exporter

## Overview

This library provides export functionality for MudBlazor DataGrid components in .NET applications, allowing easy conversion of grid data into `.xlsx`, `.csv`, and `.json` formats. Link: [datagridexporter.netlify.app](https://datagridexporter.netlify.app/)

## Acknowledgments

This project is substantially based on the work originally developed by [timmac-qmc](https://github.com/timmac-qmc). Credit is given to the original author for the foundational implementation.

## Features

- Easy one-method export for MudBlazor DataGrids
- Supports `.xlsx`, `.csv`, and `.json` formats
- Configurable column visibility
- Automatic number and text cell type detection (for `.xlsx`)
- Browser-based file download using JavaScript interop

## Installation

### Dependencies

- Microsoft.AspNetCore.Components
- MudBlazor
- DocumentFormat.OpenXml (for `.xlsx` exports)

### Setup

1. Install the required NuGet packages (see dependencies)
2. Copy the project files into your solution
3. Ensure the JavaScript file (`saveFileAs.js`) is included and referenced in your project

## Usage Example

```csharp

@inject IJSRuntime JSRuntime

<MudDataGrid @ref="grid" Items="@items">

    <!-- Your grid columns -->

</MudDataGrid>

<MudButton OnClick="ExportToXlsx">Export to Excel</MudButton>
<MudButton OnClick="ExportToCsv">Export to CSV</MudButton>
<MudButton OnClick="ExportToJson">Export to JSON</MudButton>

@code {
    private async Task ExportToExcel() => await _grid.ExportToExcel(JSRuntime, "export.xlsx");
    private async Task ExportToCsv() => await _grid.ExportToCsv(JSRuntime, "export.csv");
    private async Task ExportToJson() => await _grid.ExportToJson(JSRuntime, "export.json");
}

```

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Open a Pull Request

## License

Distributed under the MIT License.
