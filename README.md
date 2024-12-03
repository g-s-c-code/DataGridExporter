# MudBlazor DataGrid Excel Exporter

## Overview

This library provides Excel export functionality for MudBlazor DataGrid components in .NET applications, allowing easy conversion of grid data into Excel spreadsheets. Link: [datagridexporter.netlify.app](http://datagridexporter.netlify.app/)

## Acknowledgments

This project is substantially based on the work originally developed by [timmac-qmc](https://github.com/timmac-qmc). Credit is given to the original author for the foundational implementation.

## Features

- Easy one-method export for MudBlazor DataGrids
- Supports both simple and extended sheet formats
- Automatic number and text cell type detection
- Configurable column visibility
- Browser-based file download using JavaScript interop

## Installation

### Dependencies
- Microsoft.AspNetCore.Components
- MudBlazor
- DocumentFormat.OpenXml

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

<MudButton OnClick="ExportToExcel">Export to Excel</MudButton>

@code {
    private async Task ExportToExcel()
    {
        await grid.ExportMudDataGrid(JSRuntime, "exported_data.xlsx");
    }
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
