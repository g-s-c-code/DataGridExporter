using Microsoft.JSInterop;
using MudBlazor;

public static class MudDataGridFileExporter
{
	public static async Task ExportMudDataGrid<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename)
	{
		CellWriter spreadsheetlWriter = new CellWriter();
		byte[] content = spreadsheetlWriter.GenerateSpreadsheet(grid.RenderedColumns, grid.FilteredItems);

		await JSRuntimeExtensions.InvokeAsync<object>(js, "saveAsFile", new object[2]
		{
					filename, Convert.ToBase64String(content)
		});
	}
}