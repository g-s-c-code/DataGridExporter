using System.Text;
using Microsoft.JSInterop;
using MudBlazor;

public static class MudDataGridFileExporter
{
	public static async Task ExportToExcel<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename)
	{
		var excelWriter = new CellWriter();
		byte[] content = excelWriter.GenerateSpreadsheet(grid.RenderedColumns, grid.FilteredItems);

		await js.InvokeVoidAsync("saveAsFile", filename, Convert.ToBase64String(content),
								  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	}

	public static async Task ExportToCsv<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename)
	{
		byte[] content = GenerateCsv(grid.RenderedColumns, grid.FilteredItems);

		await js.InvokeVoidAsync("saveAsFile", filename, Convert.ToBase64String(content), "text/csv");
	}

	public static async Task ExportToJson<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename)
	{
		byte[] content = GenerateJson(grid.FilteredItems);

		await js.InvokeVoidAsync("saveAsFile", filename, Convert.ToBase64String(content), "application/json");
	}

	private static byte[] GenerateCsv<T>(List<Column<T>> columns, IEnumerable<T> items)
	{
		var csvBuilder = new StringBuilder();
		// Add headers
		var header = string.Join(",", columns.Where(c => !c.Hidden).Select(c => c.Title));
		csvBuilder.AppendLine(header);

		// Add rows
		foreach (var item in items)
		{
			var row = string.Join(",", columns.Where(c => !c.Hidden)
				.Select(c =>
				{
					var prop = item.GetType().GetProperty(c.PropertyName!);
					return prop?.GetValue(item)?.ToString()?.Replace(",", " ") ?? string.Empty;
				}));
			csvBuilder.AppendLine(row);
		}

		return Encoding.UTF8.GetBytes(csvBuilder.ToString());
	}

	private static byte[] GenerateJson<T>(IEnumerable<T> items)
	{
		var json = System.Text.Json.JsonSerializer.Serialize(items);
		return Encoding.UTF8.GetBytes(json);
	}
}
