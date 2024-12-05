using System.Text;
using Microsoft.JSInterop;
using MudBlazor;

public static class MudDataGridFileExporter
{
	public static async Task ExportMudDataGrid<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename, string format)
	{
		byte[] content;
		switch (format.ToLower())
		{
			case "xlsx":
				var excelWriter = new CellWriter();
				content = excelWriter.GenerateSpreadsheet(grid.RenderedColumns, grid.FilteredItems);
				break;
			case "csv":
				content = GenerateCsv(grid.RenderedColumns, grid.FilteredItems);
				break;
			case "json":
				content = GenerateJson(grid.FilteredItems);
				break;
			default:
				throw new ArgumentException("Unsupported format");
		}

		var mimeType = format.ToLower() switch
		{
			"xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			"csv" => "text/csv",
			"json" => "application/json",
			_ => "application/octet-stream"
		};

		await js.InvokeVoidAsync("saveAsFile", filename, Convert.ToBase64String(content), mimeType);
	}


	public static byte[] GenerateCsv<T>(List<Column<T>> columns, IEnumerable<T> items)
	{
		var csvBuilder = new StringBuilder();
		// Add headers
		var header = string.Join(",", columns.Where(c => !c.Hidden).Select(c => c.Title));
		csvBuilder.AppendLine(header);

		// Add rows
		foreach (var item in items)
		{
			var row = string.Join(",", columns.Where(c => !c.Hidden)
				.Select(c => {
					var prop = item.GetType().GetProperty(c.PropertyName!);
					return prop?.GetValue(item)?.ToString()?.Replace(",", " ") ?? string.Empty;
				}));
			csvBuilder.AppendLine(row);
		}

		return Encoding.UTF8.GetBytes(csvBuilder.ToString());
	}

	public static byte[] GenerateJson<T>(IEnumerable<T> items)
	{
		var json = System.Text.Json.JsonSerializer.Serialize(items);
		return Encoding.UTF8.GetBytes(json);
	}

}