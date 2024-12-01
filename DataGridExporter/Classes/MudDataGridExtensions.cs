using Microsoft.JSInterop;
using MudBlazor;

public static class MudDataGridExtensions
{
	public static async Task DownloadExcel<T>(this MudDataGrid<T> grid, IJSRuntime js, string filename)
	{
		ExcelWriter excelWriter = new ExcelWriter();
		byte[] content = excelWriter.GenerateExcel(grid.RenderedColumns, grid.FilteredItems);

		await JSRuntimeExtensions.InvokeAsync<object>(js, "saveAsFile", new object[2]
		{
					filename,
					Convert.ToBase64String(content)
		});
	}
}