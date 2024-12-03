public class Cell(string content)
{
	public string Content { get; set; } = content;
	public int ColumnIndex { get; set; }
	public int RowIndex { get; set; }
	public int ColSpan { get; set; } = 1;
	public int RowSpan { get; set; } = 1;
}