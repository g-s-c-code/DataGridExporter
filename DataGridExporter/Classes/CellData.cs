public class CellData(bool extendedSheet)
{
	public CellData() : this(false) { }
	public List<List<Cell>> Cells { get; set; } = new List<List<Cell>>();
	public string? SheetName { get; set; }
	public bool ExtendedSheet { get; set; } = extendedSheet;
}
