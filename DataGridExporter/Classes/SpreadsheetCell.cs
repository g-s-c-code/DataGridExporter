using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Reflection;

namespace MudBlazor
{
	public class Cell(string content)
	{
		public string Content { get; set; } = content;
		public int ColumnIndex { get; set; }
		public int RowIndex { get; set; }
		public int ColSpan { get; set; } = 1;
		public int RowSpan { get; set; } = 1;
	}

	public class CellStatus
	{
		public string? Message { get; set; }
		public bool Success
		{
			get { return string.IsNullOrWhiteSpace(Message); }
		}
	}

	public class CellData(bool extendedSheet)
	{
		public CellData() : this(false) { }
		public CellStatus Status { get; set; } = new CellStatus();
		public Columns? ColumnConfigurations { get; set; }
		public List<List<Cell>> Cells { get; set; } = new List<List<Cell>>();
		public string? SheetName { get; set; }
		public bool ExtendedSheet { get; set; } = extendedSheet;
	}

	public class CellWriter
	{
		private string ColumnLetter(int colIndex)
		{
			int div = colIndex + 1;
			string colLetter = string.Empty;
			int mod = 0;

			while (div > 0)
			{
				mod = (div - 1) % 26;
				colLetter = (char)(65 + mod) + colLetter;
				div = (int)((div - mod) / 26);
			}
			return colLetter;
		}

		private DocumentFormat.OpenXml.Spreadsheet.Cell CreateCell(string header, UInt32 index, string text)
		{
			DocumentFormat.OpenXml.Spreadsheet.Cell cell;

			if (double.TryParse(text, out double number))
			{
				cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
				{
					DataType = CellValues.Number,
					CellReference = header + index,
					CellValue = new CellValue(number.ToString(CultureInfo.InvariantCulture))
				};
			}
			else
			{
				cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
				{
					DataType = CellValues.InlineString,
					CellReference = header + index
				};

				var istring = new InlineString();
				var t = new Text { Text = text };
				istring.AppendChild(t);
				cell.AppendChild(istring);
			}

			return cell;
		}

		private MergeCell CreateMergeCell(Cell cell)
		{
			MergeCell mergeCell = new MergeCell
			{
				Reference = new StringValue(ColumnLetter(cell.ColumnIndex)
						+ (cell.RowIndex + 1) + ":"
						+ ColumnLetter(cell.ColumnIndex + cell.ColSpan - 1) +
						+(cell.RowIndex + cell.RowSpan)),
			};
			return mergeCell;
		}

		public byte[] GenerateSpreadsheet<T>(List<Column<T>> columns, IEnumerable<T> items)
		{
			CellData cellData = new CellData();
			cellData.SheetName = "Items";
			var header = new List<Cell>();

			foreach (var column in columns)
			{
				if (!column.Hidden && !string.IsNullOrEmpty(column.PropertyName))
				{
					header.Add(new Cell(column.Title));
				}
			}
			cellData.Cells.Add(header);

			if (items != null)
			{
				foreach (var item in items)
				{
					Type t = item!.GetType();
					List<Cell> row = new List<Cell>();
					foreach (var column in columns)
					{
						if (!column.Hidden)
						{
							if (!string.IsNullOrEmpty(column.PropertyName))
							{
								PropertyInfo prop = t.GetProperty(column.PropertyName)!;
								object val = prop.GetValue(item) ?? new string("N/A");
								row.Add(new Cell(val.ToString()!));
							}
						}
					}
					cellData.Cells.Add(row);
				}
			}
			return GenerateSpreadsheet(cellData);
		}

		public byte[] GenerateSpreadsheet(CellData data)
		{
			var stream = new MemoryStream();
			var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

			var workbookpart = document.AddWorkbookPart();
			workbookpart.Workbook = new Workbook();
			var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
			var sheetData = new SheetData();

			worksheetPart.Worksheet = new Worksheet(sheetData);

			var sheets = document.WorkbookPart.Workbook.
				AppendChild(new Sheets());

			var sheet = new Sheet()
			{
				Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = 1,
				Name = data.SheetName ?? "sheet_1"
			};
			sheets.AppendChild(sheet);

			if (data.ExtendedSheet)
				AppendDataToExtendedSheet(worksheetPart.Worksheet, sheetData, data);
			else
				AppendDataToSheet(sheetData, data);

			workbookpart.Workbook.Save();
			document.Dispose();

			return stream.ToArray();
		}

		private void AppendDataToSheet(SheetData sheetData, CellData data)
		{
			UInt32 rowIdex = 0;
			Row row;

			// Add sheet data
			foreach (var rowData in data.Cells)
			{
				int cellIdex = 0;
				row = new Row { RowIndex = ++rowIdex };
				sheetData.AppendChild(row);
				foreach (var cellData in rowData)
				{
					var cell = CreateCell(ColumnLetter(cellIdex++), rowIdex,
						cellData.Content ?? string.Empty);
					row.AppendChild(cell);
				}
			}
		}

		private void AppendDataToExtendedSheet(Worksheet worksheet, SheetData sheetData,
			CellData cellData)
		{
			UInt32 rowIdex = 0;
			Row row;
			MergeCells mergeCells = new MergeCells();

			// Add sheet data
			foreach (var rowData in cellData.Cells)
			{
				row = new Row { RowIndex = ++rowIdex };
				sheetData.AppendChild(row);
				foreach (var data in rowData)
				{
					var cell = CreateCell(ColumnLetter(data.ColumnIndex),
						(uint)(data.RowIndex + 1), data.Content ?? string.Empty);
					row.AppendChild(cell);
					if (data.ColSpan > 1 || data.RowSpan > 1)
					{
						var mergeCell = CreateMergeCell(data);
						mergeCells.Append(mergeCell);
					}
				}
			}
			worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
		}
	}
}
