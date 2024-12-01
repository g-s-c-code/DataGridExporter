using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Reflection;

namespace MudBlazor
{
	public class Cell
	{
		public string Content { get; set; }
		public int ColumnIndex { get; set; }
		public int RowIndex { get; set; }
		public int ColSpan { get; set; }
		public int RowSpan { get; set; }

		public Cell(string content)
		{
			Content = content;
			ColSpan = 1;
			RowSpan = 1;
		}
	}

	public class CellStatus
	{
		public string? Message { get; set; }
		public bool Success
		{
			get { return string.IsNullOrWhiteSpace(Message); }
		}
	}

	public class CellData
	{
		public CellStatus Status { get; set; }
		public Columns ColumnConfigurations { get; set; }
		public List<List<Cell>> Cells { get; set; }
		public string SheetName { get; set; }
		public bool ExtendedSheet { get; set; }

		public CellData() : this(false)
		{ }

		// extended sheet allows merged cells and enforce column and row index use
		public CellData(bool extendedSheet)
		{
			Status = new CellStatus();
			Cells = new List<List<Cell>>();
			ExtendedSheet = extendedSheet;
		}
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
			double number;

			if (double.TryParse(text, out number))
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

		private MergeCell CreateMergeCell(Cell excelCell)
		{
			MergeCell mergeCell = new MergeCell
			{
				Reference = new StringValue(ColumnLetter(excelCell.ColumnIndex)
						+ (excelCell.RowIndex + 1) + ":"
						+ ColumnLetter(excelCell.ColumnIndex + excelCell.ColSpan - 1) +
						+(excelCell.RowIndex + excelCell.RowSpan)),
			};
			return mergeCell;
		}

		public byte[] GenerateSpreadsheet<T>(List<Column<T>> columns, IEnumerable<T> items)
		{
			CellData excelData = new CellData();
			excelData.SheetName = "Items";
			var header = new List<Cell>();
			foreach (var column in columns)
			{
				if (!column.Hidden && !string.IsNullOrEmpty(column.PropertyName))
				{
					header.Add(new Cell(column.Title));
				}
			}
			excelData.Cells.Add(header);

			foreach (var item in items)
			{
				Type t = item.GetType();
				List<Cell> row = new List<Cell>();
				foreach (var column in columns)
				{
					if (!column.Hidden)
					{
						if (!string.IsNullOrEmpty(column.PropertyName))
						{
							PropertyInfo prop = t.GetProperty(column.PropertyName);
							object val = prop.GetValue(item);
							row.Add(new Cell(val != null ? val.ToString() : string.Empty));
						}
					}
				}
				excelData.Cells.Add(row);
			}
			return GenerateSpreadsheet(excelData);
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
				AppendChild<Sheets>(new Sheets());

			var sheet = new Sheet()
			{
				Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = 1,
				Name = data.SheetName ?? "Sheet 1"
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
			DocumentFormat.OpenXml.Spreadsheet.Row row;
			var cellIdex = 0;

			// Add sheet data
			foreach (var rowData in data.Cells)
			{
				cellIdex = 0;
				row = new DocumentFormat.OpenXml.Spreadsheet.Row { RowIndex = ++rowIdex };
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
			CellData data)
		{
			UInt32 rowIdex = 0;
			Row row;
			MergeCells mergeCells = new MergeCells();

			// Add sheet data
			foreach (var rowData in data.Cells)
			{
				row = new Row { RowIndex = ++rowIdex };
				sheetData.AppendChild(row);
				foreach (var excelCell in rowData)
				{
					var cell = CreateCell(ColumnLetter(excelCell.ColumnIndex),
						(uint)(excelCell.RowIndex + 1), excelCell.Content ?? string.Empty);
					row.AppendChild(cell);
					if (excelCell.ColSpan > 1 || excelCell.RowSpan > 1)
					{
						var mergeCell = CreateMergeCell(excelCell);
						mergeCells.Append(mergeCell);
					}
				}
			}
			worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
		}
	}
}
