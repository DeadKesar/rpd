using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DisciplineWorkProgram.Excel
{
	public class Converter
	{
		private const string ConvertedFilesPath = "Converted/";
		private const string Extension = ".xlsx";

		public static void Convert(string path, out string newPath)
		{
			newPath = GetConvertedFilePath(path);

			if (!Directory.Exists(ConvertedFilesPath)) Directory.CreateDirectory(ConvertedFilesPath);

			ConvertXlsToXlsx(path).Write(File.Create(newPath));
		}

		public static Stream Convert(string path)
		{
			var newPlan = new NpoiMemoryStream() { AllowClose = false };

			ConvertXlsToXlsx(path).Write(newPlan);
			newPlan.AllowClose = true;

			return newPlan;
		}

		private static XSSFWorkbook ConvertXlsToXlsx(string path)
		{
			using var inputStream = File.OpenRead(path);
			// var workbookIn = new HSSFWorkbook(inputStream);
			var workbookIn = new XSSFWorkbook(inputStream);
			var workbookOut = new XSSFWorkbook();

			for (var i = 0; i < workbookIn.NumberOfSheets; i++)
			{
				var sheetIn = workbookIn.GetSheetAt(i);
				var sheetOut = workbookOut.CreateSheet(sheetIn.SheetName);
				var rowEnumerator = sheetIn.GetEnumerator();
				while (rowEnumerator.MoveNext())
				{
					// var rowIn = (HSSFRow)rowEnumerator.Current;
					var rowIn = (XSSFRow)rowEnumerator.Current;
					// if (rowIn == null || rowIn.IsHidden) continue;
					if (rowIn == null) continue;
					var rowOut = sheetOut.CreateRow(rowIn.RowNum);
					CopyRowProperties(rowOut, rowIn);
				}
			}
			return workbookOut;
		}

		private static void CopyRowProperties(IRow rowOut, IRow rowIn)
		{
			rowOut.RowNum = rowIn.RowNum;

			using var cellEnumerator = rowIn.GetEnumerator();

			while (cellEnumerator.MoveNext())
			{
				var cellIn = cellEnumerator.Current;
				if (cellIn == null) continue;
				var cellOut = rowOut.CreateCell(cellIn.ColumnIndex, cellIn.CellType);
				CopyCellProperties(cellOut, cellIn);
			}
		}

		private static void CopyCellProperties(ICell cellOut, ICell cellIn)
		{
			switch (cellIn.CellType)
			{
				case CellType.Blank:
					break;
				case CellType.Boolean:
					cellOut.SetCellValue(cellIn.BooleanCellValue);
					break;
				case CellType.Error:
					cellOut.SetCellValue(cellIn.ErrorCellValue);
					break;
				case CellType.Formula:
					cellOut.SetCellFormula(cellIn.CellFormula);
					break;
				case CellType.Numeric:
					cellOut.SetCellValue(cellIn.NumericCellValue);
					break;
				case CellType.String:
					cellOut.SetCellValue(cellIn.StringCellValue);
					break;
				case CellType.Unknown:
					break;
				default:
					return;
			}
		}

		private static string GetConvertedFilePath(string path) =>
			Path.GetFullPath(ConvertedFilesPath + Path.GetFileNameWithoutExtension(path) + Extension);
	}
}
