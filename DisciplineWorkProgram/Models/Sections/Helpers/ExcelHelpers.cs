using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace DisciplineWorkProgram.Models.Sections.Helpers
{
	public static class ExcelHelpers
	{
		public static IEnumerable<IXLRow> GetRowsWithPlus(IXLWorksheet worksheet) =>
			worksheet.RowsUsed().Where(row => row.Cell("A").GetString().Equals("+") && !row.Cell("A").Style.Font.Bold);

		public static IEnumerable<IXLRow> GetRowsWithPractices(IXLWorksheet worksheet) =>
			worksheet.RowsUsed().Where(row => row.Cell("D").GetString().ToLower().Contains("практика"));
	}
}
