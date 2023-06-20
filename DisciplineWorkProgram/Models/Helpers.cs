using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DisciplineWorkProgram.Models.Sections.Helpers;
using System.Collections.Generic;
using System.Linq;

namespace DisciplineWorkProgram.Models
{
	public static class Helpers
	{
		private const string WorksheetName = "План";

		public static IDictionary<string, Discipline> GetDisciplines(IXLWorkbook workbook, HierarchicalCheckableElement section)
		{
			var disciplines = ExcelHelpers.GetRowsWithPlus(workbook.Worksheet(WorksheetName))
				.Select(row => new Discipline
				{
					Name = row.Cell("C").GetString(),
					Department = row.Cell("AB").GetString(),
					Exam = row.Cell("D").GetInt(),
					Credit = row.Cell("E").GetInt(),
					CreditWithRating = row.Cell("F").GetInt(),
					Kp = row.Cell("G").GetInt(),
					Kr = row.Cell("H").GetInt(),
					Fact = row.Cell("I").GetInt(),
					ByPlan = row.Cell("K").GetInt(),
					ContactHours = row.Cell("L").GetInt(),
					Lec = row.Cell("N").GetInt(),
					Lab = row.Cell("O").GetInt(),
					Pr = row.Cell("P").GetInt(),
					Ind = row.Cell("Q").GetInt(),
					Control = row.Cell("R").GetInt(),
					ZeAtAll = row.Cells("S", "Z").Sum(val => val.GetInt()),

					Parent = section
				})
				.ToDictionary(discipline => discipline.Name);

			if (!disciplines.Any())
			{
				disciplines = ExcelHelpers.GetRowsWithPlus(workbook.Worksheet(WorksheetName + "Свод"))
					.Select(row => new Discipline
					{
						Name = row.Cell("C").GetString(),
						Department = row.Cell("AB").GetString(),
						Exam = row.Cell("D").GetInt(),
						Credit = row.Cell("E").GetInt(),
						CreditWithRating = row.Cell("F").GetInt(),
						Kp = row.Cell("G").GetInt(),
						Kr = row.Cell("H").GetInt(),
						Fact = row.Cell("I").GetInt(),
						ByPlan = row.Cell("K").GetInt(),
						ContactHours = row.Cell("L").GetInt(),
						Lec = row.Cell("N").GetInt(),
						Lab = row.Cell("O").GetInt(),
						Pr = row.Cell("P").GetInt(),
						Ind = row.Cell("Q").GetInt(),
						Control = row.Cell("R").GetInt(),
						ZeAtAll = row.Cells("S", "Z").Sum(val => val.GetInt()),

						Parent = section
					})
					.ToDictionary(discipline => discipline.Name);
			}

			return disciplines;
		}
	}
}
