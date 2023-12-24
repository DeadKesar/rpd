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
					Name = row.Cell("B").GetString(), //название дисцыплины
					Department = row.Cell("AA").GetString(), //закреплённая кафедра
					Exam = row.Cell("C").GetInt(), //экзамен
					Credit = row.Cell("D").GetInt(), //Зачёт
					CreditWithRating = row.Cell("E").GetInt(), //Зачёт с оценкой
                    Kp = row.Cell("F").GetInt(), //КП
					Kr = row.Cell("G").GetInt(), //КР
					Fact = row.Cell("H").GetInt(), //фактическое число З.Е.
					ByPlan = row.Cell("K").GetInt(), //По плану
					ContactHours = row.Cell("L").GetInt(), //Конт. раб.
                    Lec = row.Cell("M").GetInt(), //лекции
					Lab = row.Cell("N").GetInt(), //лабораторные 
					Pr = row.Cell("O").GetInt(), //практические
					Ind = row.Cell("P").GetInt(),//самостоятелтная работа
					Control = row.Cell("Q").GetInt(), //контроль
					ZeAtAll = row.Cells("R", "Y").Sum(val => val.GetInt()), //сумируем зачётные единицы

					Parent = section
				})
				.ToDictionary(discipline => discipline.Name);

			if (!disciplines.Any())
			{
				disciplines = ExcelHelpers.GetRowsWithPlus(workbook.Worksheet(WorksheetName + "Свод"))
					.Select(row => new Discipline
					{
                        Name = row.Cell("C").GetString(), //название дисцыплины
                        Department = row.Cell("AA").GetString(), //закреплённая кафедра
                        Exam = row.Cell("D").GetInt(), //экзамен
                        Credit = row.Cell("E").GetInt(), //Зачёт
                        CreditWithRating = row.Cell("F").GetInt(), //Зачёт с оценкой
                        Kp = row.Cell("G").GetInt(), //КП
                        Kr = row.Cell("H").GetInt(), //КР
                        Fact = row.Cell("J").GetInt(), //фактическое число З.Е.
                        ByPlan = row.Cell("L").GetInt(), //По плану
                        ContactHours = row.Cell("K").GetInt(), //Конт. раб.
                        Lec = row.Cell("N").GetInt(), //лекции
                        Lab = row.Cell("N").GetInt(), //лабораторные 
                        Pr = row.Cell("N").GetInt(), //практические
                        Ind = row.Cell("O").GetInt(),//самостоятелтная работа
                        Control = row.Cell("P").GetInt(), //контроль
                        ZeAtAll = row.Cells("R", "Y").Sum(val => val.GetInt()), //сумируем зачётные единицы

                        Parent = section
					})
					.ToDictionary(discipline => discipline.Name);
			}

			return disciplines;
		}
	}
}
