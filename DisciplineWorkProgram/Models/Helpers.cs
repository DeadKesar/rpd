using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DisciplineWorkProgram.Models.Sections.Helpers;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models
{
	public static class Helpers
	{
		private const string WorksheetName = "План";

		public static IDictionary<string, Discipline> GetDisciplines(IXLWorkbook workbook, HierarchicalCheckableElement section)
		{
            

            var worksheet = workbook.Worksheet(WorksheetName);
            var add = FindCell(worksheet, "закрепленная кафедра");
            var disciplines = ExcelHelpers.GetRowsWithPlus(worksheet)
				.Select(row => new Discipline
				{
					Name = row.Cell(FindCell(worksheet, "наименование")).GetString(),
					Department = row.Cell(FindCell(worksheet, "закрепленная кафедра")).GetString(),
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
        /// <summary>
        /// Поиск заданного слова на странице
        /// </summary>
        /// <param name="worksheet">страница для поиска</param>
        /// <param name="target">слово которое ищем</param>
        /// <param name="isRegex">true если хотим передать регекс, иначе false</param>
        /// <returns>адресс ячейки где нашли слово(первый встреченный)</returns>
        /// <exception cref="Exception">нет искомого поля</exception>
        
		private static string FindCell(IXLWorksheet worksheet, string target, bool isRegex = false)
        {
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе {worksheet.Name}");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе {worksheet.Name}");
            }
        }

        private static string FindCell(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
           
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target1, StringComparison.OrdinalIgnoreCase))
                        {
                            var range = worksheet.Range(cell.CurrentRegion.ToString());
                        foreach (var cellValue in range.CellsUsed())
                        {
                            if (cellValue.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                            {
                                return cellValue.Address.ColumnLetter.ToString();
                            }
                        }
                        }
                    }
                }
                throw new Exception($"Нет поля {target1} в документе {worksheet.Name}");
        }
    }
}
