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
            var disciplines = ExcelHelpers.GetRowsWithPlus(worksheet)
				.Select(row => new Discipline
				{
                    Ind = row.Cell(FindCell(worksheet, "индекс")).GetString(),
                    Name = row.Cell(FindCell(worksheet, "наименование")).GetString(), //C
					Department = row.Cell(FindCell(worksheet, "закрепленная кафедра", "наименование")).GetString(),
					Exam = row.Cell(FindCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[З|з]\\s*[А|а]\\s*[М|м]\\s*[Е|е]\\s*[Н|н]", true)).GetInt(),
					Credit = row.Cell(FindCell(worksheet, "зачет")).GetInt(),
					CreditWithRating = row.Cell(FindCell(worksheet, "зачет с оц")).GetInt(),
					Kp = row.Cell(FindCell(worksheet, "^кп$",true)).GetInt(),
					Kr = row.Cell(FindCell(worksheet, "^кр$", true)).GetInt(),
					Fact = row.Cell(FindCell(worksheet, "факт")).GetInt(),
                    ByPlan = row.Cell(FindCellOr(worksheet, "[П|п]?\\s*[О|о]\\s*[П|п]\\s*[Л|л]\\s*[А|а]\\s*[Н|н]s*[У|у]", "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true)).GetInt(), //экспертное
                    ContactHours = row.Cell(FindCell(worksheet, "Конт. раб.")).GetInt(),
					Lec = row.Cell(FindCell(worksheet, "Лаб")).GetInt(),
					Lab = row.Cell(FindCell(worksheet, "^пр$", true)).GetInt(),
					Pr = row.Cell(FindCell(worksheet, "^ср$", true)).GetInt(),
					
					Control = row.Cell(FindCell(worksheet, "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
					ZeAtAll = row.Cells(FindCell(worksheet, "Семестр 1"), FindCell(worksheet, "Семестр 8")).Sum(val => val.GetInt()),

					Parent = section
				})
                    .Aggregate(new Dictionary<string, Discipline>(), (dict, discipline) =>
                    {
                        string originalName = discipline.Ind;
                        string nameToUse = originalName;
                        int counter = 2;

                        // Пока ключ уже существует, добавляем суффикс
                        while (dict.ContainsKey(nameToUse))
                        {
                            nameToUse = $"{originalName}{counter}";
                            counter++;
                        }

                        // Добавляем дисциплину с уникальным именем
                        dict[nameToUse] = discipline;

                        return dict;
                    });

            if (!disciplines.Any())
			{
				disciplines = ExcelHelpers.GetRowsWithPlus(workbook.Worksheet(WorksheetName + "Свод"))
					.Select(row => new Discipline
					{
                        Name = row.Cell(FindCell(worksheet, "наименование")).GetString(), //C
                        Department = row.Cell(FindCell(worksheet, "закрепленная кафедра", "наименование")).GetString(),
                        Exam = row.Cell(FindCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[З|з]\\s*[А|а]\\s*[М|м]\\s*[Е|е]\\s*[Н|н]", true)).GetInt(),
                        Credit = row.Cell(FindCell(worksheet, "зачет")).GetInt(),
                        CreditWithRating = row.Cell(FindCell(worksheet, "зачет с оц")).GetInt(),
                        Kp = row.Cell(FindCell(worksheet, "^кп$", true)).GetInt(),
                        Kr = row.Cell(FindCell(worksheet, "^кр$", true)).GetInt(),
                        Fact = row.Cell(FindCell(worksheet, "факт")).GetInt(),
                        ByPlan = row.Cell(FindCellOr(worksheet, "[П|п]?\\s*[О|о]\\s*[П|п]\\s*[Л|л]\\s*[А|а]\\s*[Н|н]s*[У|у]", "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true)).GetInt(), //экспертное
                        ContactHours = row.Cell(FindCell(worksheet, "Конт. раб.")).GetInt(),
                        Lec = row.Cell(FindCell(worksheet, "Лаб")).GetInt(),
                        Lab = row.Cell(FindCell(worksheet, "^пр$", true)).GetInt(),
                        Pr = row.Cell(FindCell(worksheet, "^ср$", true)).GetInt(),
                        Ind = row.Cell(FindCell(worksheet, "индекс")).GetText(),
                        Control = row.Cell(FindCell(worksheet, "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                        ZeAtAll = row.Cells(FindCell(worksheet, "Семестр 1"), FindCell(worksheet, "Семестр 8")).Sum(val => val.GetInt()),

                        Parent = section
					})
                        .Aggregate(new Dictionary<string, Discipline>(), (dict, discipline) =>
                        {
                            string originalName = discipline.Ind;
                            string nameToUse = originalName;
                            int counter = 2;

                            // Пока ключ уже существует, добавляем суффикс
                            while (dict.ContainsKey(nameToUse))
                            {
                                nameToUse = $"{originalName}{counter}";
                                counter++;
                            }

                            // Добавляем дисциплину с уникальным именем
                            dict[nameToUse] = discipline;

                            return dict;
                        });
            }
            //Возможно надо отфильтровать дисциплины. и убрать заголовки


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
                        var mergedRange = cell.MergedRange() ?? cell.AsRange();
                        var firstColumn = mergedRange.FirstCell().Address.ColumnLetter;
                        var lastColumn = mergedRange.LastCell().Address.ColumnLetter;
                        int startRow = mergedRange.LastCell().Address.RowNumber + 1;
                        int endRow = worksheet.LastRowUsed().RowNumber();
                        var searchRange = worksheet.Range($"{firstColumn}{startRow}:{lastColumn}{endRow}");


                        //var range = worksheet.Range($"План!{cell.Address.ToString()}:{cell.CurrentRegion.ToString().Split(':')[1]}");
                        foreach (var cellValue in searchRange.CellsUsed())
                        {
                            if (cellValue.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                            {
                                return cellValue.Address.ColumnLetter.ToString();
                            }
                        }
                        }
                    }
                }
                throw new Exception($"Нет поля {target1}, или {target2} в документе {worksheet.Name}");
        }

        private static string FindCellOr(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target1, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target2, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target1}, или {target2} в документе {worksheet.Name}");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target1, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target1}, или {target2} в документе {worksheet.Name}");
            }
        }
    }
}
