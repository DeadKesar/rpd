using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using static DisciplineWorkProgram.Models.Sections.Helpers.Competencies;
using System;
using System.Reactive.Joins;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json.Serialization;

namespace DisciplineWorkProgram.Models
{
    public class Employee : HierarchicalCheckableElement
    {
        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Enumerable.Empty<HierarchicalCheckableElement>();

        [JsonIgnore]
        public IDictionary<string, IDictionary<string, string>> Employees { get; } = new Dictionary<string, IDictionary<string, string>>();

        public Employee(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Путь к файлу не может быть пустым.", nameof(path));

            if (!File.Exists(path))
                throw new FileNotFoundException("Файл не найден.", path);

            // Открываем Excel-файл
            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheets
                .SingleOrDefault(sheet => sheet.Name.StartsWith("Сотрудники"));

            if (worksheet == null)
                throw new InvalidOperationException("Не найден лист, начинающийся с 'Сотрудники'.");

            // Проходим по всем строкам с данными
            foreach (var row in worksheet.RowsUsed().Where(row => int.TryParse(row.Cell(FindColumn(worksheet, "номер")).GetString(), out _)))
            {
                var emp = row.Cell("B").GetString();
                var employeeData = new Dictionary<string, string>
                {
                    ["nameForDoc"] = row.Cell("C").GetString(),
                    ["position"] = row.Cell("D").GetString(),
                    ["FIO"] = row.Cell("E").GetString(),
                    ["institut"] = row.Cell("F").GetString(),
                };

                // Добавляем сотрудника в словарь
                Employees[emp] = employeeData;
            }
        }

       
        
        
        private static string FindColumn(IXLWorksheet worksheet, string target, bool isRegex = false)
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
                throw new Exception($"Нет ПАТЕРНА {target} в документе");
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
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        
        
        
        private static string FindColumn(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            if (isRegex)
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
                                string cellForReg = cellValue.GetValue<string>();
                                if (Regex.IsMatch(cellForReg, target2, RegexOptions.IgnoreCase))
                                {
                                    return cellValue.Address.ColumnLetter.ToString();
                                }
                            }
                            throw new Exception($"Нет ПАТЕРНА {target2} в документе");
                        }
                    }
                }
            }

            else
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
            }
            throw new Exception($"Нет поля {target1} в документе {worksheet.Name}");
        }
    }
}
