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

namespace DisciplineWorkProgram.Models.Sections
{
    public class Section : HierarchicalCheckableElement //Section - направление
    {
        public string Name => SectionDictionary.ContainsKey("WaySection") ? SectionDictionary["WaySection"] : "";

        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Disciplines.Values;

        private readonly string _compListPath;
        private readonly string _competenciesMatrixPath;

        //Содержит значения Section. Не свойства, поскольку закладки находятся как словарь и проще
        //использовать Section как словарь
        public IDictionary<string, string> SectionDictionary { get; set; }
        public IDictionary<string, Discipline> Disciplines { get; private set; }
        public IDictionary<string, Competence> Competencies { get; set; }
        //Ключ - название дисциплины, значение - список кодов компетенций
        public IDictionary<string, List<string>> DisciplineCompetencies { get; set; }
        public static IDictionary<string, string> CompetenceClassifiers = new Dictionary<string, string>
        {
            ["УК"] = "Универсальные компетенции (УК)",
            ["ОПК"] = "Общепрофессиональные компетенции (ОПК)",
            ["ПК"] = "Профессиональные компетенции (ПК)"
        };

        public Section(string competenciesListPath, string competenciesMatrixPath)
        {
            _compListPath = competenciesListPath;
            _competenciesMatrixPath = competenciesMatrixPath;
            SectionDictionary = new Dictionary<string, string>();
            Competencies = new Dictionary<string, Competence>();
            DisciplineCompetencies = new Dictionary<string, List<string>>();
        }

        public void LoadDataFromPlan(string path)
        {
            var workbook = new XLWorkbook(path);
            LoadSection(workbook);
        }

        public void LoadDataFromPlan(Stream plan)
        {
            var workbook = new XLWorkbook(plan);
            LoadSection(workbook);
        }

        public void LoadCompetenciesData()
        {
            using (var doc = WordprocessingDocument.Open(_compListPath, false))
                LoadCompetencies(doc);

            using (var doc = WordprocessingDocument.Open(_competenciesMatrixPath, false))
                LoadCompetenciesMatrix(doc);
        }

        //Короче, обяз. часть и другие. Их по-идее надо отделять. Может, в коммент как доп. поле поместить
        //к дисциплине или типа того. Но это надо. Наверное.
        //Допустим, здесь все равно
        private void LoadCompetenciesMatrix(WordprocessingDocument document)
        {
            foreach (var table in GetTables(document))
            {
                var headers = GetHeaders(table).ToArray();    //Получить заголовки таблиц
                                                              //По строкам ориентир. Одну пропускаем, так как это заголовки
                foreach (var row in table.Descendants<TableRow>().Skip(1).ToArray())
                {
                    if (row.Descendants<TableCell>().Count() < 2) continue; //Если повторно некоторый заголовок

                    var cells = row.Descendants<TableCell>().ToArray();
                    //var disc = cells[0].Elements<Paragraph>().Single().InnerText; //название дисциплины в первой ячейке
                    var disc = cells[0].InnerText.TrimStart();

                    if (!DisciplineCompetencies.ContainsKey(disc))
                        DisciplineCompetencies[disc] = new List<string>();
                    //Если заголовок не код компетенции или ячейка пуста, то пропускаем
                    for (var i = 1; i < headers.Length; i++)
                    {
                        if (!RegexPatterns.Competence.IsMatch(headers[i]) ||
                            i - (headers.Length - cells.Length) < 0 ||
                            string.IsNullOrWhiteSpace(cells[i-(headers.Length - cells.Length)].InnerText))//string.IsNullOrWhiteSpace(cells[i].Elements<Paragraph>().Single().InnerText))
                            continue;

                        DisciplineCompetencies[disc].Add(headers[i]);
                    }
                }
            }
        }

        private void LoadSection(IXLWorkbook workbook)
        {
            //var regex = new Regex("(?<=\").*(?=\")");
            var worksheet = workbook.Worksheet("Титул");
            SectionDictionary["EducationLevel"] = worksheet.Cell(FindCell(worksheet, "квалификация", false)).Value.ToString().ToLower().Replace("квалификация:", "").Trim();
            switch (SectionDictionary["EducationLevel"])
            {
                case "бакалавр":
                    {
                        SectionDictionary["EducationLevel"] = "Бакалавриат";
                        break;
                    }
                case "магистр":
                    {
                        SectionDictionary["EducationLevel"] = "Магистратура";
                        break;
                    }
                case "аспирант":
                    {
                        SectionDictionary["EducationLevel"] = "Аспирантура";
                        break;
                    }
                case "специалист":
                    {
                        SectionDictionary["EducationLevel"] = "Специалитет";
                        break;
                    }

                default:
                    break;
            }
            SectionDictionary["WayCode"] = worksheet.Cell(FindCell(worksheet, "\\d\\d.\\d\\d.\\d\\d$", true)).Value.ToString();
            SectionDictionary["EducationForm"] = worksheet.Cell(FindCell(worksheet, "форма обучения")).Value.ToString().Replace("Форма обучения: ", "");

            if (SectionDictionary["EducationLevel"] == "Специалитет")
            {
                SectionDictionary["WayName"] = worksheet.Cell(FindCell(worksheet, "Специальность:", false)).Value.ToString().Replace("Специальность:", "").Trim();
                SectionDictionary["WaySection"] = worksheet.Cell(FindTwoCell(worksheet, "Специализация", false)[1]).Value.ToString();
            }
            else
            {
                //B18 - сложная строка, требуется разложение
                var matches = RegexPatterns.WayNameSection.Matches(worksheet.Cell(FindCell(worksheet, "направление подготовки")).Value.ToString());
                SectionDictionary["WayName"] = matches[0].Value;
                SectionDictionary["WaySection"] = matches[1].Value; //Профиль

            }
            Disciplines = DisciplineWorkProgram.Models.Helpers.GetDisciplines(workbook, this, SectionDictionary["EducationLevel"]);
            LoadDetailedDisciplineData(workbook);
        }



        private void LoadDetailedDisciplineData(IXLWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets.Where(sheet => sheet.Name.StartsWith("Курс")))
            {
                int i = 0;
                int cur = 0;
                int count = 1;
                foreach (var row in worksheet.RowsUsed().Where(row => int.TryParse(row.Cell(FindColumn(worksheet, "№")).GetString(), out _))
                    .Concat(worksheet.RowsUsed().Where(row =>
                        row.Cell(FindColumn(worksheet, "наименование")).GetString().ToLower().ContainsAny("практика", "аттестация"))))
                {
                    var discipline = row.Cell(FindColumn(worksheet, "Индекс", true)).GetString();
                    if (string.IsNullOrWhiteSpace(discipline))
                        discipline = row.Cell("E").GetString(); //вроде не актуально

                    if (!Disciplines.ContainsKey(discipline)) continue;
                    //Изменить на трайпарс после дебага

                    string[] semestrs = FindTwoCell(worksheet, "семестр");
                    var semester = 0;
                    bool isGood = int.TryParse(RegexPatterns.DigitInString.Match(worksheet.Cell(semestrs[0]).GetString()).Value, out semester);
                    if (!isGood) semester = ((cur + 1)) * 2;
                    string[] academChas = FindTwoCell(worksheet, "Академических");


                    int.TryParse(row.Cell(FindColumn(worksheet, "№")).GetString(), out cur);
                    if (cur < i)
                    {
                        count++;
                    }
                    i = cur;
                    if (count < (semester + 1) / 2)
                    {
                        continue;
                    }

                    var details = new DisciplineDetails
                    {
                        Monitoring = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetString(),
                        Contact = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[А|а]\\s*[К|к]\\s*[Т|т]", true)).GetInt(),
                        Lec = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^лек$", true)).GetInt(),
                        Lab = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^лаб$", true)).GetInt(),
                        Pr = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^пр$", true)).GetInt(),
                        Ind = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^ср$", true)).GetInt(),
                        Control = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                        Ze = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[0]), @"^з\.е\.$", true)).GetInt()
                    };

                    if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
                        Disciplines[discipline].Details.Add(semester, details);
                    //to do: заменить на поиск по странице.
                    isGood = int.TryParse(RegexPatterns.DigitInString.Match(worksheet.Cell(semestrs[0]).GetString()).Value, out semester);
                    if (!isGood) semester = ((cur + 1)) * 2 + 1;

                    details = new DisciplineDetails
                    {
                        Monitoring = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetString(),
                        Contact = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[А|а]\\s*[К|к]\\s*[Т|т]", true)).GetInt(),
                        Lec = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^лек$", true)).GetInt(),
                        Lab = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^лаб$", true)).GetInt(),
                        Pr = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^пр$", true)).GetInt(),
                        Ind = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^ср$", true)).GetInt(),
                        Control = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                        Ze = row.Cell(FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[1]), @"^з\.е\.$", true)).GetInt()
                    };

                    if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
                        Disciplines[discipline].Details.Add(semester, details);
                }
            }
        }

        private void LoadCompetencies(WordprocessingDocument document)
        {
            var competencies = ParseCompetencies(document).ToArray();
            var regex = RegexPatterns.CompetenceName2;
            //var regex = new Regex(@"^(УК-[\dЗ]+(\.\d+)*|ОПК-[\dЗ]+(\.\d+)*|ПК-[\dЗ]+(\.[\dЗ]+)*)\b");
            //Составление набора ключей-компетенций 
            foreach (var competency in competencies)
            {
                var match = regex.Match(competency);
                if (match.Success)
                {
                    var key = match.Value.Replace(" ", "").Replace("З", "3");


                    if (!Competencies.ContainsKey(key))
                    {
                        // Если ключа ещё нет, создаём новую компетенцию
                        Competencies[key] = new Competence { Name = competency };
                    }
                    else
                    {
                        // Если ключ уже есть, добавляем строку в список компетенций
                        Competencies[key].Competencies.Add(competency);
                    }
                }
            }


        }

        public IEnumerable<string> GetCheckedDisciplinesNames =>
            Disciplines
                .Where(d => d.Value.IsChecked)
                .Select(kv => kv.Key);

        public IEnumerable<string> GetAnyDisciplinesNames =>
            Disciplines
                .Where(d => true)
                .Select(kv => kv.Key);

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
                            return cell.Address.ToString();
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
                            return cell.Address.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        private static string[] FindTwoCell(IXLWorksheet worksheet, string target, bool isRegex = false)
        {
            string[] answ = new string[2];
            int count = 0;
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target, RegexOptions.IgnoreCase))
                        {
                            answ[count++] = cell.Address.ToString();
                            if (count == 2) { return answ; }

                        }
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе в количестве 2-ух штук.");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                        {
                            answ[count++] = cell.Address.ToString();
                            if (count == 2) { return answ; }
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        private static string FindCell(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
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
                                    return cell.Address.ToString();
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
                                    return cellValue.Address.ToString();
                                }
                            }
                        }
                    }
                }
            }
            throw new Exception($"Нет поля {target1} в документе {worksheet.Name}");
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

        private static string FindColumnAnderCell(IXLWorksheet worksheet, IXLCell cell, string target, bool isRegex = false)
        {
            var mergedRange = cell.MergedRange() ?? cell.AsRange();
            var firstColumn = mergedRange.FirstCell().Address.ColumnLetter;
            var lastColumn = mergedRange.LastCell().Address.ColumnLetter;
            int startRow = mergedRange.LastCell().Address.RowNumber + 1;
            int endRow = worksheet.LastRowUsed().RowNumber();
            var searchRange = worksheet.Range($"{firstColumn}{startRow}:{lastColumn}{endRow}");

            if (isRegex)
            {
                foreach (var cellValue in searchRange.CellsUsed())
                {
                    string cellForReg = cellValue.GetValue<string>();
                    if (Regex.IsMatch(cellForReg, target, RegexOptions.IgnoreCase))
                    {
                        return cellValue.Address.ColumnLetter.ToString();
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе");
            }
            else
            {
                foreach (var cellValue in searchRange.CellsUsed())
                {
                    if (cellValue.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                    {
                        return cellValue.Address.ColumnLetter.ToString();
                    }
                }
            }
            throw new Exception($"Нет поля {target} в документе {worksheet.Name}");
        }
    }
}