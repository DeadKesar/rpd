using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using static DisciplineWorkProgram.Models.Sections.Helpers.Competencies;

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
					var disc = cells[0].Elements<Paragraph>().Single().InnerText; //название дисциплины в первой ячейке

					if (!DisciplineCompetencies.ContainsKey(disc))
						DisciplineCompetencies[disc] = new List<string>();
					//Если заголовок не код компетенции или ячейка пуста, то пропускаем
					for (var i = 1; i < cells.Length; i++)
					{
						if (!RegexPatterns.Competence.IsMatch(headers[i]) ||
							string.IsNullOrWhiteSpace(cells[i].Elements<Paragraph>().Single().InnerText))
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
			try
			{
				SectionDictionary["EducationLevel"] = worksheet.Cell("G15").Value.ToString().Replace("по программе", "").Trim();
				SectionDictionary["WayCode"] = worksheet.Cell("C17").Value.ToString();
				//B18 - сложная строка, требуется разложение
				var matches = RegexPatterns.WayNameSection.Matches(worksheet.Cell("C19").Value.ToString());
				SectionDictionary["WayName"] = matches[0].Value;
				SectionDictionary["WaySection"] = matches[1].Value; //Профиль
				SectionDictionary["EducationForm"] = worksheet.Cell("B32").Value.ToString().Replace("Форма обучения: ", "");
			}
			catch 
			{
				//костыль над здравым смыслом
                SectionDictionary["EducationLevel"] = worksheet.Cell("F14").Value.ToString().Replace("по программе", "").Trim();
                SectionDictionary["WayCode"] = worksheet.Cell("B16").Value.ToString();
                //B18 - сложная строка, требуется разложение
                var matches = RegexPatterns.WayNameSection.Matches(worksheet.Cell("B18").Value.ToString());
                SectionDictionary["WayName"] = matches[0].Value;
                SectionDictionary["WaySection"] = matches[1].Value; //Профиль
                SectionDictionary["EducationForm"] = worksheet.Cell("A31").Value.ToString().Replace("Форма обучения: ", "");
            }
			Disciplines = DisciplineWorkProgram.Models.Helpers.GetDisciplines(workbook, this);
			LoadDetailedDisciplineData(workbook);
		}

		private void LoadDetailedDisciplineData(IXLWorkbook workbook)
		{
			foreach (var worksheet in workbook.Worksheets.Where(sheet => sheet.Name.StartsWith("Курс")))
			{
                //Where(row => int.TryParse(row.Cell("C").GetString(), out _)) фильтрует строки, оставляя только те, у которых значение
                //в ячейке столбца "C" может быть успешно преобразовано в целое число. Это достигается с помощью
                //метода int.TryParse(), который возвращает true, если преобразование удалось, и false в противном случае.
                foreach (var row in worksheet.RowsUsed().Where(row => int.TryParse(row.Cell("C").GetString(), out _))
					.Concat(worksheet.RowsUsed().Where(row =>
						row.Cell("E").GetString().ToLower().ContainsAny("практика", "аттестация", "квалификационной"))))
				{
					var discipline = row.Cell("F").GetString();
					if (string.IsNullOrWhiteSpace(discipline))
						discipline = row.Cell("E").GetString();

					if (!Disciplines.ContainsKey(discipline)) continue;
					//Изменить на трайпарс после дебага
					//это какойто объект описывающий семестры походу
					//Пусть это будет первый семестр
					var semester =
						int.Parse(RegexPatterns.DigitInString.Match(worksheet.Cell(3, "G").GetString()).Value);

					var details = new DisciplineDetails
					{
						//предположим:
						Monitoring = row.Cell("G").GetString(), //контроль
                        Contact = row.Cell("I").GetInt(), //Контакт
						Lec = row.Cell("J").GetInt(), //лекции
						Lab = row.Cell("N").GetInt(), //лабораторные
						Pr = row.Cell("R").GetInt(), //Пр.
						Ind = row.Cell("V").GetInt(), //хз... пусть будет СР, самостоятельные работы, т.е. индивидуальные...
						Control = row.Cell("Z").GetInt(), //опять контроль
						Ze = row.Cell("AD").GetInt() //зачётные единицы
					};

					if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
						Disciplines[discipline].Details.Add(semester, details);

					
					semester = int.Parse(RegexPatterns.DigitInString.Match(worksheet.Cell(3, "AF").GetString()).Value);
					//а это второй
					details = new DisciplineDetails
					{
						Monitoring = row.Cell("AF").GetString(),
						Contact = row.Cell("AH").GetInt(),
						Lec = row.Cell("AI").GetInt(),
						Lab = row.Cell("AM").GetInt(),
						Pr = row.Cell("AQ").GetInt(),
						Ind = row.Cell("AU").GetInt(),
						Control = row.Cell("AY").GetInt(),
						Ze = row.Cell("BC").GetInt()
					};

					if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
						Disciplines[discipline].Details.Add(semester, details);
				}
			}
		}

		private void LoadCompetencies(WordprocessingDocument document)
		{
			var competencies = ParseCompetencies(document).ToArray();
			//Составление набора ключей-компетенций
			foreach (var competency in competencies.Where(text => RegexPatterns.CompetenceName.IsMatch(text)))
				Competencies[competency.Substring(0, competency.IndexOf('.')).Replace(" ", "")] =
					new Competence { Name = competency };

			foreach (var competency in competencies.Where(text => !RegexPatterns.CompetenceName.IsMatch(text)))
				Competencies[competency.Substring(0, competency.IndexOf('.')).Replace(" ", "")]
					.Competencies.Add(competency);
		}

		public IEnumerable<string> GetCheckedDisciplinesNames =>
			Disciplines
				.Where(d => d.Value.IsChecked)
				.Select(kv => kv.Key);
	}
}
