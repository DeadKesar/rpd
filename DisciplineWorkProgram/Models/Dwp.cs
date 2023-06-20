using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using static DisciplineWorkProgram.Word.Helpers.Ooxml;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using DisciplineWorkProgram.Models.Sections;
using DocumentFormat.OpenXml;

namespace DisciplineWorkProgram.Models
{
	public class Dwp
	{
		public Dwp(Section section)
		{
			Section = section;
		}

		private Section Section { get; }

		//Должно обрабатывать только 1 дисциплину, чтобы "масштабировать" без доп. кода
		public void MakeDwp(string templatePath, string dwpDir, string discipline)
		{
			using var doc = WordprocessingDocument.CreateFromTemplate(templatePath, true);
			var bookmarkMap = GetBookmarks(doc, "Autofill");

			WriteSectionData(bookmarkMap);
			WriteDisciplineData(bookmarkMap, discipline);
			WriteCompetenciesTable(bookmarkMap, discipline);
			WriteDisciplinePartitionTable(bookmarkMap, discipline);
			WriteSemesters(bookmarkMap, discipline);
			WriteCompetencies(bookmarkMap, discipline);
			WriteYear(bookmarkMap);
			// Не реализовано занесение данных по дисциплине
			WriteLaboriousnessTable(bookmarkMap, discipline);
			SaveDoc(doc, dwpDir, discipline);
			doc.Dispose();
		}

		private void WriteSectionData(IDictionary<string, BookmarkStart> bookmarkMap)
		{
			foreach (var (key, bookmark) in bookmarkMap)
			{
				var actualKey = key.Substring(0, key.Length - 1);

				try
				{
					if (Section.SectionDictionary.ContainsKey(actualKey))
						FindElementsByBookmark<Text>(bookmark, 1)
							.First(elem => elem.Text.Contains("Autofill" + actualKey))
							.Text = Section.SectionDictionary[actualKey];
				}
				catch
				{
					Console.WriteLine(key);
					Environment.Exit(1);
				}
			}
		}

		private void WriteDisciplineData(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline))
				return;

			foreach (var (key, bookmark) in bookmarkMap)
			{
				//Надежда на то, что понадобится нумерация в рамках только одноразрядного числа
				var actualKey = key.Substring(0, key.Length - 1);

				if (Section.Disciplines[discipline].Props.ContainsKey(actualKey))
					FindElementsByBookmark<Text>(bookmark, 1)
						.First(elem => elem.Text.Contains("Autofill" + actualKey))
						.Text = Section.Disciplines[discipline].Props[actualKey];
			}
		}

		private void WriteCompetenciesTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(discipline))
				return;

			var table = FindElementsByBookmark<Table>(bookmarkMap["CompetenciesTable1"], 2).First();

			foreach (var competence in Section.DisciplineCompetencies[discipline])
			{
				if (!Section.Competencies.ContainsKey(competence)) continue;

				var row = new TableRow();
				row.AppendChild(GetTableCellByString(Section.Competencies[competence].Name));
				row.AppendChild(GetTableCellByStrings(Section.Competencies[competence].Competencies));
				table.AppendChild(row);
			}
		}

		private void WriteDisciplinePartitionTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(discipline))
				return;

			var rows = new List<TableRow>();

			foreach (var (semester, _) in Section.Disciplines[discipline].Details)
			{
				var row = new TableRow();
				row.Append(
					GetTableCellByString(""),
					GetTableCellByString(""),
					new TableCell(
						new Paragraph(
							new Run(
								new Text(semester.ToString())))
						{
							ParagraphProperties = new ParagraphProperties(
								new Justification { Val = JustificationValues.Center })
						}
					));
				row.Append(
					GetTableCellsByStrings("",
					"",
					"",
					"",
					"",
					"",
					""));

				rows.Add(row);
			}

			rows.Add(new TableRow(
				GetTableCellsByStrings(
					"",
					"Итого",
					"",
					"",
					Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Lec).ToString(),
					Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Lab).ToString(),
					Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Pr).ToString(),
					Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Ind).ToString(),
					"",
					"")));
			rows.Add(new TableRow(
				GetTableCellsByStrings(
					"",
					"Промежуточный контроль",
					"",
					"",
					"",
					"",
					"",
					"",
					"",
					Section.Disciplines[discipline].Details
						.FirstOrDefault(elem =>
							elem.Key == Section.Disciplines[discipline].Details.Keys.Max())
						.Value.Monitoring)));

			FindElementsByBookmark<Table>(bookmarkMap["DisciplinePartitionTable1"], 2)
				.First()
				.Append(rows);
		}

		private void WriteSemesters(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(discipline)) return;

			FindElementsByBookmark<Text>(bookmarkMap["Semester1"], 1)
					.First(elem => elem.Text.Contains("Autofill" + "Semester"))
					.Text =
				Section.Disciplines[discipline].Details.Keys
					.Select(elem => elem.ToString())
					.Aggregate((curr, next) => curr + ", " + next);
		}

		private void WriteCompetencies(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(discipline))
				return;

			var bookmarkElement = FindElementsByBookmark<Paragraph>(bookmarkMap["Competencies1"], 1).First();
			var currentElement = bookmarkElement;

			foreach (var (name, classifier) in Section.CompetenceClassifiers)
			{
				if (!Section.DisciplineCompetencies[discipline].Any(competence => competence.StartsWith(name)))
					continue;

				var element = new Paragraph
				{
					ParagraphProperties = new ParagraphProperties
					{
						ParagraphStyleId = new ParagraphStyleId { Val = "Default" },
						Indentation = new Indentation { FirstLine = "709" },
						Justification = new Justification { Val = JustificationValues.Both }
					}
				};
				currentElement.InsertAfterSelf(element);
				currentElement = element;

				element.AppendChild(
					new Run(new Text(classifier + ":"))
					{
						RunProperties = new RunProperties { Bold = new Bold() }
					});

				var classifiedCompetencies = Section.Competencies
					.Where(elem => Section.DisciplineCompetencies[discipline].Contains(elem.Key));

				Text text = default;

				foreach (var (_, competence) in classifiedCompetencies.Where(elem => elem.Key.StartsWith(name)))
				{
					element = new Paragraph
					{
						ParagraphProperties = new ParagraphProperties
						{
							SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
							Indentation = new Indentation { FirstLine = "720" },
							Justification = new Justification { Val = JustificationValues.Both }
						}
					};
					currentElement.InsertAfterSelf(element);
					currentElement = element;

					var dot = competence.Name.IndexOf('.');
					var rightText = competence.Name.Substring(0, dot + 1);

					text = new Text(rightText) { Space = SpaceProcessingModeValues.Preserve };
					element.AppendChild(
						new Run(text)
						{
							RunProperties = new RunProperties
							{
								Bold = new Bold(),
								RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
								FontSize = new FontSize { Val = "24" },
							}
						});

					rightText = competence.Name.Substring(dot + 1);
					if (rightText[^1] == '.' || rightText[^1] == ';')
						rightText = rightText.Substring(0, rightText.Length - 1);

					text = new Text(rightText + ";") { Space = SpaceProcessingModeValues.Preserve };
					element.AppendChild(
						new Run(text)
						{
							RunProperties = new RunProperties
							{
								RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
								FontSize = new FontSize { Val = "24" },
							}
						});
				}

				if (!(text is null))
					text.Text = text.Text.Substring(0, text.Text.LastIndexOf(';')) + ".";

				//пустая строка
				element = new Paragraph();
				currentElement.InsertAfterSelf(element);
				currentElement = element;
			}

			bookmarkElement.Remove();
		}

		private void WriteYear(IDictionary<string, BookmarkStart> bookmarkMap) =>
			FindElementsByBookmark<Text>(bookmarkMap["Year1"], 1)
				.First(elem => elem.Text.Contains("AutofillYear"))
				.Text = DateTime.Today.Year.ToString();

		private void WriteLaboriousnessTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline)
		{
			if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(discipline))
				return;

			var rows = FindElementsByBookmark<Table>(bookmarkMap["LaboriousnessTable1"], 2)
				.First()
				.Elements<TableRow>()
				.ToArray();


			//Добавляет в первую строку ячейку, которая далее будет сливаться с последующими
			rows.First().AppendChild(
				new TableCell(
					new TableCellProperties(
						new HorizontalMerge { Val = MergedCellValues.Restart }),
					new Paragraph(
						new Run(
							new Text("Трудоемкость, академических часов")))));

			foreach (var (semester, details) in Section.Disciplines[discipline].Details)
			{
				var i = 0;

				rows.First().AppendChild(new TableCell(
					new TableCellProperties(
						new HorizontalMerge { Val = MergedCellValues.Continue }),
					new Paragraph()));
				rows.Skip(1).First().AppendChild(GetTableCellByString($"{semester} семестр"));

				foreach (var row in rows.Skip(2))
					row.AppendChild(GetTableCellByString(details[i++]));
			}

			WriteAtAllColumn();


			void WriteAtAllColumn()
			{
				//if (Disciplines[discipline].Details.Count == 0) return;
				var i = 0;
				var atAll = new DisciplineDetails
				{
					Monitoring = Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring)
						.Aggregate((current, next) => current + ", " + next),
					Contact = Section.Disciplines[discipline].Details.Values.Sum(details => details.Contact),
					Lec = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lec),
					Lab = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lab),
					Pr = Section.Disciplines[discipline].Details.Values.Sum(details => details.Pr),
					Ind = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ind),
					Control = Section.Disciplines[discipline].Details.Values.Sum(details => details.Control),
					Ze = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ze)
				};

				rows.First().AppendChild(new TableCell(
					new TableCellProperties(
						new HorizontalMerge { Val = MergedCellValues.Continue }),
					new Paragraph()));
				rows.Skip(1).First().AppendChild(GetTableCellByString("всего"));

				foreach (var row in rows.Skip(2))
					row.AppendChild(GetTableCellByString(atAll[i++]));
			}
		}
	}
}
