using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using static DisciplineWorkProgram.Word.Helpers.Ooxml;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using DisciplineWorkProgram.Models.Sections;
using DocumentFormat.OpenXml;
using DisciplineWorkProgram.Models.Sections.Helpers;
using System.Text;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models
{
    public class Fos
    {
        public Fos(Section section)
        {
            Section = section;
        }

        private Section Section { get; }

        //Должно обрабатывать только 1 дисциплину, чтобы "масштабировать" без доп. кода
        public void MakeFos(string templatePath, string fosDir, string discipline)
        {

            using var doc = WordprocessingDocument.CreateFromTemplate(templatePath, true);
            var bookmarkMap = GetBookmarks(doc, "Autofill");
            
            WriteSectionData(bookmarkMap, doc);
            WriteDisciplineData(bookmarkMap, discipline, doc);
            //WriteRequirements(bookmarkMap, discipline, doc);
            WriteCompetenciesTable(bookmarkMap, discipline, doc); //заполняет табличку компетенций
            //WriteDisciplinePartitionTable(bookmarkMap, discipline, doc);
            //WritePracticleClassTable(bookmarkMap, discipline, doc);
            //WriteSemesters(bookmarkMap, discipline, doc);
            //WriteCompetencies(bookmarkMap, discipline, doc);//записываем компетенции в самом начале
            //WriteYear(bookmarkMap, doc);
            // Не реализовано занесение данных по дисциплине
            //WriteLaboriousnessTable(bookmarkMap, discipline, doc);
            //WriteLaboratiesClassTable(bookmarkMap, discipline, doc);

            SaveDoc(doc, fosDir, Section.Disciplines[discipline].Name);
            doc.Dispose();
        }

        private void WriteSectionData(IDictionary<string, BookmarkStart> bookmarkMap, WordprocessingDocument doc)
        {
            foreach (var (key, bookmark) in bookmarkMap)
            {
                var actualKey = key.Substring(0, key.Length - 1);

                try
                {
                    if (Section.SectionDictionary.ContainsKey(actualKey))
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
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

        private void WriteDisciplineData(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline))
                return;

            int countForPartFormForPartners = 0;
            foreach (var (key, bookmark) in bookmarkMap)
            {
                //Надежда на то, что понадобится нумерация в рамках только одноразрядного числа, это о том что количество
                //закладок с одинаковым именем не может быть больше 10
                var actualKey = key.Substring(0, key.Length - 1);

                var d = Section.Disciplines[discipline].Props.ContainsKey(actualKey);

                if (actualKey == "Discipline")
                {
                    if (Section.Disciplines[discipline].Props.ContainsKey(actualKey))
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Section.Disciplines[discipline].Props["Name"];
                    continue;
                }

                if (actualKey == "PartType")
                {
                    if (countForPartFormForPartners == 0)
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Regex.IsMatch(discipline, @".О.") ? "Обязательная часть." : "Часть, формируемая участниками образовательных отношений.";
                    else
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Regex.IsMatch(discipline, @".О.") ? "обязательной части" : "части, формируемой участниками образовательных отношений";
                    countForPartFormForPartners++;
                    continue;
                }

                if (Section.Disciplines[discipline].Props.ContainsKey(actualKey))
                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = Section.Disciplines[discipline].Props[actualKey];
            }
        }
        private void WriteCompetenciesTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;
            //AutofillCompetenciesTable1
            var table = FindElementsByBookmark<Table>(bookmarkMap["CompetenciesTable1"], 2, doc).First();

            foreach (var competence in Section.DisciplineCompetencies[Section.Disciplines[discipline].Name])
            {
                if (!Section.Competencies.ContainsKey(competence)) continue;


                var firstCell = new TableCell(
                new TableCellProperties(
                    new VerticalMerge { Val = MergedCellValues.Restart }
                    ),
                    new Paragraph(new Run(new Text(Section.Competencies[competence].Name)))
                );

                var relatedCompetencies = Section.Competencies
                    .Where(kvp => kvp.Key.StartsWith(competence + "."))
                    .ToList();
                var cell = new TableCell();

                var paragraph = new Paragraph();
                bool isFirst = true;

                foreach (var s in relatedCompetencies)
                {
                    var row = new TableRow();
                    if (isFirst)
                    {
                        row.Append(firstCell);
                        isFirst = false;
                    }
                    else
                    {
                        row.Append(new TableCell(
                            new TableCellProperties(
                                new VerticalMerge { Val = MergedCellValues.Continue }
                            )
                        ));
                    }
                    paragraph = new Paragraph();
                    paragraph.AppendChild(new Run(new Text(s.Value.Name)));
                    var tableCell = new TableCell(paragraph);
                    row.AppendChild(tableCell);
                    row.AppendChild(new TableCell());
                    row.AppendChild(new TableCell());
                    table.AppendChild(row);

                }

            }
        }
    }
}
