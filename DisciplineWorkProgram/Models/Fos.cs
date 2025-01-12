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
            WriteRequirements(bookmarkMap, discipline, doc);
            WriteCompetenciesTable(bookmarkMap, discipline, doc); //заполняет табличку компетенций
            WriteDisciplinePartitionTable(bookmarkMap, discipline, doc);
            WritePracticleClassTable(bookmarkMap, discipline, doc);
            WriteSemesters(bookmarkMap, discipline, doc);
            WriteCompetencies(bookmarkMap, discipline, doc);//записываем компетенции в самом начале
            WriteYear(bookmarkMap, doc);
            // Не реализовано занесение данных по дисциплине
            WriteLaboriousnessTable(bookmarkMap, discipline, doc);
            WriteLaboratiesClassTable(bookmarkMap, discipline, doc);

            SaveDoc(doc, dwpDir, Section.Disciplines[discipline].Name);
            doc.Dispose();
        }
    }
}
