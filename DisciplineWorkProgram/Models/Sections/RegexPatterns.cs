using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models.Sections
{
	public class RegexPatterns
	{
		//Наименование компетенции
		public static readonly Regex CompetenceName = new Regex(@"^(УК-\d{1,2}.\D)|^(ОПК-\d{1,2}.\D)|^(ПК-\d{1,2}.\D)");
		//Любая строка, содержащая в себе информацию о компетенции
		public static readonly Regex Competence = new Regex(@"^(УК-\d{1,2})|^(ОПК-\d{1,2})|^(ПК-\d{1,2})");
		//Поиск в строке ...?
		public static readonly Regex WayNameSection = new Regex("(?<=\").*(?=\")");
		//Поиск числа в строке
		public static readonly Regex DigitInString = new Regex(@"\d");
		//будем сепарировать
        public static readonly Regex Separator = new Regex(@"(УК-(\d{1,2}.){1,3})|(ОПК-(\d{1,2}.){1,3})|(ПК-(\d{1,2}.){1,3})");

        public static readonly Regex Competence2 = new Regex(@"(УК-\d{1,2})|(ОПК-\d{1,2})|(ПК-\d{1,2})");
    }
}
