using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models.Sections
{
	public class RegexPatterns
	{
		//Наименование компетенции
		public static readonly Regex CompetenceName = new Regex(@"^(УК-\d{1,2}.?(! )?(?!\d))|^(ОПК-\d{1,2}.?(! )?(?!\d))|^(ПК-\d{1,2}.?(! )?(?!\d))");
		//Любая строка, содержащая в себе информацию о компетенции
		public static readonly Regex Competence = new Regex(@"^(УК-\d)|^(ОПК-\d)|^(ПК-\d)");
		//Поиск в строке ...?
		public static readonly Regex WayNameSection = new Regex("(?<=\").*(?=\")");
		//Поиск числа в строке
		public static readonly Regex DigitInString = new Regex(@"\d");
	}
}
