using DisciplineWorkProgram.Extensions;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models.Sections.Helpers
{
	public class Competencies
	{
		public static IEnumerable<string> ParseCompetencies(WordprocessingDocument document)
		{

            /*foreach (var cell in GetTablesCells(document)
				.Where(cell => cell.Descendants<Text>().Any(text => RegexPatterns.Competence.IsMatch(text.Text))))
			{

				var tmp = string.Empty;
				foreach (var text in cell.Descendants<Text>())
				{
					if (RegexPatterns.Competence.IsMatch(text.Text) && !string.IsNullOrEmpty(tmp))
					{
						competencies.Add(tmp.RemoveMultipleSpaces());
						tmp = string.Empty;
					}
					tmp += text.Text + " "; //тут добавлялся пробел... зачем?
					tmp.Replace(". ", ".");
				}
				if (!string.IsNullOrEmpty(tmp)) competencies.Add(tmp.RemoveMultipleSpaces());
			}*/
            var competencies = new List<string>();
            foreach (var cell in GetTablesCells(document))
            {
                var tmp2 = cell.InnerText;
				var ss = Regex.Split(tmp2, RegexPatterns.Separator.ToString(), RegexOptions.ExplicitCapture);
                var ss2 = RegexPatterns.Separator.Matches(tmp2);
				for(int i = 1; i < ss.Length; i++)
				{
					competencies.Add(ss2[i-1] + " " + ss[i]);
				}
            }
            return competencies;
		}
	}
}
