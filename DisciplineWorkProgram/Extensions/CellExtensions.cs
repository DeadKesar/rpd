using ClosedXML.Excel;

namespace DisciplineWorkProgram.Extensions
{
	public static class CellExtensions
	{
		public static int GetInt(this IXLCell cell)
		{
			//кто так делает то...
			try
			{
				return cell.GetValue<int>();
			}
			catch
			{
				return 0;
			}
		}
	}
}
