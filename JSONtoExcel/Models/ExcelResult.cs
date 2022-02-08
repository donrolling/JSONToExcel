using OfficeOpenXml;

namespace JSONtoExcel.Models
{
	public class ExcelResult : IDisposable
	{
		public string Directory { get; set; }
		public ExcelPackage ExcelPackage { get; set; }
		public string Filename { get; set; }

		public void Dispose()
		{
			ExcelPackage.Dispose();
		}
	}
}