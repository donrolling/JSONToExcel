namespace JSONtoExcel.Models
{
	public class ReadResult
	{
		public List<string> Headers { get; set; } = new List<string>();

		public List<List<string>> Data { get; set; } = new List<List<string>>();
	}
}