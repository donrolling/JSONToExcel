using OfficeOpenXml;

namespace JSONtoExcel.Excel
{
	public static class DynamicListToExcel
	{
		public static void DynamicItemToExcelWorksheet(ExcelPackage excelPackage, string worksheetName, Dictionary<string, object> item)
		{
			if (item == null)
			{
				return;
			}
			var ws = excelPackage.Workbook.Worksheets.Add(worksheetName);
			var row = 1;
			var column = 1;
			// create the header row
			//var properties = JObjectToObject.GetPropertyKeysForDynamic(item);
			foreach (var header in item.Keys)
			{
				ExcelFormattingEngine.SetHeader(ws.Cells[row, column], JObjectToObject.FirstCharToUpper(header));
				column++;
			}
			row++;
			column = 1;
			foreach (var prop in item.Keys)
			{
				var value = item[prop].ToString();
				if (value != null && !string.IsNullOrWhiteSpace(value))
				{
					ws.Cells[row, column].Value = value;
				}
				column++;
			}
			ExcelFormattingEngine.AutoFit_All_Columns(ws);
		}

		public static void DynamicListToExcelWorksheet(ExcelPackage excelPackage, string worksheetName, dynamic items)
		{
			if (items == null)
			{
				return;
			}
			var ws = excelPackage.Workbook.Worksheets.Add(worksheetName);
			var row = 1;
			var column = 1;
			// create the header row
			var properties = JObjectToObject.GetPropertyKeysForDynamic(items[0]);
			foreach (var header in properties)
			{
				ExcelFormattingEngine.SetHeader(ws.Cells[row, column], JObjectToObject.FirstCharToUpper(header));
				column++;
			}
			foreach (var item in items)
			{
				row++;
				column = 1;
				foreach (var prop in properties)
				{
					var value = item[prop].ToString();
					if (value != null && !string.IsNullOrWhiteSpace(value))
					{
						ws.Cells[row, column].Value = value;
					}
					column++;
				}
			}
			ExcelFormattingEngine.AutoFit_All_Columns(ws);
		}
	}
}