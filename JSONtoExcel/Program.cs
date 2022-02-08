using JSONtoExcel.Excel;
using JSONtoExcel.Utilities;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

var filename = "data.json";
var path = FileUtility.GetBasePath<Program>();
var localPath = $"{path}\\Data";
var fullPath = $"{localPath }\\{filename}";
var json = FileUtility.ReadTextFile(fullPath);

var data = JArray.Parse(json);
foreach (var item in data)
{
	using (var excelPackage = new ExcelPackage())
	{
		var worksheetName = $"{item["id"]}-{item["firstName"]}-{item["lastName"]}";
		var services = JObjectToObject.ConvertArray(item["service"] as JArray);
		var specialties = JObjectToObject.ConvertArray(item["specialty"] as JArray);
		var confirmedRelationships = JObjectToObject.ConvertArray(item["confirmedRelationships"] as JArray);
		var reviews = JObjectToObject.ConvertArray(item["reviews"] as JArray);

		var revisedItem = JObjectToObject.RemoveProperties(item, new List<string> { "service", "specialty", "confirmedRelationships", "reviews" });
		DynamicListToExcel.DynamicItemToExcelWorksheet(excelPackage, worksheetName, revisedItem);
		DynamicListToExcel.DynamicListToExcelWorksheet(excelPackage, "Services", services);
		DynamicListToExcel.DynamicListToExcelWorksheet(excelPackage, "Specialties", specialties);
		DynamicListToExcel.DynamicListToExcelWorksheet(excelPackage, "Confirmed Relationships", confirmedRelationships);
		DynamicListToExcel.DynamicListToExcelWorksheet(excelPackage, "Reviews", reviews);
		
		var bytes = ExcelFileUtility.ToByteArray(excelPackage);
		ExcelFileUtility.Save($"{worksheetName}.xlsx", localPath, bytes);
	}
}