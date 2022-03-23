using JSONtoExcel.Excel;
using JSONtoExcel.Models;
using JSONtoExcel.Utilities;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

var filename = "data.json";
var path = FileUtility.GetBasePath<Program>();
var localPath = $"{path}\\Data";
var fullPath = $"{localPath }\\{filename}";
var json = FileUtility.ReadTextFile(fullPath);
var comma = ", ";
var lineBreak = "\r\n";
var wrappingCells = new List<string>
{
	"ReviewerName",
	"ReviewerUsername",
	"SubmittedDate",
	"LittlerPrivate",
	"BasicReview",
	"OverallPerformance",
	"AverageRating",
	"SkillsExpertisePerformance"
};

using (var excelPackage = new ExcelPackage())
{
	var worksheetName = "Worksheet";
	var ws = excelPackage.Workbook.Worksheets.Add(worksheetName);
	var row = 1;
	var column = 1;
	// create the header row
	var properties = typeof(Output).GetProperties();
	foreach (var property in properties)
	{
		var displayName = string.Concat(property.Name.Select(x => Char.IsUpper(x) ? " " + x : x.ToString())).TrimStart(' ');
		ExcelFormattingEngine.SetHeader(ws.Cells[row, column], displayName);
		column++;
	}
	var data = JArray.Parse(json);
	foreach (var item in data)
	{
		if (item == null)
		{
			continue;
		}

		var services = JObjectToObject.ConvertArray(item["service"] as JArray);
		var specialties = JObjectToObject.ConvertArray(item["specialty"] as JArray);
		var confirmedRelationships = JObjectToObject.ConvertArray(item["confirmedRelationships"] as JArray);
		var reviews = JObjectToObject.ConvertArray(item["reviews"] as JArray);

		var serviceString = ConcatenateValues(services, "Description", comma);
		var skillsExpertisePerformance = ConcatenateValues(reviews, "skillsExpertisePerformance", lineBreak);
		var overallPerformance = ConcatenateValues(reviews, "overallPerformance", lineBreak);
		var basicReview = ConcatenateValues(reviews, "basicReview", lineBreak);
		var averageRating = ConcatenateValues(reviews, "averageRating", lineBreak);
		var reviewerName = ConcatenateValues(reviews, "reviewerName", lineBreak);
		var reviewerUsername = ConcatenateValues(reviews, "reviewerUsername", lineBreak);
		var submittedDate = ConcatenateValues(reviews, "submittedDate", lineBreak);
		var littlerPrivate = ConcatenateValues(reviews, "littlerPrivate", lineBreak);

		var output = new Output
		{
			FirstName = GetTokenValue(item, "firstName"),
			MiddleName = GetTokenValue(item, "middleName"),
			LastName = GetTokenValue(item, "lastName"),
			Suffix = GetTokenValue(item, "suffix"),
			CourtName = GetTokenValue(item, "courtName"),
			City = GetTokenValue(item, "city"),
			State = GetTokenValue(item, "State"),
			Services = serviceString,
			ReviewerName = reviewerName,
			ReviewerUsername = reviewerUsername,
			SubmittedDate = submittedDate,
			LittlerPrivate = littlerPrivate,
			BasicReview = basicReview,
			OverallPerformance = overallPerformance,
			AverageRating = averageRating,
			SkillsExpertisePerformance = skillsExpertisePerformance
		};

		row++;
		column = 1;
		foreach (var prop in properties)
		{
			var objValue = prop.GetValue(output);
			var value = objValue == null
				? string.Empty
				: objValue.ToString();
			if (!string.IsNullOrWhiteSpace(value))
			{
				ws.Cells[row, column].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
				ws.Cells[row, column].Value = value;
			}
			column++;
		}
	}
	// wrapped columns need a width to be set
	ExcelFormattingEngine.AutoFit_All_Columns(ws);
	SetWidth(ws.Column(8), 30, true);
	SetWidth(ws.Column(9), 30, true);
	SetWidth(ws.Column(10), 30, true);
	SetWidth(ws.Column(11), 30, true);
	SetWidth(ws.Column(12), 30, true);
	SetWidth(ws.Column(13), 30, true);
	SetWidth(ws.Column(14), 30, true);
	SetWidth(ws.Column(15), 30, true);
	SetWidth(ws.Column(16), 500, true);
	var date = DateTime.Now.ToString("dd-MM-yyyy");
	var bytes = ExcelFileUtility.ToByteArray(excelPackage);
	ExcelFileUtility.Save($"Mediator-{date}.xlsx", localPath, bytes);
}

bool IsWrappingCell(string name)
{
	return wrappingCells.Contains(name);
}

string GetDyanmicValue(dynamic item, string propertyName)
{
	var valueObject = item[propertyName];
	if (valueObject == null)
	{
		return string.Empty;
	}
	return valueObject.ToString();
}

string GetTokenValue(JToken item, string propertyName)
{
	var valueObject = item[propertyName];
	if (valueObject == null)
	{
		return string.Empty;
	}
	return valueObject.ToString();
}

string ConcatenateValues(List<dynamic> input, string propertyName, string join)
{
	if (input == null)
	{
		return string.Empty;
	}
	var list = new List<string>();
	foreach (var item in input)
	{
		var value = GetDyanmicValue(item, propertyName);
		list.Add(value);
	}
	var result = !list.Any()
		? string.Empty
		: string.Join(join, list);
	return result;
}

void SetWidth(ExcelColumn excelColumn, int width, bool bestFit)
{
	excelColumn.Style.WrapText = true;
	excelColumn.SetTrueColumnWidth(width);
	excelColumn.BestFit = bestFit;
}