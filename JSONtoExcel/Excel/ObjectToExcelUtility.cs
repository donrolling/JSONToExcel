using JSONtoExcel.Excel.Enum;
using JSONtoExcel.Utilities;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Reflection;

namespace JSONtoExcel.Excel
{
	public static class ObjectToExcelUtility
	{
		public static void Object_To_Excel(object data, ExcelPackage package, string worksheetName, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			if (data == null)
			{
				throw new Exception("Object_To_ExcelPackage: Null or empty input table!\n");
			}
			var properties = GetProperties(data);
			var ws = package.Workbook.Worksheets.Add(worksheetName);
			var json = JsonConvert.SerializeObject(data, Formatting.Indented, new JsonConverter[] { new StringEnumConverter() });
			var range = ws.Cells[1, 1];
			range.Value = json;
			range.Style.WrapText = true;
			range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
			ExcelFormattingEngine.AutoFit_All_Columns(ws);
			ws.Column(1).Width = 200;
		}

		public static void ObjectList_To_Excel(IEnumerable<object> data, ExcelPackage package, string worksheetName, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			if (data == null || !data.Any())
			{
				throw new Exception("Object_To_ExcelPackage: Null or empty input table!\n");
			}
			var properties = GetProperties(data.First());
			var ws = package.Workbook.Worksheets.Add(worksheetName);
			SetHeaders(properties, ws, propertyListOptions);
			SetData(properties, data.ToList(), ws, propertyListOptions);
			ExcelFormattingEngine.AutoFit_All_Columns(ws);
		}

		public static ExcelPackage ObjectList_To_ExcelPackage(IEnumerable<object> data, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var package = new ExcelPackage();
			return ObjectList_To_ExcelPackage(data, package, "Worksheet", propertyListOptions);
		}

		public static ExcelPackage ObjectList_To_ExcelPackage(IEnumerable<object> data, string worksheetName, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var package = new ExcelPackage();
			return ObjectList_To_ExcelPackage(data, package, worksheetName, propertyListOptions);
		}

		public static ExcelPackage ObjectList_To_ExcelPackage(IEnumerable<object> data, ExcelPackage package, string worksheetName, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			if (data == null || !data.Any())
			{
				throw new Exception("Object_To_ExcelPackage: Null or empty input table!\n");
			}
			var properties = GetProperties(data.First());
			var ws = package.Workbook.Worksheets.Add(worksheetName);
			SetHeaders(properties, ws, propertyListOptions);
			SetData(properties, data.ToList(), ws, propertyListOptions);
			ExcelFormattingEngine.AutoFit_All_Columns(ws);
			return package;
		}

		public static void Save_Object_To_Excel(string filename, string outputDirectory, IEnumerable<object> data, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var result = ObjectList_To_ExcelPackage(data, propertyListOptions);
			ExcelFileUtility.SaveExcel(filename, outputDirectory, result);
		}

		private static Tuple<int, int> GetDataRange(ExcelWorksheet ws, int itemIndex, int propertyIndex, ObjectToExcel_PropertyListOptions propertyListOptions, Tuple<int, int> pickupRange)
		{
			var column = 0;
			var row = 0;
			switch (propertyListOptions)
			{
				case ObjectToExcel_PropertyListOptions.HorizontalPropertyNames:
					if (pickupRange.Item1 == 0)
					{
						row = itemIndex + 2;
					}
					else
					{
						row = pickupRange.Item1;
					}
					if (pickupRange.Item2 == 0)
					{
						column = itemIndex + 1;
					}
					else
					{
						column = pickupRange.Item2 + 1;
					}
					break;

				case ObjectToExcel_PropertyListOptions.VerticalPropertyNames:
					if (pickupRange.Item1 == 0)
					{
						row = propertyIndex + 1;
					}
					else
					{
						row = pickupRange.Item1 + 1;
					}
					if (pickupRange.Item2 == 0)
					{
						column = itemIndex + 2;
					}
					else
					{
						column = pickupRange.Item2;
					}
					break;

				default:
					throw new Exception("Case not matched.");
			}
			return Tuple.Create(row, column);
		}

		private static ExcelRange GetFullHeadingRange(ExcelWorksheet ws, int count, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			switch (propertyListOptions)
			{
				case ObjectToExcel_PropertyListOptions.HorizontalPropertyNames:
					return ws.Cells[1, 1, 1, count];

				case ObjectToExcel_PropertyListOptions.VerticalPropertyNames:
					return ws.Cells[1, 1, count, 1];

				default:
					throw new Exception("Case not matched.");
			}
		}

		private static ExcelRange GetHeadingRange(ExcelWorksheet ws, int itemIndex, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var column = 0;
			var row = 0;
			switch (propertyListOptions)
			{
				case ObjectToExcel_PropertyListOptions.HorizontalPropertyNames:
					row = 1;
					column = itemIndex + 1;
					break;

				case ObjectToExcel_PropertyListOptions.VerticalPropertyNames:
					column = 1;
					row = itemIndex + 1;
					break;

				default:
					throw new Exception("Case not matched.");
			}
			return ws.Cells[row, column];
		}

		private static List<PropertyInfo> GetProperties(object obj)
		{
			return obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).ToList();
		}

		private static Tuple<int, int> SetData(List<PropertyInfo> properties, object data, ExcelWorksheet ws, ObjectToExcel_PropertyListOptions propertyListOptions, int pickupRow = 0, int pickupColumn = 0)
		{
			var itemIndex = 0;
			var pickupRange = Tuple.Create(pickupRow, pickupColumn);
			for (int propertyIndex = 0; propertyIndex < properties.Count(); propertyIndex++)
			{
				pickupRange = SetField(properties, data, ws, propertyListOptions, itemIndex, propertyIndex, pickupRange);
			}
			return pickupRange;
		}

		private static void SetData(List<PropertyInfo> properties, List<object> data, ExcelWorksheet ws, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var pickupRange = Tuple.Create(0, 0);
			for (int itemIndex = 0; itemIndex < data.Count(); itemIndex++)
			{
				for (int propertyIndex = 0; propertyIndex < properties.Count(); propertyIndex++)
				{
					pickupRange = SetField(properties, data[itemIndex], ws, propertyListOptions, itemIndex, propertyIndex, pickupRange);
				}
			}
		}

		private static Tuple<int, int> SetField(List<PropertyInfo> properties, object data, ExcelWorksheet ws, ObjectToExcel_PropertyListOptions propertyListOptions, int itemIndex, int propertyIndex, Tuple<int, int> pickupRange)
		{
			var value = properties[propertyIndex].GetValue(data, null);
			var isValueTypeOrString = true;
			Type valueType = null;
			var valueTypeName = "";
			if (value != null)
			{
				valueType = value.GetType();
				valueTypeName = valueType.Name;
				if (valueTypeName == "String")
				{
					isValueTypeOrString = true;
				}
				else
				{
					isValueTypeOrString = valueType.IsValueType;
				}
			}
			if (isValueTypeOrString)
			{
				var range = GetDataRange(ws, itemIndex, propertyIndex, propertyListOptions, pickupRange);
				var excelRange = ws.Cells[range.Item1, range.Item2];
				excelRange.Value = value;
				ExcelFormattingEngine.AlignRight(excelRange);
				return range;
			}
			else
			{
				var newProperties = GetProperties(value);
				var range = GetDataRange(ws, itemIndex, propertyIndex, propertyListOptions, pickupRange);
				var excelRange = ws.Cells[range.Item1, range.Item2];
				excelRange.Value = valueTypeName;
				ExcelFormattingEngine.AlignRight(excelRange);
				switch (propertyListOptions)
				{
					case ObjectToExcel_PropertyListOptions.HorizontalPropertyNames:
						SetData(newProperties, value, ws, ObjectToExcel_PropertyListOptions.VerticalPropertyNames, range.Item1 + 1, range.Item2);
						return Tuple.Create(range.Item1, range.Item2);

					case ObjectToExcel_PropertyListOptions.VerticalPropertyNames:
						SetData(newProperties, value, ws, ObjectToExcel_PropertyListOptions.HorizontalPropertyNames, range.Item1, range.Item2 + 1);
						return Tuple.Create(range.Item1, range.Item2);

					default:
						throw new Exception("Case not matched.");
				}
			}
		}

		private static void SetHeaders(List<PropertyInfo> properties, ExcelWorksheet ws, ObjectToExcel_PropertyListOptions propertyListOptions)
		{
			var count = properties.Count();
			for (int i = 0; i < count; i++)
			{
				var property = properties[i];
				var range = GetHeadingRange(ws, i, propertyListOptions);
				range.Value = property.Name;
			}
			var headerRange = GetFullHeadingRange(ws, count, propertyListOptions);
			switch (propertyListOptions)
			{
				case ObjectToExcel_PropertyListOptions.HorizontalPropertyNames:
					ExcelFormattingEngine.Format_Background_Text(headerRange, ExcelFillStyle.Solid, ExcelHorizontalAlignment.Center, Color.LightGray, true);
					break;

				case ObjectToExcel_PropertyListOptions.VerticalPropertyNames:
					ExcelFormattingEngine.Format_Background_Text(headerRange, ExcelFillStyle.Solid, ExcelHorizontalAlignment.Left, Color.LightGray, true);
					break;

				default:
					break;
			}
		}
	}
}