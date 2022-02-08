using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace JSONtoExcel.Excel
{
	public static class ExcelFormattingEngine
	{
		public static void AlignRight(ExcelRange excelRange)
		{
			excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
		}

		public static void AutoFit_All_Columns(ExcelWorksheet excelWorksheet)
		{
			AutoFitColumns(excelWorksheet.Cells[excelWorksheet.Dimension.Address]);
		}

		public static void AutoFitColumns(ExcelRange excelRange)
		{
			excelRange.AutoFitColumns();
		}

		public static string ColumnLetter_FromColumnNumber(int columnNumber)
		{
			var dividend = columnNumber;
			var columnName = string.Empty;
			var modulo = 0;
			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (dividend - modulo) / 26;
			}
			return columnName;
		}

		public static int ColumnNumber_FromColumnLetter(string columnName)
		{
			columnName = columnName.ToUpperInvariant();
			int sum = 0;
			for (int i = 0; i < columnName.Length; i++)
			{
				sum *= 26;
				sum += columnName[i] - 'A' + 1;
			}
			return sum;
		}

		public static void Format_Background_Text(ExcelRange excelRange, ExcelFillStyle excelFillStyle, ExcelHorizontalAlignment excelHorizontalAlignment, Color color, bool boldText)
		{
			excelRange.Style.Fill.PatternType = excelFillStyle;
			excelRange.Style.Fill.BackgroundColor.SetColor(color);
			excelRange.Style.Font.Bold = true;
			excelRange.Style.HorizontalAlignment = excelHorizontalAlignment;
		}

		public static void Format_RedBackground_BlackText(ExcelRange excelRange, string message)
		{
			excelRange.Value = message;
			excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
			var colFromHex = ColorTranslator.FromHtml("#FF0000");
			excelRange.Style.Fill.BackgroundColor.SetColor(colFromHex);
		}

		public static void FormatNumber(ExcelRange excelRange, int precision, decimal value)
		{
			excelRange.Value = value;
			FormatNumber(excelRange, precision);
		}

		public static void FormatNumber(ExcelRange excelRange, int precision)
		{
			var format = GetNumberFormat_WithPrecision(precision);
			excelRange.Style.Numberformat.Format = format;
		}

		public static void FormatPercentage(ExcelRange excelRange, int precision, decimal value, bool divideBy100 = true)
		{
			excelRange.Value = value == 0 ? 0 : divideBy100 ? value / 100 : value;
			FormatPercentage(excelRange, precision);
		}

		public static void FormatPercentage(ExcelRange excelRange, int precision)
		{
			excelRange.Style.Numberformat.Format = GetPercentageFormat_WithPrecision(precision);
		}

		public static ExcelPackage FromMemoryStream(MemoryStream ms)
		{
			var excelPackage = new ExcelPackage(ms);
			return excelPackage;
		}

		public static void Merge(ExcelRange range)
		{
			range.Merge = true;
		}

		public static void MergeAndCenter(ExcelRange range, string value)
		{
			MergeAndCenter(range);
			range.Value = value;
		}

		public static void MergeAndCenter(ExcelRange range)
		{
			range.Merge = true;
			range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
		}

		public static void MergeAndCenterVerticallyAndHorizontally(ExcelRange range, bool wrap = false)
		{
			range.Merge = true;
			if (wrap)
			{
				range.Style.WrapText = true;
			}
			range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
		}

		public static void SetHeader(ExcelRange cell, string value)
		{
			cell.Value = value;
			SetHeaderStyles(cell);
		}

		public static void SetHeaderStyles(ExcelRange excelRange)
		{
			excelRange.Style.Font.Bold = true;
			excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
		}

		private static string GetNumberFormat_WithPrecision(int precision)
		{
			var baseFormat = "#,##0_);[Red](#,##0)";
			if (precision == 0) { return baseFormat; }
			var zeros = new string('0', precision);
			return baseFormat.Replace("#0", $"#0.{ zeros }");
		}

		private static string GetPercentageFormat_WithPrecision(int precision)
		{
			var baseFormat = "#,##0%_);[Red](#,##0%)";
			if (precision == 0) { return baseFormat; }
			var zeros = new string('0', precision);
			return baseFormat.Replace("#0", $"#0.{ zeros }");
		}
	}
}