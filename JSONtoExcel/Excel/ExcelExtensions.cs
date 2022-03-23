﻿using OfficeOpenXml;

namespace JSONtoExcel.Excel
{
	public static class ExcelExtensions
	{
		public static void SetTrueColumnWidth(this ExcelColumn column, double width)
		{
			//// Deduce what the column width would really get set to.
			//var z = width >= (1 + 2 / 3)
			//	? Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2)
			//	: Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);

			//// How far off? (will be less than 1)
			//var errorAmt = width - z;

			//// Calculate what amount to tack onto the original amount to result in the closest possible setting.
			//var adj = width >= 1 + 2 / 3
			//	? Math.Round(7 * errorAmt - 7 / 256, 0) / 7
			//	: Math.Round(12 * errorAmt - 12 / 256, 0) / 12 + (2 / 12);

			//// Set width to a scaled-value that should result in the nearest possible value to the true desired setting.
			//if (z > 0)
			//{
			//	column.Width = width + adj;
			//	return;
			//}

			//column.Width = 0d;
			if (width < 1)
			{
				column.Width = (12.0 / 7) * width;
			}
			else
			{
				column.Width = width + (5.0 / 7);
			}
		}
	}
}