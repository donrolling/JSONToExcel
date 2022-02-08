using JSONtoExcel.Models;
using OfficeOpenXml;

namespace JSONtoExcel.Utilities
{
	public static class ExcelFileUtility
	{
		public static ExcelResult CopyExcelFile(ExcelResult fileToCopy, string newFileName)
		{
			//gotta read it again, because the stream will be closed
			using (var file = ReadFile(fileToCopy.Filename, fileToCopy.Directory))
			{
				var result = new ExcelResult
				{
					Filename = newFileName,
					Directory = fileToCopy.Directory,
					ExcelPackage = file.ExcelPackage
				};
				Save(result);
				return result;
			}
		}

		public static ExcelResult ReadFile(string filename, string directory)
		{
			using (var filestream = FileUtility.OpenRead<ExcelResult>(CleanseFilename(filename), directory))
			{
				return new ExcelResult
				{
					Filename = filename,
					Directory = directory,
					ExcelPackage = new ExcelPackage(filestream)
				};
			}
		}

		public static ExcelResult ReadFile(string path)
		{
			var filename = Path.GetFileName(path);
			var directory = Path.GetDirectoryName(path);
			return ReadFile(filename, directory);
		}

		public static void Save(string filename, string directory, byte[] data)
		{
			FileUtility.WriteFile<ExcelResult>(filename, directory, data);
		}

		public static void Save(ExcelResult excelResult)
		{
			if (string.IsNullOrEmpty(excelResult.Directory))
			{
				throw new Exception("excelResult.Directory cannot be null or empty.");
			}
			var filename = CleanseFilename(excelResult.Filename);
			FileUtility.WriteFile<ExcelResult>(filename, excelResult.Directory, excelResult.ExcelPackage.GetAsByteArray());
		}

		public static void SaveExcel(string filename, string directory, ExcelPackage result)
		{
			FileUtility.WriteFile<ExcelResult>(CleanseFilename(filename), directory, result);
		}

		public static MemoryStream ToMemoryStream(ExcelResult excelResult)
		{
			return ToMemoryStream(excelResult.ExcelPackage);
		}

		public static MemoryStream ToMemoryStream(ExcelPackage excelPackage)
		{
			return new MemoryStream(excelPackage.GetAsByteArray());
		}

		public static byte[] ToByteArray(ExcelPackage excelPackage)
		{
			return excelPackage.GetAsByteArray();
		}

		public static string CleanseFilename(string filename)
		{
			var fn = filename.Replace(".xlsx", "").Replace(".xls", "");
			return FileUtility.CleanFilename(fn, "_") + ".xlsx";
		}
	}
}