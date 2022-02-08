using Newtonsoft.Json.Linq;

namespace JSONtoExcel.Excel
{
	public static class JObjectToObject
	{
		public static List<dynamic> ConvertArray(JArray array)
		{
			if (array == null)
			{
				return null;
			}
			var result = new List<dynamic>();
			foreach (var item in array)
			{
				var x = Convert(item);
				if (x != null)
				{
					result.Add(x);
				}
			}
			return result;
		}

		public static dynamic Convert(JToken token)
		{
			if (token == null)
			{
				return null;
			}
			var jobject = token as JObject;
			return jobject.ToObject<dynamic>();
		}

		public static List<string> GetPropertyKeysForDynamic(dynamic dynamicToGetPropertiesFor)
		{
			JObject attributesAsJObject = dynamicToGetPropertiesFor;
			Dictionary<string, object> values = attributesAsJObject.ToObject<Dictionary<string, object>>();
			List<string> toReturn = new List<string>();
			foreach (string key in values.Keys)
			{
				toReturn.Add(key);
			}
			return toReturn;
		}

		public static Dictionary<string, object> RemoveProperties(JToken item, List<string> removeProps)
		{
			var dictionary = item.ToObject<Dictionary<string, object>>();
			foreach (var key in removeProps)
			{
				dictionary.Remove(key);
			}
			return dictionary;
		}

		public static string FirstCharToUpper(this string input) => input switch
		{
			null => throw new ArgumentNullException(nameof(input)),
			"" => throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
			_ => input[0].ToString().ToUpper() + input.Substring(1)
		};
	}
}