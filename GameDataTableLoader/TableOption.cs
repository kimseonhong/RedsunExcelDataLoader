using System.Text.Json;

namespace GameDataTableLoader
{
	public class TableOption
	{
		public static string ExcelPath = string.Empty;
		public static string OutputPath = string.Empty;

		public static bool UseExportXMLParse = false;
		public static string ExportXMLPath = string.Empty;

		public delegate string DelegateTableJsonSerializer(object data);
		public static DelegateTableJsonSerializer TableJsonSerializer = JsonSerialize;

		public delegate object? DelegateTableJsonDeserializer(string json, Type type);
		public static DelegateTableJsonDeserializer TableJsonDeserializer = JsonDeserialize;

		private static string JsonSerialize(object data) => JsonSerializer.Serialize(data);
		private static object? JsonDeserialize(string data, Type type) => JsonSerializer.Deserialize(data, type);
	}
}
