using OfficeOpenXml;

namespace GameDataTableLoader.Exporter
{
	public class TableExporter<T> where T : class, new()
	{
		private T _data = new();
		private Type _tableDataType;
		private Dictionary<string /* TableName */, DataType> _dataTypeMap = new();

		public TableExporter()
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			_tableDataType = _data.GetType();
			_dataTypeMap = DataType.PropertyParse(_tableDataType);

			Directory.CreateDirectory(@$"{TableOption.OutputPath}\_tmpl");
		}

		public void Run(string tableName = "")
		{
			// 원하는것만
			if (false == string.IsNullOrEmpty(tableName))
			{
				if (false == _dataTypeMap.ContainsKey(tableName))
				{
					throw new Exception("테이블 이름이 존재하지 않습니다.");
				}
				new ExportParser(_dataTypeMap[tableName]).Run();
				return;
			}

			// 전체 출력
			foreach (var dataType in _dataTypeMap)
			{
				new ExportParser(dataType.Value).Run();
			}
		}
	}
}
