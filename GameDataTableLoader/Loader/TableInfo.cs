using OfficeOpenXml;

namespace GameDataTableLoader.Loader
{
	public class TableInfo
	{
		private DataType _dataType;
		private string _tableName;

		private List<dynamic> _tableData = new();
		private Dictionary<long, dynamic> _tableDataMap = new();
		private Dictionary<long, string> _keyFileFullName = new();

		private List<string> _types = new List<string>();
		private List<string> _names = new List<string>();

		private Parser? _parser;

		public List<dynamic> GetData() { return _tableData; }

		public TableInfo(DataType dataType, string tableName)
		{
			_dataType = dataType;
			_tableName = tableName;
		}

		public void Clear()
		{
			_tableData.Clear();
			_tableDataMap.Clear();
			_keyFileFullName.Clear();

			_parser = null;
			_types.Clear();
			_names.Clear();
		}

		public void Run(FileInfo file)
		{
			using (var excel = new ExcelPackage(file))
			{
				var worksheet = excel.Workbook.Worksheets[$"{_tableName}"];

				if (_parser == null)
				{
					int columnCount = worksheet.Dimension.Columns;
					for (int column = 1; column <= columnCount; column++)
					{
						_types.Add(worksheet.Cells[1, column].Value.ToString() ?? string.Empty);
					}

					for (int column = 1; column <= columnCount; column++)
					{
						_names.Add(worksheet.Cells[2, column].Value.ToString() ?? string.Empty);
					}

					_parser = new Parser(_dataType, _tableName, _types, _names);
				}

				_parser.SetWorkSheet(worksheet, file.FullName);
				var infos = _parser.RunParser();

				foreach (var info in infos)
				{
					if (true == _tableDataMap.ContainsKey(info.Key))
					{
						throw new Exception($"FileName: {file.FullName}, Has Already Tid, {info.Key}, Already File: {_keyFileFullName[info.Key]} | TableInfo.Run()");
					}
					_tableData.Add(info.Value);
					_tableDataMap.Add(info.Key, info.Value);
					_keyFileFullName.Add(info.Key, file.FullName);
				}

				return;
			}
		}
	}
}
