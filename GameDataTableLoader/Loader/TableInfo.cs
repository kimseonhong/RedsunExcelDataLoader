using GameDataTableLoader.Exporter;
using OfficeOpenXml;

namespace GameDataTableLoader.Loader
{
	public class FileTableData
	{
		public List<dynamic> Data { get; set; }
		public string FileName { get; set; }
	}

	public class TableInfo
	{
		private DataType _dataType;
		private string _tableName;

		private List<dynamic> _tableData = new();
		private Dictionary<long, dynamic> _tableDataMap = new();
		private Dictionary<long, string> _keyFileFullName = new();
		private Dictionary<string, FileTableData> _fileTableData = new();

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
			_fileTableData.Clear();

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
						string type = worksheet.Cells[1, column].Value.ToString() ?? string.Empty;
						_types.Add(type);
						string name = worksheet.Cells[2, column].Value.ToString() ?? string.Empty;
						_names.Add(name);
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

				// 파일기준 전체 데이터
				_fileTableData.Add(file.FullName, new FileTableData() { FileName = file.Name, Data = infos.Values.ToList() });

				return;
			}
		}


		private Dictionary<string, ExportParser> _excels = new();
		public void Save()
		{
			foreach (var tableData in _fileTableData)
			{
				if (false == _excels.TryGetValue(tableData.Key, out var excel))
				{
					excel = new ExportParser(_dataType);
					excel.Run(isFileSaved: false);
					_excels.Add(tableData.Key, excel);
				}
				excel.AddData(tableData.Value.Data, out var a);
			}

			foreach (var excel in _excels)
			{
				File.Delete(excel.Key);
				excel.Value.SaveFile(excel.Key);
			}
		}
	}
}
