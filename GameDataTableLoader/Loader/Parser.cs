using OfficeOpenXml;
using System.Reflection;
using System.Text.Json;

namespace GameDataTableLoader.Loader
{
	public class Parser
	{
		private DataType _dataType;
		private string _tableName = string.Empty;
		private ExcelWorksheet _worksheet = default;
		private List<string> _types = new List<string>();
		private List<string> _names = new List<string>();

		private string _currentFileName = string.Empty;
		private int _rowCount = 0;
		private int _columnCount = 0;

		public Parser(DataType dataType, string tableName, List<string> types, List<string> names)
		{
			_dataType = dataType;
			_tableName = tableName;
			_types = types;
			_names = names;
		}

		public void SetWorkSheet(ExcelWorksheet worksheet, string currentFileName)
		{
			_worksheet = worksheet;
			_currentFileName = currentFileName;
			_rowCount = worksheet.Dimension.Rows;
			_columnCount = worksheet.Dimension.Columns;
		}

		public Dictionary<long, dynamic> RunParser()
		{
			int rowIndex = 3;
			int columnIndex = 1;

			List<dynamic> infos = new List<dynamic>();
			Dictionary<long, dynamic> _tableDataMap = new Dictionary<long, dynamic>();

			long tid = 0;
			string tidValue = "";

			for (; rowIndex < _rowCount;)
			{
				string selectTid = _worksheet.Cells[rowIndex, columnIndex].Value?.ToString() ?? string.Empty;
				if (false == string.IsNullOrEmpty(tidValue))
				{
					if (true == string.IsNullOrEmpty(selectTid)
					|| tidValue == selectTid)
					{
						rowIndex++;
						continue;
					}
				}

				tidValue = selectTid;
				if (false == long.TryParse(tidValue, out tid))
				{
					throw new Exception($"FileName: {_currentFileName}, Tid value is not Int64 | Parser.RunParse()");
				}

				if (true == _tableDataMap.ContainsKey(tid))
				{
					throw new Exception($"FileName: {_currentFileName}, Has Already Tid, {tid} | Parser.RunParse()");
				}

				var data = ColParser(ref rowIndex, ref columnIndex);
				infos.Add(data);
				_tableDataMap.Add(tid, data);

				rowIndex++;
				columnIndex = 1;
			}

			return _tableDataMap;
		}

		private Dictionary<string, dynamic> ColParser(ref int row, ref int col, string currentClassName = "")
		{
			var dataInfo = new Dictionary<string, dynamic>();
			for (; col <= _columnCount;)
			{
				string value = _worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
				string colType = _types[col - 1];
				if (false == string.IsNullOrEmpty(currentClassName))
				{
					if (false == colType.StartsWith(currentClassName))
					{
						break;
					}
					colType = _types[col - 1].Replace($"{currentClassName}.", string.Empty);
				}

				if (true == PrimitiveType(colType))
				{
					dataInfo.Add(_names[col - 1], ValueParser(value, colType));
					col++;
					continue;
				}
				else if (colType.StartsWith("Class"))
				{
					string name = _names[col - 1];
					string className = colType.Replace("Class::", "");
					col++;
					dataInfo.Add(name, ColParser(ref row, ref col, className));
					continue;
				}
				else if (colType.StartsWith("List"))
				{
					dataInfo.Add(_names[col - 1], ListParser(col, row, out col));
					continue;
				}
				else
				{
					// 만약에...Class 라면?

				}
			}
			return dataInfo;
		}

		private List<dynamic> ListParser(int indexColValue, int row, out int maxCol)
		{
			List<dynamic> infos = new List<dynamic>();
			bool isFrist = true;
			maxCol = indexColValue;

			for (; row <= _rowCount;)
			{
				int listColValue = indexColValue + 1;
				string? indexValue = _worksheet.Cells[row, indexColValue].Value?.ToString() ?? string.Empty;
				if (true == string.IsNullOrEmpty(indexValue)
					|| false == isFrist && indexValue == "0")
				{
					break;
				}
				isFrist = false;

				string colType = _types[indexColValue - 1];
				string className = "";
				// 클래스 형 리스트임으로 클래스를 분리해야함 // 자르는게 낫겠다
				if (colType.Contains("List<Class::"))
				{
					className = colType.Replace(">", string.Empty).Split("List<Class::")[1];
				}

				infos.Add(ColParser(ref row, ref listColValue, className));
				maxCol = listColValue; // 아마 최대치일듯
				row++;
			}
			return infos;
		}

		private bool PrimitiveType(string cellType)
		{
			switch (cellType)
			{
				case "Int64":
				case "Int32":
				case "Int16":
				case "UInt64":
				case "UInt32":
				case "UInt16":
				case "Floot":
				case "Double":
				case "Boolean":
				case "String":
				case string type when type.StartsWith("Enum"):
					return true;
			}
			return false;
		}

		private dynamic ValueParser(string value, string cellType)
		{
			switch (cellType)
			{
				case "Int64":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return long.Parse(value);
				case "Int32":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return int.Parse(value);
				case "Int16":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return short.Parse(value);
				case "UInt64":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return ulong.Parse(value);
				case "UInt32":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return uint.Parse(value);
				case "UInt16":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return ushort.Parse(value);
				case "Floot":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return float.Parse(value);
				case "Double":
					if (string.IsNullOrEmpty(value))
					{
						value = "0";
					}
					return double.Parse(value);
				case "Boolean":
					if (string.IsNullOrEmpty(value))
					{
						value = "False";
					}
					return bool.Parse(value);
				case "String":
					return value;
				case string type when type.StartsWith("Enum"):
					{
						string typeName = type.Replace("Enum::", "");
						Type? enumType = _dataType.FindPropertyType(typeName);
						if (null != enumType)
						{
							return Enum.Parse(enumType, value.ToUpper());
						}
					}
					break;

				case string type when type.StartsWith("Class"):
					return null;
				case string type when type.StartsWith("List"):
					return null;
			}

			return null;
		}
	}
}
