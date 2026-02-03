using OfficeOpenXml;
using System.Reflection;
using System.Text.Json;

namespace GameDataTableLoader.Loader
{
	public class LoaderParser
	{
		private DataType _dataType;
		private string _tableName = string.Empty;
		private ExcelWorksheet _worksheet = default;
		private List<string> _types = new List<string>();
		private List<string> _names = new List<string>();

		private string _currentFileName = string.Empty;
		private int _rowCount = 0;
		private int _columnCount = 0;

		public LoaderParser(DataType dataType, string tableName, List<string> types, List<string> names)
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
			int rowIndex = 4;
			int columnIndex = 1;

			List<dynamic> infos = new List<dynamic>();
			Dictionary<long, dynamic> _tableDataMap = new Dictionary<long, dynamic>();

			long tid = 0;
			string tidValue = "";

			for (; rowIndex <= _rowCount;)
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

				if (!string.IsNullOrEmpty(currentClassName))
				{
					if (!colType.StartsWith(currentClassName))
						break;

					colType = _types[col - 1].Replace($"{currentClassName}.", string.Empty);
				}

				if (colType.StartsWith("OneOf::"))
				{
					string oneofName = colType.Replace("OneOf::", "");
					col++; // oneof 마커 소비

					// ✅ currentClassName을 넘겨서 prefix 제거 기준을 맞춤
					var selected = ParseOneOf(ref row, ref col, oneofName, currentClassName);

					if (selected != null)
						dataInfo[selected.Value.FieldName] = selected.Value.Value;

					continue;
				}

				if (PrimitiveType(colType))
				{
					dataInfo.Add(_names[col - 1], ValueParser(value, colType, emptyAsNull: false)!);
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
					col++;
				}
			}

			return dataInfo;
		}

		private (string FieldName, dynamic Value)? ParseOneOf(ref int row, ref int col, string oneofName, string currentClassName)
		{
			string oneofPrefix = oneofName + "."; // "param."
			(string FieldName, dynamic Value)? selected = null;

			while (col <= _columnCount)
			{
				// ✅ 원본 타입에서 currentClassName prefix 제거(ColParser와 동일)
				string rawType = _types[col - 1];
				if (!string.IsNullOrEmpty(currentClassName))
				{
					if (!rawType.StartsWith(currentClassName + "."))
						break;

					rawType = rawType.Substring(currentClassName.Length + 1); // remove "currentClassName."
				}

				// 이제 rawType은 "param.Int32" / "param.Class::X" 형태여야 함
				if (!rawType.StartsWith(oneofPrefix))
					break;

				string inner = rawType.Substring(oneofPrefix.Length); // "Int32" or "Class::TableInfo.Param_MeleeArc"
				string rawValue = _worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;

				// (A) class 케이스
				if (inner.StartsWith("Class::"))
				{
					string className = inner.Replace("Class::", ""); // "TableInfo.Param_MeleeArc"
					string fieldName = _names[col - 1];             // "MeleeArc" 같은 이름
					col++; // 후보 시작 칸 소비

					var dict = ColParserSparse(ref row, ref col, className);
					if (dict.Count > 0)
					{
						if (selected != null)
							throw new Exception($"FileName: {_currentFileName}, OneOf({oneofName}) multiple: {selected.Value.FieldName}, {fieldName}");
						selected = (fieldName, dict);
					}
					continue;
				}

				// (B) primitive/enum 케이스
				{
					string fieldName = _names[col - 1];
					var parsed = ValueParser(rawValue, inner, emptyAsNull: true);

					if (parsed != null)
					{
						if (selected != null)
							throw new Exception($"FileName: {_currentFileName}, OneOf({oneofName}) multiple: {selected.Value.FieldName}, {fieldName}");
						selected = (fieldName, parsed);
					}

					col++; // primitive 후보 칸 소비
					continue;
				}
			}

			return selected;
		}


		private (string FieldName, Dictionary<string, dynamic> Value)? ParseOneOfCandidates(
	ref int row, ref int col, string oneofPrefix)
		{
			(string FieldName, Dictionary<string, dynamic> Value)? selected = null;

			while (col <= _columnCount)
			{
				string t = _types[col - 1]; // 예: "param.Class::Param_MeleeArc" 또는 "Param_MeleeArc.Double"
				if (!t.StartsWith(oneofPrefix))
					break;

				string inner = t.Substring(oneofPrefix.Length); // "Class::Param_MeleeArc"
				if (!inner.StartsWith("Class::"))
					throw new Exception($"FileName: {_currentFileName}, OneOf candidate must be Class, but: {t}");

				string candidateMsgName = inner.Replace("Class::", ""); // "Param_MeleeArc"
				string candidateFieldName = _names[col - 1];           // 보통 "Param_MeleeArc"

				// 후보 클래스 컬럼 소비
				col++;

				// 후보 클래스 내부 필드 파싱:
				// 기존 ColParser를 재사용하되, 이 경우는 "빈칸=0"이 아니라 "빈칸=null"이어야 함.
				// => currentClassName = candidateMsgName 으로 파싱하고,
				//    ValueParser 호출에 emptyAsNull=true가 적용되도록 별도 함수가 필요.

				var candidateDict = ColParserSparse(ref row, ref col, candidateMsgName);

				bool hasAny = candidateDict.Count > 0; // sparse 파싱이라 값이 있으면 키가 생김

				if (hasAny)
				{
					if (selected != null)
						throw new Exception($"FileName: {_currentFileName}, OneOf has multiple candidates set: {selected.Value.FieldName}, {candidateFieldName}");

					selected = (candidateFieldName, candidateDict);
				}
				else
				{
					// 비었으면 선택 안 함 (그대로 진행)
				}
			}

			return selected;
		}

		private Dictionary<string, dynamic> ColParserSparse(ref int row, ref int col, string currentClassName)
		{
			var dataInfo = new Dictionary<string, dynamic>();

			for (; col <= _columnCount;)
			{
				string rawValue = _worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
				string colType = _types[col - 1];

				// class boundary
				if (!colType.StartsWith(currentClassName))
					break;

				colType = colType.Replace($"{currentClassName}.", string.Empty);

				if (PrimitiveType(colType))
				{
					var parsed = ValueParser(rawValue, colType, emptyAsNull: true);
					if (parsed != null)
						dataInfo[_names[col - 1]] = parsed;

					col++;
					continue;
				}
				else if (colType.StartsWith("Class"))
				{
					string name = _names[col - 1];
					string className = colType.Replace("Class::", "");
					col++;

					var child = ColParserSparse(ref row, ref col, className);
					if (child.Count > 0)
						dataInfo[name] = child;

					continue;
				}
				else if (colType.StartsWith("List"))
				{
					// 필요하면 Sparse List도 구현 가능.
					throw new Exception($"FileName: {_currentFileName}, OneOf candidate contains List (not supported in sparse mode).");
				}
				else
				{
					col++;
				}
			}

			return dataInfo;
		}

		private List<dynamic> ListParser(int indexColValue, int row, out int maxCol)
		{
			List<dynamic> infos = new List<dynamic>();
			bool isFirst = true;
			maxCol = indexColValue + 1;

			for (; row <= _rowCount;)
			{
				int listColValue = indexColValue + 1;
				string? indexValue;
				if (true == isFirst)
				{
					indexValue = _worksheet.Cells[row, indexColValue].Value?.ToString() ?? string.Empty;
					if (true == string.IsNullOrEmpty(indexValue))
					{
						break;
					}
				}
				else
				{
					indexValue = _worksheet.Cells[row + 1, indexColValue].Value?.ToString() ?? string.Empty;
					if (true == string.IsNullOrEmpty(indexValue)
						|| indexValue == "0")
					{
						break;
					}
					row++;
				}
				isFirst = false;

				string colType = _types[indexColValue - 1];
				string className = "";
				// 클래스 형 리스트임으로 클래스를 분리해야함 // 자르는게 낫겠다
				if (colType.Contains("List<Class::"))
				{
					className = colType.Replace(">", string.Empty).Split("List<Class::")[1];
				}

				infos.Add(ColParser(ref row, ref listColValue, className));
				maxCol = listColValue; // 아마 최대치일듯
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

		private dynamic? ValueParser(string value, string cellType, bool emptyAsNull)
		{
			if (emptyAsNull && string.IsNullOrEmpty(value))
				return null;

			switch (cellType)
			{
				case "Int64": return long.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "Int32": return int.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "Int16": return short.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "UInt64": return ulong.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "UInt32": return uint.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "UInt16": return ushort.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "Floot": return float.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "Double": return double.Parse(string.IsNullOrEmpty(value) ? "0" : value);
				case "Boolean": return bool.Parse(string.IsNullOrEmpty(value) ? "False" : value);
				case "String": return value;

				case string type when type.StartsWith("Enum"):
					{
						string typeName = type.Replace("Enum::", "");
						Type? enumType = _dataType.FindPropertyType(typeName);
						if (enumType == null) return null;

						if (string.IsNullOrEmpty(value))
						{
							if (emptyAsNull) return null; // oneof 후보면 “미선택”
							value = Enum.GetValues(enumType).GetValue(0)?.ToString() ?? string.Empty;
						}
						return Enum.Parse(enumType, value);
					}
			}

			return null;
		}

	}
}
