using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GameDataTableLoader.Exporter
{
	public class ExportParser
	{
		private DataType _dataType;

		private List<string> _types = new List<string>();
		private List<string> _names = new List<string>();

		private ExcelPackage? _excel;
		private ExcelWorksheet? _infoWorksheet;
		private XmlCommentReader? _xmlCommentReader = default;

		public ExportParser(DataType dataType)
		{
			_dataType = dataType;
		}

		public void Run(bool isFileSaved = true)
		{
			if (true == TableOption.UseExportXMLParse)
			{
				_xmlCommentReader = new XmlCommentReader(TableOption.ExportXMLPath);
			}

			_excel = new ExcelPackage();

			_infoWorksheet = _excel.Workbook.Worksheets.Add($"{_dataType.Name}");
			TypeParser(_dataType.Type);

			// 기본 폰트와 사이즈 설정
			_infoWorksheet.Cells.Style.Font.Name = "맑은 고딕";
			_infoWorksheet.Cells.Style.Font.Size = 9;
			_infoWorksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

			for (int col = 1; col <= _types.Count; col++)
			{
				// Row 1은 타입
				SetCellData(_infoWorksheet, 1, col, _types[col - 1], ExcelHorizontalAlignment.Left, Color.FromArgb(142, 169, 219));

				// Row 2는 이름
				SetCellData(_infoWorksheet, 2, col, _names[col - 1], ExcelHorizontalAlignment.Left, Color.FromArgb(180, 198, 231));

				// Row 3은 설명(주석) 
				if (TableOption.UseExportXMLParse)
				{
					string enumValue = _xmlCommentReader?.GetMemberComment(eMemberType.PROPERTY, _dataType.FullName, _names[col - 1]) ?? string.Empty;
					//SetCellData(worksheet, 3, col, enumValue, ExcelHorizontalAlignment.Center, Color.FromArgb(51, 63, 79), Color.White);
					SetCellData(_infoWorksheet, 3, col, enumValue, ExcelHorizontalAlignment.Center, Color.FromArgb(217, 225, 242));
				}

				// 열 너비 자동 조정
				_infoWorksheet.Column(col).AutoFit();
			}

			// Drop Down Checker
			for (int i = 0; i < _types.Count; i++)
			{
				switch (_types[i])
				{
					case string type when type.Contains("Enum::"):
						{
							string enumName = type.Split("Enum::")[1];
							Type enumType = _dataType.FindPropertyType(enumName);
							AddEnumPage(_excel, enumType);
							AddEnumDropdown(_infoWorksheet, enumName, 4, i + 1);
						}
						break;
					case string type when type.Contains("Boolean"):
						{
							AddBooleanDropdown(_infoWorksheet, 4, i + 1);
						}
						break;

					default:
						break;
				}
			}

			if (true == isFileSaved)
			{
				// 파일 저장
				string tableName = _dataType.Name;
				SaveFile(@$"{TableOption.OutputPath}\_tmpl\{tableName}_tmpl.xlsx");
			}
		}
		public void SaveFile(string filePath)
		{
			if (_excel == null)
			{
				return;
			}

			// 파일 저장
			var fileInfo = new FileInfo(filePath);
			_excel.SaveAs(fileInfo);
			_excel.Dispose();
			return;
		}

		//private int currentRow = 3;
		public void AddData(List<dynamic> infos, out int lastRow, bool isList = false, int startRow = 4, int startCol = 1, int tid = 0)
		{
			lastRow = startRow;
			if (_infoWorksheet == null)
			{
				return;
			}

			if (infos.Count == 0)
			{
				return;
			}

			int row = startRow;
			int listIndex = 0;
			foreach (var info in infos)
			{
				foreach (var data in info)
				{
					AddData(data, row, ref lastRow, ref tid);
				}

				if (true == isList)
				{
					_infoWorksheet.Cells[row, startCol].Value = listIndex;
					_infoWorksheet.Cells[row, 1].Value = tid;
					listIndex++;
				}

				if (lastRow != row)
				{
					row = lastRow;
				}
				else
				{
					row++;
					lastRow = row;
				}
			}
		}

		public void AddData(KeyValuePair<string, dynamic> data, int row, ref int lastRow, ref int tid)
		{
			lastRow = Math.Max(lastRow, row);
			int col = _names.FindIndex(x => x.Equals(data.Key)) + 1;
			if (data.Value is IList)
			{
				AddData(data.Value, out lastRow, true, row, col, tid);
			}
			else if (data.Value is IDictionary)
			{
				foreach (var dicValue in data.Value)
				{
					AddData(dicValue, row, ref lastRow, ref tid);
				}
			}
			else
			{
				if (data.Key == "MetaTid")
				{
					tid = (int)data.Value;
				}

				if (col == 0) { return; }

				// 중복이슈로 처음에 사용하면 없앤다.
				_infoWorksheet.Cells[row, col].Value = data.Value;

				// Dropdown 도 추가
				if (_types[col - 1].Contains("Enum::"))
				{
					string enumName = _types[col - 1].Split("Enum::")[1];
					AddEnumDropdown(_infoWorksheet, enumName, row, col);
				}
				else if (_types[col - 1].Contains("Boolean::"))
				{
					AddBooleanDropdown(_infoWorksheet, row, col);
				}
			}
		}

		private void TypeParser(Type type, bool isClass = false)
		{
			Type? tableDataType = type;
			if (null == tableDataType)
			{
				return;
			}

			foreach (var property in tableDataType.GetProperties())
			{
				if (null == property
					|| null == property.PropertyType)
				{
					return;
				}

				Type propertyType = property.PropertyType;
				string typeName = getStringType(propertyType);
				if (true == isClass)
				{
					typeName = $"{type.FullName}.{typeName}";
				}

				_types.Add(typeName);
				_names.Add(property.Name);

				// 일단 클래스 타입이 아니라면 continue
				if (false == propertyType.IsClass)
				{
					continue;
				}

				// String 이면 continue
				if (typeof(string) == propertyType)
				{
					continue;
				}

				// String 이 아닌 클래스 타입
				Type classType = propertyType;
				if (true == propertyType.IsGenericType)
				{
					classType = classType.GenericTypeArguments[0];
				}
				TypeParser(classType, true);
			}
		}

		private string getStringType(Type type)
		{
			if (true == type.IsEnum)
			{
				return $"Enum::{type.FullName}";
			}

			if (typeof(string) == type)
			{
				return "String";
			}

			if (true == type.IsValueType)
			{
				// Single 은 Floot 인데..
				if (typeof(float) == type)
				{
					return "Floot";
				}
				return type.Name;
			}

			//클래스
			if (true == type.IsClass)
			{
				if (type.Name.StartsWith("List"))
				{
					string className = type.GetGenericArguments()[0].FullName ?? string.Empty;
					return $"List<Class::{className}>";
				}
				return $"Class::{type.FullName}";
			}
			return "null";
		}

		private void AddEnumPage(ExcelPackage excel, Type enumType)
		{
			// 워크시트 데이터 변경
			string enumName = enumType.FullName ?? string.Empty;
			var worksheet = excel.Workbook.Worksheets.Add($"{enumName}");

			worksheet.Cells[1, 1].Value = "Name";
			worksheet.Cells[1, 2].Value = "Value";
			worksheet.Cells[1, 3].Value = "";

			int row = 2;
			foreach (var value in enumType.GetEnumValues())
			{
				worksheet.Cells[row, 1].Value = value;
				worksheet.Cells[row, 2].Value = (short)value;

				if (TableOption.UseExportXMLParse)
				{
					worksheet.Cells[row, 3].Value = _xmlCommentReader?.GetMemberComment(eMemberType.FILED, enumName, value.ToString());
				}

				row++;
			}
			// 열 너비 자동 조정
			worksheet.Column(1).AutoFit();
			worksheet.Column(2).AutoFit();
			worksheet.Column(3).AutoFit();
		}

		private void AddEnumDropdown(ExcelWorksheet worksheet, string enumName, int row, int col)
		{
			int enumRow = 0;
			{
				var enumWorksheet = _excel.Workbook.Worksheets[enumName];
				enumRow = enumWorksheet.Dimension.Rows;
			}

			var cellAddress = worksheet.Cells[row, col].Address;
			var existingValidations = worksheet.DataValidations.Where(v => v.Address.Address == cellAddress);

			bool dropdownExists = existingValidations.Any(v => v is ExcelDataValidationList);

			if (false == dropdownExists)
			{
				// 첫 번째 워크시트에 드롭다운 목록 생성
				var validation = worksheet.DataValidations.AddListValidation(cellAddress);

				// 참조 범위를 동적으로 설정
				string excelAddress = new ExcelAddress(2, 1, enumRow, 1).AddressAbsolute; // "Name" 열의 데이터 범위
				validation.Formula.ExcelFormula = $"{enumName}!{excelAddress}";
			}
		}

		private void AddBooleanDropdown(ExcelWorksheet worksheet, int row, int col)
		{
			var cellAddress = worksheet.Cells[row, col].Address;
			var existingValidations = worksheet.DataValidations.Where(v => v.Address.Address == cellAddress);

			bool dropdownExists = existingValidations.Any(v => v is ExcelDataValidationList);

			if (false == dropdownExists)
			{
				// 첫 번째 워크시트에 드롭다운 목록 생성
				var validation = worksheet.DataValidations.AddListValidation(cellAddress);

				// 드롭다운 목록에 "true"와 "false"만 추가
				validation.Formula.Values.Add("TRUE");
				validation.Formula.Values.Add("FALSE");
			}
		}


		private void SetCellData(ExcelWorksheet worksheet, int row, int col, object value, ExcelHorizontalAlignment horizontalAlignment, Color backgroundColor)
		{
			var cell = worksheet.Cells[row, col];
			cell.Value = value;
			cell.Style.HorizontalAlignment = horizontalAlignment;
			cell.Style.Font.Bold = true;
			cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
			cell.Style.Fill.BackgroundColor.SetColor(backgroundColor);

			cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
		}
		private void SetCellData(ExcelWorksheet worksheet, int row, int col, object value, ExcelHorizontalAlignment horizontalAlignment, Color backgroundColor, Color fontColor)
		{
			var cell = worksheet.Cells[row, col];
			cell.Value = value;
			cell.Style.HorizontalAlignment = horizontalAlignment;
			cell.Style.Font.Bold = true;
			cell.Style.Font.Color.SetColor(fontColor);
			cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
			cell.Style.Fill.BackgroundColor.SetColor(backgroundColor);

			cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
		}
	}
}
