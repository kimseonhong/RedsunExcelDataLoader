using Google.Protobuf;
using Google.Protobuf.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;

namespace GameDataTableLoader.Exporter
{
	public class ExportParser
	{
		private DataType _dataType;

		private List<string> _types = new List<string>();
		private List<string> _names = new List<string>();
		private List<string> _realType = new List<string>();

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
					string enumValue = _xmlCommentReader?.GetMemberComment(eMemberType.PROPERTY, _realType[col - 1], _names[col - 1]) ?? string.Empty;
					//SetCellData(worksheet, 3, col, enumValue, ExcelHorizontalAlignment.Center, Color.FromArgb(51, 63, 79), Color.White);
					SetCellData(_infoWorksheet, 3, col, enumValue, ExcelHorizontalAlignment.Center, Color.FromArgb(217, 225, 242));
				}
				else
				{
					SetCellData(_infoWorksheet, 3, col, "주석", ExcelHorizontalAlignment.Center, Color.FromArgb(217, 225, 242));
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
			if (type == null) return;

			// ✅ Protobuf 메시지면 Descriptor 기반으로 처리
			if (typeof(IMessage).IsAssignableFrom(type))
			{
				TypeParserProtobuf(type);
				return;
			}

			// ---- 기존 로직 그대로 ----
			Type? tableDataType = type;
			if (null == tableDataType) return;

			string preName = "";
			foreach (var property in tableDataType.GetProperties())
			{
				if (property == null || property.PropertyType == null) return;

				Type propertyType = property.PropertyType;
				string typeName = getStringType(propertyType);

				// 구글 Protobuf 대응 (기존)
				if (string.IsNullOrEmpty(propertyType.Namespace) || propertyType.Namespace.Contains("Google.Protobuf"))
				{
					if (property.Name.Equals("Parser") || property.Name.Equals("Descriptor"))
						continue;
				}

				if (true == isClass)
					typeName = $"{type.FullName}.{typeName}";

				if (property.Name.Equals($"Has{preName}"))
					continue;

				preName = property.Name;
				_types.Add(typeName);
				_names.Add(property.Name);
				_realType.Add(property.DeclaringType?.FullName ?? string.Empty);

				if (false == propertyType.IsClass) continue;
				if (typeof(string) == propertyType) continue;

				Type classType = propertyType;
				if (true == propertyType.IsGenericType)
					classType = classType.GenericTypeArguments[0];

				TypeParser(classType, true);
			}
		}

		private void TypeParserProtobuf(Type messageType)
		{
			// static Descriptor 가져오기
			var descProp = messageType.GetProperty("Descriptor", BindingFlags.Public | BindingFlags.Static);
			if (descProp?.GetValue(null) is not MessageDescriptor md)
			{
				// Descriptor 못 얻으면 fallback: 기존 방식
				TypeParserFallback(messageType);
				return;
			}

			// oneof 처리용: 이미 출력한 oneof인지 체크
			var printedOneofs = new HashSet<string>();

			foreach (var field in md.Fields.InDeclarationOrder())
			{
				// oneof field면: oneof 마커 + 케이스 블럭으로 출력
				if (field.ContainingOneof != null)
				{
					string oneofName = field.ContainingOneof.Name; // proto oneof 이름(예: "param")

					if (!printedOneofs.Contains(oneofName))
					{
						printedOneofs.Add(oneofName);
						_types.Add($"OneOf::{oneofName}");
						_names.Add(oneofName);
						_realType.Add(messageType.FullName ?? string.Empty);
					}

					// 케이스는 보통 메시지 타입(Param_MeleeArc 같은)일 거라 가정
					if (field.FieldType != FieldType.Message)
					{
						// oneof가 primitive 케이스도 있을 수 있으니, 이 경우는 후보 1칸만 내보내는 정책으로 가능
						// (원하면 여기서 별도 포맷 정의)
						AddOneOfPrimitiveCase(oneofName, messageType, field);
						continue;
					}

					// C# 프로퍼티명(파스칼)로 찾아서 Name에 사용
					string csPropName = FindCSharpPropertyName(messageType, field);

					var caseMsgType = field.MessageType.ClrType;
					string caseFullName = caseMsgType.FullName ?? caseMsgType.Name;

					// 케이스 시작 칸
					_types.Add($"{oneofName}.Class::{caseFullName}");
					_names.Add(csPropName);
					_realType.Add(messageType.FullName ?? string.Empty);

					// 케이스 메시지 내부 필드들 펼치기 (기존 규칙과 동일하게 fullName prefix 형태)
					AddMessageFields(caseMsgType);

					continue;
				}

				// 일반 필드 출력
				AddNormalField(messageType, field);
			}
		}

		private void AddNormalField(Type ownerType, FieldDescriptor field)
		{
			string csPropName = FindCSharpPropertyName(ownerType, field);

			// repeated
			if (field.IsRepeated)
			{
				// List<Class::X> 형태로 맞추기
				if (field.FieldType == FieldType.Message)
				{
					string fn = field.MessageType.ClrType.FullName ?? field.MessageType.ClrType.Name;
					_types.Add($"List<Class::{fn}>");
					_names.Add(csPropName);
					_realType.Add(ownerType.FullName ?? string.Empty);

					// 리스트 내부 메시지는 “Class 파싱” 방식이니까, 템플릿은 메시지 필드를 펼쳐줘야 함
					AddMessageFields(field.MessageType.ClrType);
				}
				else
				{
					// repeated primitive는 현재 Parser가 직접 지원 안 하니까(지금은 List<Class::>만 파싱),
					// 여기서 정책 결정 필요.
					// 일단 막거나, List<Primitive> 포맷을 새로 정의해야 함.
				}
				return;
			}

			// message (non-oneof)
			if (field.FieldType == FieldType.Message)
			{
				string fn = field.MessageType.ClrType.FullName ?? field.MessageType.ClrType.Name;
				_types.Add($"Class::{fn}");
				_names.Add(csPropName);
				_realType.Add(ownerType.FullName ?? string.Empty);

				AddMessageFields(field.MessageType.ClrType);
				return;
			}

			// primitive/enum
			_types.Add(ProtoFieldToCellType(field));
			_names.Add(csPropName);
			_realType.Add(ownerType.FullName ?? string.Empty);
		}

		private void AddMessageFields(Type msgType)
		{
			var descProp = msgType.GetProperty("Descriptor", BindingFlags.Public | BindingFlags.Static);
			if (descProp?.GetValue(null) is not MessageDescriptor md) return;

			var printedOneofs = new HashSet<string>();

			foreach (var f in md.Fields.InDeclarationOrder())
			{
				// ✅ oneof 처리
				if (f.ContainingOneof != null)
				{
					string oneofName = f.ContainingOneof.Name;

					// oneof 마커 1회만 출력
					if (printedOneofs.Add(oneofName))
					{
						_types.Add($"{msgType.FullName}.OneOf::{oneofName}");
						_names.Add(oneofName);
						_realType.Add(msgType.FullName ?? string.Empty);
					}

					// 케이스 출력 (message / primitive)
					if (f.FieldType == FieldType.Message)
					{
						// 케이스 시작 칸
						var caseMsgType = f.MessageType.ClrType;
						string caseFullName = caseMsgType.FullName ?? caseMsgType.Name;

						// Name은 C# 프로퍼티명으로 맞추는 게 가장 안전 (MeleeArc, MeleeCircle)
						string csPropName = FindCSharpPropertyName(msgType, f);

						_types.Add($"{msgType.FullName}.{oneofName}.Class::{caseFullName}");
						_names.Add(csPropName);
						_realType.Add(msgType.FullName ?? string.Empty);

						// 케이스 메시지 내부 필드 펼치기
						AddMessageFields(caseMsgType);
					}
					else
					{
						// primitive/enum 케이스는 "후보 컬럼 1칸"만 출력
						string csPropName = FindCSharpPropertyName(msgType, f);
						string cellType = ProtoFieldToCellType(f);

						_types.Add($"{msgType.FullName}.{oneofName}.{cellType}");
						_names.Add(csPropName);
						_realType.Add(msgType.FullName ?? string.Empty);
					}

					continue;
				}

				// ✅ 일반 필드 처리 (oneof 아님)
				if (f.IsRepeated)
				{
					// repeated message만 지원 (너 Loader가 List<Class::>에 최적화)
					if (f.FieldType == FieldType.Message)
					{
						var elemType = f.MessageType.ClrType;
						string elemFullName = elemType.FullName ?? elemType.Name;

						_types.Add($"{msgType.FullName}.List<Class::{elemFullName}>");
						_names.Add(FindCSharpPropertyName(msgType, f));
						_realType.Add(msgType.FullName ?? string.Empty);

						AddMessageFields(elemType);
					}
					else
					{
						// repeated primitive는 포맷 정의 필요 (원하면 추가)
						// 일단 스킵하거나 예외로 막는 게 안전
						// throw new Exception($"Repeated primitive is not supported: {msgType.FullName}.{f.Name}");
					}
					continue;
				}

				if (f.FieldType == FieldType.Message)
				{
					var childType = f.MessageType.ClrType;
					string childFullName = childType.FullName ?? childType.Name;

					_types.Add($"{msgType.FullName}.Class::{childFullName}");
					_names.Add(FindCSharpPropertyName(msgType, f));
					_realType.Add(msgType.FullName ?? string.Empty);

					AddMessageFields(childType);
					continue;
				}

				// primitive/enum
				{
					string cellType = ProtoFieldToCellType(f);
					_types.Add($"{msgType.FullName}.{cellType}");
					_names.Add(FindCSharpPropertyName(msgType, f));
					_realType.Add(msgType.FullName ?? string.Empty);
				}
			}
		}


		private string ProtoFieldToCellType(FieldDescriptor f)
		{
			// enum
			if (f.FieldType == FieldType.Enum)
				return $"Enum::{f.EnumType.ClrType.FullName}";

			return f.FieldType switch
			{
				FieldType.Int32 => "Int32",
				FieldType.Int64 => "Int64",
				FieldType.UInt32 => "UInt32",
				FieldType.UInt64 => "UInt64",
				FieldType.SInt32 => "Int32",
				FieldType.SInt64 => "Int64",
				FieldType.Fixed32 => "UInt32",
				FieldType.Fixed64 => "UInt64",
				FieldType.SFixed32 => "Int32",
				FieldType.SFixed64 => "Int64",
				FieldType.Bool => "Boolean",
				FieldType.String => "String",
				FieldType.Float => "Floot",
				FieldType.Double => "Double",
				// bytes는 현재 포맷 정의 필요
				_ => "String"
			};
		}

		private string FindCSharpPropertyName(Type ownerType, FieldDescriptor field)
		{
			// proto name(param_melee_arc) -> C# property(ParamMeleeArc)로 매핑
			// 가장 안전한 건 Reflection으로 실제 존재하는 프로퍼티를 찾는 것.
			// Google.Protobuf는 보통 PascalCase 규칙이라 아래 정도면 잘 맞음.
			string pascal = ToPascal(field.Name);

			// 실제 프로퍼티 있으면 그걸 사용
			var pi = ownerType.GetProperty(pascal, BindingFlags.Public | BindingFlags.Instance);
			return pi?.Name ?? pascal;
		}

		private string ToPascal(string protoName)
		{
			// param_melee_arc -> ParamMeleeArc
			var parts = protoName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
			return string.Concat(parts.Select(p => char.ToUpperInvariant(p[0]) + p.Substring(1)));
		}

		private void AddOneOfPrimitiveCase(string oneofName, Type messageType, FieldDescriptor field)
		{
			// 1) 엑셀에 찍힐 "Name"(2행): C# 프로퍼티명으로 맞추기
			//    ex) "int_value" -> "IntValue"
			string csPropName = FindCSharpPropertyName(messageType, field);

			// 2) 엑셀에 찍힐 "Type"(1행): oneofPrefix + primitive type
			//    ex) param.Int32 / param.String / param.Enum::FullName
			string cellType = ProtoFieldToCellType(field);
			string excelType = $"{oneofName}.{cellType}";

			_types.Add(excelType);
			_names.Add(csPropName);
			_realType.Add(messageType.FullName ?? string.Empty);
		}



		private void TypeParserFallback(Type type, bool isClass = false)
		{
			Type? tableDataType = type;
			if (null == tableDataType)
			{
				return;
			}

			string preName = "";
			foreach (var property in tableDataType.GetProperties())
			{
				if (null == property
					|| null == property.PropertyType)
				{
					return;
				}

				Type propertyType = property.PropertyType;
				string typeName = getStringType(propertyType);

				// 구글 Protobuf 대응
				if (string.IsNullOrEmpty(propertyType.Namespace) || propertyType.Namespace.Contains("Google.Protobuf"))
				{
					if (property.Name.Equals("Parser") || property.Name.Equals("Descriptor"))
					{
						continue;
					}
				}

				if (true == isClass)
				{
					typeName = $"{type.FullName}.{typeName}";
				}

				// 구글 Protobuf 대응
				if (property.Name.Equals($"Has{preName}"))
				{
					continue;
				}

				preName = property.Name;
				_types.Add(typeName);
				_names.Add(property.Name);
				_realType.Add(property.DeclaringType?.FullName ?? string.Empty);

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
				TypeParserFallback(classType, true);
			}
		}

		private string getStringType(Type type)
		{
			// protobuf message면 class로 고정
			if (typeof(IMessage).IsAssignableFrom(type))
				return $"Class::{type.FullName}";

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
				if (type.Name.StartsWith("List")
					|| type.Name.StartsWith("Repeat"))
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
				worksheet.Cells[row, 2].Value = (int)value;

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
