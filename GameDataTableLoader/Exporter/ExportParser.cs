using OfficeOpenXml;

namespace GameDataTableLoader.Exporter
{
	public class ExportParser
	{
		private DataType _dataType;

		private ExcelWorksheet worksheet = default;
		private List<string> types = new List<string>();
		private List<string> names = new List<string>();

		public ExportParser(DataType dataType)
		{
			_dataType = dataType;
		}

		public void Run()
		{
			using (var excel = new ExcelPackage())
			{
				worksheet = excel.Workbook.Worksheets.Add($"{_dataType.Name}");
				TypeParser(_dataType.Type);

				for (int col = 1; col <= types.Count; col++)
				{
					worksheet.Cells[1, col].Value = types[col - 1];
					worksheet.Cells[1, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
					worksheet.Cells[2, col].Value = names[col - 1];
					worksheet.Cells[2, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

					// 열 너비 자동 조정
					worksheet.Column(col).AutoFit();
				}

				// Enum Table Check
				for (int i = 0; i < types.Count; i++)
				{
					string type = types[i];
					if (false == type.StartsWith("Enum::"))
					{
						continue;
					}

					string enumName = type.Replace("Enum::", string.Empty);
					Type enumType = _dataType.FindPropertyType(enumName);
					EnumLink(excel, enumType, i);
				}


				// 파일 저장
				string tableName = _dataType.Name;
				var fileInfo = new FileInfo(@$"{TableOption.OutputPath}\_tmpl\{tableName}_tmpl.xlsx");
				excel.SaveAs(fileInfo);
			}
		}

		public void TypeParser(Type type, bool isClass = false)
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

				types.Add(typeName);
				names.Add(property.Name);

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

		public string getStringType(Type type)
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

		private void EnumLink(ExcelPackage excel, Type enumType, int index)
		{
			// 워크시트 데이터 변경
			string enumName = enumType.FullName;
			worksheet = excel.Workbook.Worksheets.Add($"{enumName}");

			worksheet.Cells[1, 1].Value = "Name";
			worksheet.Cells[1, 2].Value = "Value";

			int row = 2;
			foreach (var value in enumType.GetEnumValues())
			{
				worksheet.Cells[row, 1].Value = value;
				worksheet.Cells[row, 2].Value = (short)value;
				row++;
			}
			// 열 너비 자동 조정
			worksheet.Column(1).AutoFit();
			worksheet.Column(2).AutoFit();

			worksheet = excel.Workbook.Worksheets[0];

			// 첫 번째 워크시트에 드롭다운 목록 생성
			var validation = worksheet.DataValidations.AddListValidation(worksheet.Cells[3, index + 1].Address);

			// 참조 범위를 동적으로 설정
			string excelAddress = new ExcelAddress(2, 1, row - 1, 1).Address; // "Name" 열의 데이터 범위
			validation.Formula.ExcelFormula = $"{enumName}!{excelAddress}";

		}
	}
}
