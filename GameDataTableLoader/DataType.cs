namespace GameDataTableLoader
{
	public class DataType
	{
		public Type Type { get; set; }
		public string Name { get; set; }
		public string FullName { get; set; }
		public string AssemblyName { get; set; }

		public Type FindPropertyType(string propertyName) => FindPropertyType(Type, propertyName);
		public Type FindPropertyType(Type type, string propertyName)
		{
			foreach (var property in type.GetProperties())
			{
				if (property.PropertyType.FullName == propertyName)
				{
					return property.PropertyType;
				}

				if (true == property.PropertyType.IsGenericType)
				{
					var data = FindPropertyType(property.PropertyType.GenericTypeArguments[0], propertyName);
					if (data != null)
					{
						return data;
					}
				}
				else if (true == property.PropertyType.IsClass
					&& typeof(String) != property.PropertyType)
				{
					var data = FindPropertyType(property.PropertyType, propertyName);
					if (data != null)
					{
						return data;
					}
				}
			}
			return null;
		}

		public static Dictionary<string /* TableName */, DataType> PropertyParse(Type type)
		{
			var properies = new Dictionary<string, DataType>();
			foreach (var property in type.GetProperties())
			{
				if (null == property
					|| null == property.PropertyType)
				{
					continue;
				}

				// 구글 Protobuf 면 아래는 제거됨.
				Type propertyType = property.PropertyType;
				if (string.IsNullOrEmpty(propertyType.Namespace) || propertyType.Namespace.Contains("Google.Protobuf"))
				{
					if (property.Name.Equals("Parser") || property.Name.Equals("Descriptor"))
					{
						continue;
					}
				}

				if (false == propertyType.IsGenericType)
				{
					throw new Exception("무조건 List만을 보유한 객체여야만 합니다..");
				}

				propertyType = property.PropertyType.GenericTypeArguments[0];
				properies.Add(propertyType.Name, new DataType()
				{
					Type = propertyType,
					Name = propertyType.Name,
					FullName = propertyType.FullName ?? string.Empty,
					AssemblyName = propertyType.Assembly.GetName().Name ?? string.Empty
				});
			}
			return properies;
		}
	}
}
