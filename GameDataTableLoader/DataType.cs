namespace GameDataTableLoader
{
	public class DataType
	{
		public Type Type { get; set; }
		public string Name { get; set; }
		public string FullName { get; set; }
		public string AssemblyName { get; set; }

		public Type FindPropertyType(string propertyName)
		{
			foreach (var property in Type.GetProperties())
			{
				if (property.PropertyType.FullName == propertyName)
				{
					return property.PropertyType;
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

				Type propertyType = property.PropertyType;
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
