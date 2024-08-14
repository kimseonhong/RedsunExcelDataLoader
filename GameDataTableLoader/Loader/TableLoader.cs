using GameDataTableLoader.Serializer;
using OfficeOpenXml;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text.Json;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Runtime.InteropServices.Marshalling.IIUnknownCacheStrategy;

namespace GameDataTableLoader.Loader
{
	public class TableLoader<T>
		where T : class, new()
	{
		public delegate byte[] DelegatePacker(T data);
		public DelegatePacker? Packer;

		public delegate T DelegateUnPacker(byte[] data);
		public DelegateUnPacker? UnPacker;

		private T _data = new();
		private Dictionary<string /* TableName */, TableInfo> _tableInfos = new();

		private Type _tableDataType;
		private Dictionary<string /* TableName */, DataType> _properties = new();

		public Dictionary<string /* TableName */, TableInfo> TableInfos() { return _tableInfos; }

		public TableLoader(DelegatePacker? packer = null, DelegateUnPacker? unpacker = null)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			_tableDataType = _data.GetType();
			_properties = DataType.PropertyParse(_tableDataType);

			Packer = packer;
			if (null == packer)
			{
				Packer = DataTableSerializer.SerializeToUtf8Binary<T>;
			}

			UnPacker = unpacker;
			if (null == unpacker)
			{
				UnPacker = DataTableSerializer.DeserializeToJsonSerializer<T>;
			}

			Directory.CreateDirectory(@$"{TableOption.OutputPath}");
			Directory.CreateDirectory(@$"{TableOption.OutputPath}\_DataPack");
		}

		public void Run(bool useReflection = true)
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(TableOption.ExcelPath);
			AllFiles(directoryInfo);

			if (useReflection)
			{
				_RunReflectionMode();
			}
			else
			{
				_RunJsonMode();
			}
		}

		private void _RunReflectionMode()
		{
			foreach (var table in _tableInfos)
			{
				PropertyInfo? propertyInfo = _tableDataType.GetProperty($"{table.Key}s");
				if (null == propertyInfo)
				{
					Console.WriteLine($"{table.Key} No Property");
					continue;
				}

				var data = TableOption.TableJsonSerializer(table.Value.GetData());
				var realData = TableOption.TableJsonDeserializer(data, propertyInfo.PropertyType);
				propertyInfo.SetValue(_data, realData);
			}
		}

		private void _RunJsonMode()
		{
			Dictionary<string /* ObjectName */, dynamic> rawData = new Dictionary<string, dynamic>();
			foreach (var table in _tableInfos)
			{
				rawData.Add($"{table.Key}s", table.Value.GetData());
			}

			var data = TableOption.TableJsonSerializer(rawData);
			var realData = TableOption.TableJsonDeserializer(data, _data.GetType());
			_data = realData as T ?? new T();
		}

		public void SavePack() => Packing().Wait();

		private void AllFiles(DirectoryInfo directoryInfo)
		{
			// 현재 디렉토리의 모든 파일을 출력
			foreach (var file in directoryInfo.GetFiles())
			{
				string fileName = file.Name;
				string tableName = $"{fileName.Split(".")[0]}";

				// 프로퍼티에 존재하지 않으면 패스
				DataType? dataType;
				if (false == _properties.TryGetValue(tableName, out dataType))
				{
					continue;
				}

				TableInfo? table;
				if (false == _tableInfos.TryGetValue(tableName, out table))
				{
					table = new TableInfo(dataType, tableName);
					_tableInfos.Add(tableName, table);
				}

				Console.WriteLine($"Run TableParse : {tableName} / {file.FullName}");
				table.Run(file);
				Console.WriteLine($"End TableParse : {tableName} / {file.FullName}");
			}

			// 현재 디렉토리의 모든 하위 디렉토리를 가져옴
			foreach (var directory in directoryInfo.GetDirectories())
			{
				// 각 하위 디렉토리에 대해 재귀적으로 함수 호출
				AllFiles(directory);
			}
		}

		private async Task Packing()
		{
			if (null == Packer)
			{
				throw new Exception("Serializer is NULL!!!!!!");
			}

			Console.WriteLine("Start Packing...");
			var time = DateTime.Now.ToString("yyyyMMdd_HHmmss");
			var binary = Packer.Invoke(_data);
			var binaryArray = binary.ToArray();
			Console.WriteLine("Finish Packing...");

			File.Delete(@$"{TableOption.OutputPath}\GameTable.pack");

			Console.WriteLine("Save...");
			await Task.WhenAll(
				File.WriteAllBytesAsync(@$"{TableOption.OutputPath}\GameTable.pack", binaryArray),
				File.WriteAllBytesAsync(@$"{TableOption.OutputPath}\_DataPack\GameTable_{time}.pack", binaryArray)
			);
			Console.WriteLine("Finish...");
		}
	}
}
