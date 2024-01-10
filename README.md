# RedsunExcelDataLoader

기본적인 사용은 새로운 Console 프로젝트를 만들어, Dll 파일을 참조로 넣는 방법과 혹은 프로젝트 참조를 하여 Console 프로젝트를 만드는 것이다.
(주의: .net8.0 이상 지원)

```csharp
namespace GameDataTable
{
	internal class Program
	{
		static void Main(string[] args)
		{
			TableOption.ExcelPath = @$"..\..\..\..\..\GameTable";
			TableOption.OutputPath = @$"..\..\..\..\..\GameTable";

			var tableLoader = new TableLoader<Table.GameData>();
			tableLoader.Run();
		}
	}
}
```



기본적으로 Memorypack 이 포함되어 있으며,  Serailizer / Deserializer 를 각각 수동으로 등록할 수 있도록 되어있다.

```csharp
namespace GameDataTable
{
	internal class Program
	{
		static void Main(string[] args)
		{
			TableOption.ExcelPath = @$"..\..\..\..\..\GameTable";
			TableOption.OutputPath = @$"..\..\..\..\..\GameTable";
	
            // 직접만든 Serializer
			var tableLoader = new TableLoader<Table.GameData>(
				PacketSerializer.TableSerialize<Table.GameData>
				, PacketSerializer.Deserialize<Table.GameData>);

			tableLoader.Run();
		}
	}
}
```



Serializer 를 등록하려고 할 때, 아래와 같은 `Delegate` 규격을 지켜주면 된다.

```csharp
public delegate byte[] DelegateSerializer(T data);
public delegate T DelegateDeserializer(byte[] data);
```



기본적으로 아무 Serializer 를 넣지 않았을 때, 기본적으로 JsonSerializer 가 동작되도록 설계되어 있으며, 내장 MemoryPack Serializer 는 다음과 같다.

```csharp
DataTableSerializer.Serialize<Table.GameData>
DataTableSerializer.Deserialize<Table.GameData>
```