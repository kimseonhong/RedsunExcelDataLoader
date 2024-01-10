using MemoryPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace GameDataTableLoader.Serializer
{
	/// <summary>
	/// 해당 Class 는 MemoryPack 을 기반으로 Serializer 를 진행한다.
	/// </summary>
	public class DataTableSerializer
	{
		public static byte[] Serialize<T>(T memoryPackData) where T : IMemoryPackable<T> => Serialize(memoryPackData, eCompressionType.GZip);

		public static byte[] Serialize<T>(T memoryPackData, eCompressionType compressionType)
			where T : IMemoryPackable<T>
		{
			var data = MemoryPackSerializer.Serialize(memoryPackData);
			data = Compression.Compress(data, compressionType);
			return data;
		}

		public static T Deserialize<T>(byte[] binaryData)
			where T : IMemoryPackable<T>
		{
			var compressionType = Compression.DetectCompressionType(binaryData);
			if (compressionType == eCompressionType.NONE)
			{
				throw new Exception();
			}

			var data = Compression.Decompress(binaryData, compressionType);
			var result = MemoryPackSerializer.Deserialize<T>(data);

			if (result == null)
			{
				throw new Exception();
			}

			return result;
		}

		public static byte[] SerializeToUtf8Binary<T>(T data)
		{
			return JsonSerializer.SerializeToUtf8Bytes(data);
		}

		public static T DeserializeToJsonSerializer<T>(byte[] data)
		{
			return JsonSerializer.Deserialize<T>(data);
		}
	}
}
