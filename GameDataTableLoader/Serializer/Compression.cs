using ICSharpCode.SharpZipLib.BZip2;
using ICSharpCode.SharpZipLib.GZip;
using K4os.Compression.LZ4;
using K4os.Compression.LZ4.Streams;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GameDataTableLoader.Serializer
{
	public enum eCompressionType
	{
		NONE = 0,
		BZip2 = 1,
		GZip = 2,
		LZ4Stream = 3,
	}

	public class CompressionException : Exception
	{
		public eCompressionType CompressionType = eCompressionType.NONE;
		public CompressionException() { }
		public CompressionException(string message) : base(message) { }
		public CompressionException(string message, eCompressionType eCompressionType) : base(message)
		{
			CompressionType = eCompressionType;
		}
	}

	public class Compression
	{
		public static eCompressionType DetectCompressionType(byte[] data)
		{
			if (data.Length < 2)
			{
				return eCompressionType.NONE;
			}

			// BZip2 의 매직넘버
			if (data[0] == 0x42
				&& data[1] == 0x5A)
			{
				return eCompressionType.BZip2;
			}

			// GZip2 의 매직넘버
			if (data[0] == 0x1F
				&& data[1] == 0x8B)
			{
				return eCompressionType.GZip;
			}

			// LZ4 의 매직넘버, 단 Stream 과 Codec 에는 매직넘버 차이가 없음
			if (data[0] == 0x04
				&& data[1] == 0x22
				&& data[2] == 0x4D
				&& data[3] == 0x18)
			{
				return eCompressionType.LZ4Stream;
			}

			return eCompressionType.NONE;
		}

		public static byte[] Compress(byte[] data, eCompressionType type)
		{
			try
			{
				switch (type)
				{
					case eCompressionType.BZip2:
						return _BZip2(data, 0, data.Length);
					case eCompressionType.GZip:
						return _GZip(data, 0, data.Length);
					case eCompressionType.LZ4Stream:
						return _LZ4StreamEncode(data, 0, data.Length);
					default:
						throw new CompressionException("Compression Type is None. Plz Checked!");
				}
			}
			catch
			{
				throw;
			}
		}

		public static byte[] Decompress(byte[] data, eCompressionType type)
		{
			try
			{
				switch (type)
				{
					case eCompressionType.BZip2:
						return _UnBZip2(data, 0, data.Length);
					case eCompressionType.GZip:
						return _UnGZip(data, 0, data.Length);
					case eCompressionType.LZ4Stream:
						return _LZ4StreamDecode(data, 0, data.Length);
					default:
						throw new CompressionException("Decompression Type is None. Plz Checked!");
				}
			}
			catch
			{
				throw;
			}
		}


		private static byte[] _BZip2(byte[] data, int offset, int length)
		{
			using (var inStream = new MemoryStream(data, offset, length))
			{
				using (var outStream = new MemoryStream())
				{
					BZip2.Compress(inStream, outStream, false, 3);
					return outStream.ToArray();
				}
			}
		}

		private static byte[] _UnBZip2(byte[] data, int offset, int length)
		{
			using (var inStream = new MemoryStream(data, offset, length))
			{
				using (var outStream = new MemoryStream())
				{
					BZip2.Decompress(inStream, outStream, false);
					return outStream.ToArray();
				}
			}
		}

		private static byte[] _GZip(byte[] data, int offset, int size)
		{
			using (var ms = new MemoryStream())
			{
				using (var gzip = new GZipOutputStream(ms))
				{
					gzip.Write(data, offset, size);
					gzip.Flush();
					gzip.Finish();

					byte[] buffer = new byte[ms.Length];
					ms.Seek(0, SeekOrigin.Begin);
					ms.Read(buffer, 0, buffer.Length);
					return buffer;
				}
			}
		}

		private static byte[] _UnGZip(byte[] data, int offset, int size)
		{
			using (var ms = new MemoryStream(data, offset, size))
			{
				using (var gzip = new GZipInputStream(ms))
				{
					using (var outputStream = new MemoryStream())
					{
						gzip.CopyTo(outputStream);
						return outputStream.ToArray();
					}
				}
			}
		}

		// LZ4Codec 과 LZ4Stream 은 동일하나 Codec 이 더 빠르긴함
		// 왜냐면 OriginalSize 가 이미 변수로 전달되어 사용하기 때문...
		private static byte[] _LZ4CodecEncode(byte[] data, int offset, int size)
		{
			var target = new byte[LZ4Codec.MaximumOutputSize(size)];
			var encoded = LZ4Codec.Encode(data, offset, size
											, target, 0, target.Length);

			if (encoded == -1)
			{
				return data;
			}
			return target.AsSpan().Slice(0, encoded).ToArray();
		}

		private static byte[] _LZ4CodecDecode(byte[] data, int offset, int size, int originalSize)
		{
			var target = new byte[originalSize];
			var decoded = LZ4Codec.Decode(data, offset, size
											, target, 0, target.Length);

			if (decoded == -1)
			{
				return new byte[0];
			}
			return target;
		}

		private static byte[] _LZ4StreamEncode(byte[] data, int offset, int size)
		{
			using (var input = new MemoryStream())
			{
				using (var lz4Stream = LZ4Stream.Encode(input, LZ4Level.L06_HC))
				{
					lz4Stream.Write(data, offset, size);
				}
				return input.ToArray();
			}
		}

		private static byte[] _LZ4StreamDecode(byte[] data, int offset, int size)
		{
			using (var output = new MemoryStream())
			{
				using (var lz4Stream = LZ4Stream.Decode(new MemoryStream(data, offset, size)))
				{
					lz4Stream.CopyTo(output);
				}
				return output.ToArray();
			}
		}
	}
}
