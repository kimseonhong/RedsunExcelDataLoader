using System.Xml;

namespace GameDataTableLoader.Exporter
{
	public enum eMemberType
	{
		_NONE,

		TYPE, // 타입 (클래스, 인터페이스, 구조체, 열거형, 대리자)
		METHOD, // 메서드
		PROPERTY, // 프로퍼티 (get;set;)
		FILED, // 필드
		EVENT, //이벤트 

		_END
	}

	public class XmlCommentReader
	{
		private XmlDocument xmlDocument;

		public XmlCommentReader(string xmlPath)
		{
			xmlDocument = new XmlDocument();
			xmlDocument.Load(xmlPath);
		}

		public string GetMemberComment(eMemberType memberType, string typeName, string memberName)
		{
			string type = "F";

			switch (memberType)
			{
				case eMemberType.TYPE:
					type = "T";
					break;
				case eMemberType.METHOD:
					type = "M";
					break;
				case eMemberType.PROPERTY:
					type = "P";
					break;
				case eMemberType.FILED:
					type = "F";
					break;
				case eMemberType.EVENT:
					type = "E";
					break;
				default:
					return string.Empty;
			}

			string xpath = $"/doc/members/member[@name='{type}:{typeName}.{memberName}']/summary";
			XmlNode? node = xmlDocument.SelectSingleNode(xpath);
			return node?.InnerText.Trim() ?? string.Empty;
		}
	}

}
