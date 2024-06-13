////////////////////////////////////////////////////////////////////////////////
//
// SpcDocumentEvent  - ���������� ������� �� ��������� ������������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcDocumentEvent : BaseEvent, ksSpcDocumentNotify
	{
		public SpcDocumentEvent(object obj, object doc)
			: base(obj, typeof(ksSpcDocumentNotify).GUID, doc) {}

		// sdDocumentBeginAdd - ������ ���������� ��������� ���������� �������
		public bool DocumentBeginAdd()
		{
			return true;
		}


		// sdDocumentAdd - ���������� ��������� ���������� �������
		public bool DocumentAdd(string docName)
		{
			return true;
		}
      

		// sdDocumentBeginRemove - ������ �������� ��������� ���������� �������
		public bool DocumentBeginRemove(string docName)
		{
			return true;
		}


		// sdDocumentRemove - �������� ��������� ���������� �������
		public bool DocumentRemove(string docName)
		{
			return true;
		}


		// sdSpcStyleBeginChange - ������ ��������� ����� ������������
		public bool SpcStyleBeginChange(string libName, int numb)
		{
			return true;
		}


		// sdSpcStyleChange - ����� ������������ ���������
		public bool SpcStyleChange(string libName, int numb)
		{
			Global.UpdateSlideBox(null);
			return true;
		}	


		// ������� ���������� ������� ��������� c�����������
		public static BaseEvent NewSpcDocEvent(object doc, int objType)
		{
			SpcDocumentEvent res = null;
			if (doc != null)
			{
				if (!BaseEvent.FindEvent(typeof(SpcDocumentEvent), doc)) 
				{
					ksSpcDocument spcDoc = (ksSpcDocument)doc;
					if (spcDoc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// ����� ��� �������� ��������� ����������
						res = new SpcDocumentEvent(spcDoc.GetSpcDocumentNotify(), doc);
						res.Advise();
					}
				}
			}
			return res;
		}
	}
}