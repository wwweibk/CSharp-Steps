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
		public SpcDocumentEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksSpcDocumentNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// sdDocumentBeginAdd - ������ ���������� ��������� ���������� �������
		public bool DocumentBeginAdd()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = m_LibName + " --> DocumentBeginAdd";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// sdDocumentAdd - ���������� ��������� ���������� �������
		public bool DocumentAdd(string docName)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = m_LibName + " --> DocumentAdd\ndocName = " + docName;
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}
      

		// sdDocumentBeginRemove - ������ �������� ��������� ���������� �������
		public bool DocumentBeginRemove(string docName)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = m_LibName + " --> DocumentBeginRemove\ndocName =" + docName;
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// sdDocumentRemove - �������� ��������� ���������� �������
		public bool DocumentRemove(string docName)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = m_LibName + " --> DocumentRemove\ndocName =" + docName;
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// sdSpcStyleBeginChange - ������ ��������� ����� ������������
		public bool SpcStyleBeginChange(string libName, int numb)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcStyleBeginChange\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// sdSpcStyleChange - ����� ������������ ���������
		public bool SpcStyleChange(string libName, int numb)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcStyleChange\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}	
	}
}
