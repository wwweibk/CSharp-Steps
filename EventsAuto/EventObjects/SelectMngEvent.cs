////////////////////////////////////////////////////////////////////////////////
//
// SelectMngEvent  - ���������� ������� �� ��������� �������������� ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SelectMngEvent : BaseEvent, ksSelectionMngNotify
	{
		public SelectMngEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksSelectionMngNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// ksmSelect - ������ ������������
		public bool Select(object obj) 
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSelectEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> Select\nobj = {0}", m_LibName, obj);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ksmUnselect - ������ ���������������
		public bool Unselect(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSelectEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> Unselect\nobj = {1}", m_LibName, obj);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ksmUnselectAll - ��� ������� ����������������
		public bool UnselectAll() 
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSelectEvents.Checked)
			{
				string str = m_LibName + " --> UnselectAll";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}	
	}
}
