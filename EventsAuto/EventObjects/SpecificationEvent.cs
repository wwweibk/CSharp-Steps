////////////////////////////////////////////////////////////////////////////////
//
// SpecificationEvent  - ���������� ������� �� ������������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpecificationEvent : BaseEvent, ksSpecificationNotify
	{
		public SpecificationEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksSpecificationNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// ssTuningSpcStyleBeginChange - ������ ��������� �������� ������������
		public bool TuningSpcStyleBeginChange(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> TuningSpcStyleBeginChange\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}


		// ssTuningSpcStyleChange - ���������� ������������ ����������
		public bool TuningSpcStyleChange(string libName, int numb, bool isOk)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> TuningSpcStyleChange\nlibName = {1}\nnumb = {2}\nisOk = {3}", m_LibName, libName, numb, isOk ? "true" : "false");
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;      
		}


		// ssChangeCurrentSpcDescription - ���������� ������� �������� ������������
		public bool ChangeCurrentSpcDescription(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChangeCurrentSpcDescription\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// ssSpcDescriptionAdd - ���������� �������� ������������
		public bool SpcDescriptionAdd(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionAdd\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssSpcDescriptionRemove - ��������� �������� ������������
		public bool SpcDescriptionRemove(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionRemove\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssSpcDescriptionBeginEdit - ������ �������������� �������� ������������
		public bool SpcDescriptionBeginEdit(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionBeginEdit\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}


		// ssSpcDescriptionEdit - ��������������� �������� ������������
		public bool SpcDescriptionEdit(string libName, int numb, bool isOk)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionEdit\nlibName = {1}\nnumb = {2}\nisOk = {3}", m_LibName, libName, numb, isOk ? "true" : "false");
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// ssSynchronizationBegin - ������ �������������
		public bool SynchronizationBegin()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> SynchronizationBegin";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// ssSynchronization - ������������� ���������
		public bool Synchronization()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> Synchronization";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssBeginCalcPositions - ������  ������� �������
		public bool BeginCalcPositions()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> BeginCalcPositions";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// ssCalcPositions - �������� ������ ������� 
		public bool CalcPositions()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> CalcPositions";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssBeginCreateObject - ������ �������� ������� ������������ (�� ������� ������ �������) 
		public bool BeginCreateObject(int typeObj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginCreateObject\ntypeObj = {1}", m_LibName, typeObj);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}
	}
}
