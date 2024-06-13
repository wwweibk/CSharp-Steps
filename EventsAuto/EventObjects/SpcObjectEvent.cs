////////////////////////////////////////////////////////////////////////////////
//
// SpcObjectEvent - ���������� ������� �������� ��������� ������������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcObjectEvent : BaseEvent, ksSpcObjectNotify
	{
		public SpcObjectEvent(object obj, object doc, int objType, bool selfAdvise)
			: base(obj, typeof(ksSpcObjectNotify).GUID, doc,
			objType, null, selfAdvise) {}

		// koBeginDelete - ������� �������� �������
		public bool BeginDelete(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginDelete\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// koDelete - ������ ������
		public bool Delete(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> Delete\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			// ���� ��������� ������ ������� � �������� �� ��� �������
			// ���� ��� ������� : 0, SPC_BASE_OBJECT, SPC_COMMENT, ���� ��������� �� ������
			if (objRef != 0 && objRef == m_ObjType)
				this.Delete(objRef); 
			return true;
		}
       

		// soCellDblClick - ������� ���� � ������ 
		public bool CellDblClick(int objRef, int number)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CellDblClick\nobjRef = {1}\nnumber = {2}", m_LibName, objRef, number);
				str += "\n��� ��������� = " + GetDocName(); 
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;  
		}


		// soCellBeginEdit - ������ �������������� � ������   
		public bool CellBeginEdit(int objRef, int number) 
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CellBeginEdit\nobjRef = {1}\nnumber = {2}", m_LibName, objRef, number);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soChangeCurrent - ��������� ������� ������  
		public bool ChangeCurrent(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChangeCurrent\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;  
		}


		// soDocumentBeginAdd - ������ ���������� ���������
		public bool DocumentBeginAdd(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentBeginAdd\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soDocumentAdd - ���������� ��������� � ������� ��   
		public bool DocumentAdd(int objRef, string docName) 
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentAdd\nobjRef = {1}\ndocName = {2}", m_LibName, objRef, docName);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// soDocumentRemove - �������� ��������� �� ������� ��  
		public bool DocumentRemove(int objRef, string docName)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentRemove\nobjRef = {1}\ndocName = {2}", m_LibName, objRef, docName);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// soBeginGeomChange - ������ ������� ��������� ������� ��
		public bool BeginGeomChange(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginGeomChange\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soGeomChange - ��������� ������� �� ����������    
		public bool GeomChange(int objRef) 
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> GeomChange\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;  
		}


		// koBeginProcess - ������ ��������������\�������� �������
		public bool BeginProcess(int pType, int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginProcess\npType = {1}\nobjRef = {2}", m_LibName, pType, objRef);
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;  
		}
       

		// koEndProcess - ����� ��������������\�������� �������
		public bool EndProcess(int pType)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> EndProcess\npType = {1}", m_LibName, pType);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// koCreate - �������� ��������
		public bool CreateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CreateObject\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}

    
		// koUpdateObject - �������������� �������
		public bool UpdateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> UpdateObject\nobjRef = {1}", m_LibName, objRef);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}   
			return true;
		} 

    // koBeginCopy - ����������� �������
    public bool BeginCopy( int objRef )
    {
      return true;
    }

    // koCopy - ����������� �������
    public bool copy( int objRef )
    {
      return true;
    }

    }
}
