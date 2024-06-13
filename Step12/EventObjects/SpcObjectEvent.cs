////////////////////////////////////////////////////////////////////////////////
//
// SpcObjectEvent - ���������� ������� �������� ��������� ������������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using System.Reflection;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcObjectEvent : BaseEvent, ksSpcObjectNotify
	{
		public SpcObjectEvent(object obj, object doc, int objType)
			: base(obj, typeof(ksSpcObjectNotify).GUID, doc, objType, 0) {}

		// koBeginDelete - ������� �������� �������
		public bool BeginDelete(int objRef)
		{
			return true;
		}


		// koDelete - ������ ������
		public bool Delete(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}
       

		// soCellDblClick - ������� ���� � ������ 
		public bool CellDblClick(int objRef, int number)
		{ 
			return true;  
		}


		// soCellBeginEdit - ������ �������������� � ������   
		public bool CellBeginEdit(int objRef, int number) 
		{ 
			return true;
		}


		// soChangeCurrent - ��������� ������� ������  
		public bool ChangeCurrent(int objRef)
		{ 
			return true;  
		}


		// soDocumentBeginAdd - ������ ���������� ���������
		public bool DocumentBeginAdd(int objRef)
		{ 
			return true;
		}


		// soDocumentAdd - ���������� ��������� � ������� ��   
		public bool DocumentAdd(int objRef, string docName) 
		{ 
			return true; 
		}


		// soDocumentRemove - �������� ��������� �� ������� ��  
		public bool DocumentRemove(int objRef, string docName)
		{ 
			return true;
		}


		// soBeginGeomChange - ������ ������� ��������� ������� ��
		public bool BeginGeomChange(int objRef)
		{ 
			return true;
		}


		// soGeomChange - ��������� ������� �� ����������    
		public bool GeomChange(int objRef) 
		{
			return true;  
		}


		// koBeginProcess - ������ ��������������\�������� �������
		public bool BeginProcess(int pType, int objRef)
		{
			return true;  
		}
       

		// koEndProcess - ����� ��������������\�������� �������
		public bool EndProcess(int pType)
		{
			Global.UpdateSlideBox(null);
			return true; 
		}


		// koCreate - �������� ��������
		public bool CreateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true; 
		}

    
		// koUpdateObject - �������������� �������
		public bool UpdateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		} 


		// ������� ���������� ������� ��� ������� ������������
		public static BaseEvent NewSpcObjectEvent(object doc, int objType)
		{
			SpcObjectEvent res = null;
			if (doc != null) 
			{
				if (!BaseEvent.FindEvent(typeof(SpcObjectEvent), doc, objType)) 
				{
					object spcObj = doc.GetType().InvokeMember("GetSpecification", BindingFlags.InvokeMethod, null, doc, null);
					ksSpecification spc = (ksSpecification)spcObj;
					if (spc != null)
					{
						DocumentEvent.NewDocumentEvent(doc); // ����� ��� �������� ��������� ����������
						res = new SpcObjectEvent(spc.GetSpcObjectNotify(objType), doc, objType);
						res.Advise();
					}
				}
			}
			return res;
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