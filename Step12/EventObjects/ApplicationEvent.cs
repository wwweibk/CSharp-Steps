////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - ���������� ������� �� ����������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.InteropServices;

namespace Steps.NET
{
	public class ApplicationEvent : BaseEvent, ksKompasObjectNotify
	{
		public ApplicationEvent(object obj)
			: base(obj, typeof(ksKompasObjectNotify).GUID, null) {}
		

		// koApplicatinDestroy - �������� ����������
		public bool ApplicationDestroy()
		{
			TerminateEvents();
			return true;
		}


		// koBeginCloseAllDocument - ������ �������� ���� �������� ����������
		public bool BeginCloseAllDocument()
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginCreate - ������ �������� ���������(�� ������� ������ ����)
		public bool BeginCreate(int docType)
		{
			return true;
		}


		// koOpenDocumenBegin - ������ �������� ���������
		public bool BeginOpenDocument(string fileName)
		{
			return true;
		}


		// koBeginOpenFile - ������ �������� ���������(�� ������� ������ �����)
		public bool BeginOpenFile()
		{
			return true;
		}


		// koActiveDocument - ������������ �� ������ �������� ��������
		public bool ChangeActiveDocument(object newDoc, int docType)
		{
			Global.UpdateSlideBox(newDoc); 
			return true;
		}


		// koCreateDocument - �������� ������
		public bool CreateDocument(object newDoc, int docType)
		{
			Global.AdviseDoc(newDoc, docType);
			Global.UpdateSlideBox(newDoc); 
			return true;
		}


		// koOpenDocumen - �������� ������
		public bool OpenDocument(object newDoc, int docType)
		{
			Global.AdviseDoc(newDoc, docType);
			Global.UpdateSlideBox(newDoc); 
			return true;
		}

    // koKeyDown - ������� ����������
    public bool KeyDown( ref int key, int flags, bool system )
    {
      return true;
    }

    // koKeyUp - ������� ����������
    public bool KeyUp( ref int key, int flags, bool system )
    {
      return true;
    }

    // koKeyPress - ������� ����������
    public bool KeyPress( ref int key, bool system )
    {
      return true;
    }

    public bool BeginReguestFiles( int type, ref object files )
    {
      return true;
    }
    public bool BeginChoiceMaterial(int MaterialPropertyId)
    {
      return true;
    }
    public bool ChoiceMaterial(int MaterialPropertyId, string material, double density)
    {
      return true;
    }
    public bool IsNeedConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam, ref bool needConvert)
    {
      return true;
    }

    public bool BeginConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
    {
      return true;
    }

    public bool EndConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
    {
      return true;
    }

    public bool ChangeTheme(int newTheme)
    {
      return true;
    }

    public bool BeginDragOpenFiles([MarshalAs(UnmanagedType.Struct)] ref object filesList, bool insert, ref bool filesListChanged)
    {
      return true;
    }

    public static BaseEvent NewApplicationEvent()
		{
			ApplicationEvent res = null;
			if (!BaseEvent.FindEvent(typeof(ApplicationEvent), null))
			{
				res = new ApplicationEvent(Global.Kompas);
				res.Advise();
			}
			return res;
		}
	}
}
