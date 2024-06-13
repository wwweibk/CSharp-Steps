////////////////////////////////////////////////////////////////////////////////
//
// DocumentEvent  - ���������� ������� �� ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using KompasAPI7;

using System;
using Kompas6Constants;
using KAPITypes;
 

using reference = System.Int32;

namespace Steps.NET
{
	public class DocumentEvent : BaseEvent, Kompas6API5.ksDocumentFileNotify
	{
		public DocumentEvent(object doc)
			: base(doc, typeof(Kompas6API5.ksDocumentFileNotify).GUID, doc) {}


		// kdBeginCloseDocument - ������ �������� ���������
		public bool BeginCloseDocument()
		{
			return true;
		}


		// kdCloseDocument - �������� ������
		public bool CloseDocument()
		{
			BaseEvent.TerminateEvents(null, this.m_Doc);
			Global.UpdateSlideBox(null);
			return true;
		}


		// kdBeginSaveDocument - ������ ���������� ���������
		public bool BeginSaveDocument(string fileName)
		{
			return true;
		}


		// kdSaveDocument - �������� ��������
		public bool SaveDocument()
		{
			return true;
		}


		// kdActiveDocument - �������� �������������.
		public bool Activate()
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// kdDeactiveDocument - �������� ���������������.
		public bool Deactivate()
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// kdBeginSaveAsDocument - ������ ���������� ��������� c ������ ������ (�� ������� ������ �����)
		public bool BeginSaveAsDocument()
		{
			return true;
		}


		// kdDocumentFrameOpen - ���� ��������� ���������
		public bool DocumentFrameOpen(object v) 
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		public bool ProcessActivate(int id)
		{
			return true;
		}


		public bool ProcessDeactivate(int id)
		{
			Global.UpdateSlideBox(null);
			return true;
		}

    public bool BeginProcess(int iD)
    {
      return true;
    }
    public bool EndProcess(int iD, bool Success)
    {
      return true;
    }

    public bool BeginAutoSaveDocument(string fileName)
    {
      return true;
    }

    public bool AutoSaveDocument()
    {
      return true;
    }

    // ������� ���������� ������� ���������
    public static BaseEvent NewDocumentEvent(object doc)
		{
			DocumentEvent res = null;
			if (doc != null) 
			{
				if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc)) 
				{
					res = new DocumentEvent(doc);
					res.Advise();
				}
			}
			return res;
		}


		// ������� ���������� ������� 2D ���������
		public static BaseEvent NewDocument2DEvent(object doc)
		{
			Document2DEvent res = null;
			if (doc != null) 
			{
				if (BaseEvent.FindEvent(typeof(Document2DEvent), doc)) 
				{
					ksDocument2D doc2d = (ksDocument2D)doc;
					if (doc2d != null)
					{
						NewDocumentEvent(doc);	// ����� ��� �������� ��������� ����������
            res = new Document2DEvent( doc2d.GetDocument2DNotify(),  doc );
						res.Advise();
					}
				}
			}
			return res;
		}

    // ������� ���������� ������� 3D ���������
    public static BaseEvent NewDocument3DEvent(object doc)
    {
      Document3DEvent res = null;
      if (doc != null)
      {
        if (BaseEvent.FindEvent(typeof(Document3DEvent), doc))
        {
          ksDocument3D doc3d = (ksDocument3D)doc;
          if (doc3d != null)
          {
            NewDocumentEvent(doc);	// ����� ��� �������� ��������� ����������
            res = new Document3DEvent(doc3d.GetDocument3DNotify(), doc);
            res.Advise();
          }
        }
      }
      return res;
    }

	}
}
