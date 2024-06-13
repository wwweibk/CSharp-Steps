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

namespace Steps.NET
{
	public class DocumentEvent : BaseEvent, Kompas6API5.ksDocumentFileNotify
	{
		public DocumentEvent(object doc, bool selfAdvise)
			: base(doc, typeof(Kompas6API5.ksDocumentFileNotify).GUID, doc,
			-1, null, selfAdvise) {}


		// kdBeginCloseDocument - ������ �������� ���������
		public bool BeginCloseDocument()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginCloseDocument";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// kdCloseDocument - �������� ������
		public bool CloseDocument()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> CloseDocument";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			TerminateEvents(null, m_Doc, -1, null);
			return true;
		}


		// kdBeginSaveDocument - ������ ���������� ���������
		public bool BeginSaveDocument(string fileName)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginSaveDocument\nfileName = " + fileName;
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// kdSaveDocument - �������� ��������
		public bool SaveDocument()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> SaveDocument";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// kdActiveDocument - �������� �������������.
		public bool Activate()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> Activate";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// kdDeactiveDocument - �������� ���������������.
		public bool Deactivate()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> Deactivate";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// kdBeginSaveAsDocument - ������ ���������� ��������� c ������ ������ (�� ������� ������ �����)
		public bool BeginSaveAsDocument()
		{
			bool res = true;
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginSaveAsDocument\n" +
					"�� - ��������� ������\n" +
					"��� - ��������� ����������� ������ ���������� �����\n" +
					"������ - �� ��������� ����";
				int comm = Global.Kompas.ksYesNo(str);
				switch (comm) 
				{
					case 1: 
					{
						if (m_Doc != null)
						{ 
							ksDocument2D doc2D = (ksDocument2D)m_Doc;
							if (doc2D != null)
							{
								switch(Global.Kompas.ksGetDocumentType(doc2D.reference)) 
								{
									case (short)DocType.lt_DocFragment:   
										res = !doc2D.ksSaveDocument("C:\\1.frw");
										break;
									case (short)DocType.lt_DocSheetStandart:
									case (short)DocType.lt_DocSheetUser:
										res = !doc2D.ksSaveDocument("C:\\1.cdw");
										break;
								}
							}
							else
							{
								ksDocument3D doc3D = (ksDocument3D)m_Doc;
								if (doc3D != null)
								{  
									if (doc3D.IsDetail())
										res = !doc3D.SaveAs("C:\\1.m3d"); 
									else
										res = !doc3D.SaveAs("C:\\1.a3d");
								}
								else
								{
									ksSpcDocument docSpc = (ksSpcDocument)m_Doc;
									if (docSpc != null) 
									{   
										res = !docSpc.ksSaveDocument("C:\\1.spw");
									}
									else
									{
										ksDocumentTxt docTxt = (ksDocumentTxt)m_Doc;
										if (docTxt != null) 
										{   
											res = !docTxt.ksSaveDocument("C:\\1.kdw");
										}
									}
								}
							}
						}
					}	
						break;
					case 0:
						res = true;
						break;
					case -1:
						res = false;
						break;
				}
			}	
			return res;
		}


		// kdDocumentFrameOpen - ���� ��������� ���������
		public bool DocumentFrameOpen(object v) 
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = m_LibName + " --> DocumentFrameOpen";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
 
			if (!FindEvent(typeof(DocumentFrameEvent), (object)m_Container, 0, null))
			{ 
				if (v != null) 
				{
					DocumentFrameEvent frmEvent = new DocumentFrameEvent((ksDocumentFrameNotify_Event)v, (object)m_Container, true); 
					frmEvent.Advise();                                               
				}
			}
			return true;
		}

		public bool ProcessActivate(int id)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = string.Format("{0} --> ProcessActivate, ProcessID = {1}", m_LibName, id);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		public bool ProcessDeactivate(int id)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocEvents.Checked)
			{
				string str = string.Format("{0} --> ProcessDeactivate, ProcessID = {1}", m_LibName, id);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

    public bool BeginProcess(int iD)
    {
      return true;
    }
    public bool EndProcess(int iD, bool Success )
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
  }
}
