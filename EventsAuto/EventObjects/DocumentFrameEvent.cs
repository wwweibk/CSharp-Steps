////////////////////////////////////////////////////////////////////////////////
//
// DocumentFrameEvent  - ���������� ������� �� ���� ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using KompasAPI7;


using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class DocumentFrameEvent : BaseEvent, ksDocumentFrameNotify
	{
		public DocumentFrameEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksDocumentFrameNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// frBeginPaint - ������ ��������� ���������
		public bool BeginPaint(IPaintObject paintObj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> BeginPaint";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frClosePaint - ����� ��������� ���������
		public bool ClosePaint(IPaintObject paintObj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> ClosePaint";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frMouseDown - ������� ������ ����
		public bool MouseDown(short nButton, short nShiftState, int x, int y)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> MouseDown\nnButton = {1}\nnShiftState = {2}\nX = {3}\nY = {4}", 
					m_LibName, nButton, nShiftState, x, y);
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frMouseUp - ���������� ������ ����
		public bool MouseUp(short nButton, short nShiftState, int x, int y)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> MouseUp\nnButton = {1}\nnShiftState = {2}\nX = {3}\nY = {4}", 
					m_LibName, nButton, nShiftState, x, y);
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frMouseMove - �������� ����
		public bool MouseMove(short nShiftState, int x, int y)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> MouseMove\nnShiftState = {1}\nX = {2}\nY = {3}", 
					m_LibName, nShiftState, x, y);
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frMouseDblClick - ������� ���� ������ ����
		public bool MouseDblClick(short nButton, short nShiftState, int x, int y)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> MouseDblClick\nnButton = {1}\nnShiftState = {2}\nX = {3}\nY = {4}", 
					m_LibName, nButton, nShiftState, x, y);
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frBeginPaintGL - ������ ��������� � ��������� OpenGL
		public bool BeginPaintGL(ksGLObject glObj, int drawMode)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> BeginPaintGL";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frClosePaintGL - ��������� ��������� � ��������� OpenGL
		public bool ClosePaintGL(ksGLObject glObj, int drawMode)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> CloseClosePaintGL";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frAddGabarit - ����������� ��������� ���������
		public bool AddGabarit(IGabaritObject gabObj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> AddGabarit";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frBeginCurrentProcess - ������ �������� ��������
		public bool BeginCurrentProcess(int id)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> BeginCurrentProcess";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frStopCurrentProcess - ��������� �������� ��������
		public bool StopCurrentProcess(int id)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> StopCurrentProcess";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frActivate - ���� ����������������
		public bool Activate()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> Activate";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frDeactivate - ���� ������������������
		public bool Deactivate()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> Deactivate";
				Global.Kompas.ksMessage(str);
			}
			return true;
		}

		// frCloseFrame - �������� ����
		public bool CloseFrame()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbDocFrameEvents.Checked)
			{
				string str = m_LibName + " --> CloseFrame";
				Global.Kompas.ksMessage(str);
				BaseEvent.TerminateEvents();
			}
			return true;
		}	

    public bool ShowOcxTree( object ocx, bool show )
    {
      return true;
    }

    // 
    public bool BeginPaintTmpObjects()
    {
      return true;
    }

    // 
    public bool ClosePaintTmpObjects()
    {
      return true;
    }
  }
}
