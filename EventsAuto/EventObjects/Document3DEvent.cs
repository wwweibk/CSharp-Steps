////////////////////////////////////////////////////////////////////////////////
//
// Document3DEvent - ���������� ������� �� 3D ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Steps.NET
{
	public class Document3DEvent : BaseEvent, ksDocument3DNotify
	{
		public Document3DEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksDocument3DNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// d3BeginRebuild - ������ ������������ ������
		public bool BeginRebuild()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginRebuild";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			else
				return true;
		}


		// d3Rebuild - ������ �����������
		public bool Rebuild()
		{   
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> Rebuild";
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d3BeginChoiceMaterial - ������ ������ ���������
		public bool BeginChoiceMaterial()
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginChoiceMaterial";
				str += "\n��� ��������� = " + GetDocName();
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (!res)
				{   
					ksDocument3D doc3D = (ksDocument3D)m_Doc;
					if (doc3D != null) 
					{  
						ksPart partObj = (ksPart)doc3D.GetPart((int)Part_Type.pTop_Part);
						partObj.SetMaterial("Material", 36.6);
						partObj.Update();
					}
				}
			}
			return res;
		}


		// d3�hoiceMaterial - �������� ����� ���������
		public bool ChoiceMaterial(string material, double density)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChoiceMaterial\nmaterial = {1}\ndensity = {2}", m_LibName, material, density);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d3BeginChoiceMarking - ������ ������ �����������
		public bool BeginChoiceMarking()
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginChoiceMarking";
				str += "\n��� ��������� = " + GetDocName();
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (!res)
				{   
					ksDocument3D doc3D = (ksDocument3D)m_Doc;
					if (doc3D != null) 
					{  
						ksPart partObj = (ksPart)doc3D.GetPart((int)Part_Type.pTop_Part);
						partObj.marking = "Marking";
						partObj.Update();
					}
				}
			}
			return res;
		}


		// d3�hoiceMarking - �������� ����� �����������
		public bool ChoiceMarking(string marking)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChoiceMarking\nmarking = {1}", m_LibName, marking);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d3BeginSetPartFromFile - ������ ��������� ���������� � ������ (�� ������� ������ �����)
		public bool BeginSetPartFromFile()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginSetPartFromFile";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// d3BeginCreatePartFromFile - ������ �������� ���������� � ������ (�� ������� ������ �����)
		public bool BeginCreatePartFromFile(bool part, entity plane)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginCreatePartFromFile";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}

    public bool CreateEmbodiment(string marking)
    {
      return true;
    }

    public bool DeleteEmbodiment(string marking)
    {
      return true;
    }
    public bool ChangeCurrentEmbodiment(string marking)
    {
      return true;
    }
    public bool BeginChoiceProperty(object obj, double propID)
    {
      return true;
    }

		public bool ChoiceProperty(object obj, double propID)
		{
			return true;
		}
		public bool BeginRollbackFeatures()
    {
      return true;
    }
    public bool RollbackFeatures()
    {
      return true;
    }
    public bool BedinLoadCombinationChange(int index)
    {
      return true;
    }
    public bool LoadCombinationChange(int index)
    {
      return true;
    }
    public bool BeginDeleteMaterial()
    {
      return true;
    }
    public bool DeleteMaterial()
    {
      return true;
    }
    public bool BeginDeleteProperty([MarshalAs(UnmanagedType.IDispatch)] object obj, double propID)
    {
      return true;
    }
    public bool DeleteProperty([MarshalAs(UnmanagedType.IDispatch)] object obj, double propID)
    {
      return true;
    }
  }
}
