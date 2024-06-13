////////////////////////////////////////////////////////////////////////////////
//
// Document2DEvent - ���������� ������� �� 2D ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class Document2DEvent : BaseEvent, ksDocument2DNotify
	{
		public Document2DEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksDocument2DNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// d3BeginRebuild - ������ ������������ ������
		public bool BeginRebuild()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginRebuild";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// d3Rebuild - ������ �����������
		public bool Rebuild()
		{   
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
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
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginChoiceMaterial";
				str += "\n��� ��������� = " + GetDocName();
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (!res)
				{   
					ksTextLineParam parLine = (ksTextLineParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
					ksTextItemParam parItem = (ksTextItemParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

					// C������ ������ ����� ������
					ksDynamicArray arMaterial = (ksDynamicArray)Global.Kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
					if (arMaterial != null && parLine != null && parItem != null)
					{
						parLine.Init();
						parItem.Init();

						// ������ ��������� ������ ������
						ksDynamicArray item = (ksDynamicArray)parLine.GetTextItemArr();
						if (item != null) 
						{
							ksTextItemFont font = (ksTextItemFont)parItem.GetItemFont();
							if (font != null) 
							{
								// ������� ������ ������ ������
								font.height = 10;   // ������ ������
								font.ksu = 1;       // ������� ������
								font.color = 1000;  // ����
								font.bitVector = 1; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
								parItem.s = "(1-� ����������)";
								// �������� 1-� ����������  � ������ ���������
								item.ksAddArrayItem(-1, parItem);

								font.height = 20;   // ������ ������
								font.ksu = 2;       // ������� ������
								font.color = 2000;  // ����
								font.bitVector = 2; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
								parItem.s = "(2-� ����������)";
								// �������� 2-� ����������  � ������ ���������
								item.ksAddArrayItem(-1, parItem);
				    
								parLine.style = 1;

								// 1-� ������ ������ ������� �� ���� ��������� ������� 
								// ������ ������ � ������ ����� ������
								arMaterial.ksAddArrayItem(-1, parLine);

								ksDocument2D doc2D = (ksDocument2D)m_Doc;
								if (doc2D != null) 
								{
									doc2D.ksSetMaterialParam(arMaterial, 36.6);
								}
							}
						}
					}
				}
			}
			return res;
		}


		// d3�hoiceMaterial - �������� ����� ���������
		public bool ChoiceMaterial(string material, double density)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChoiceMaterial\nmaterial = {1}\ndensity = {2}", m_LibName, material, density);
				str += "\n��� ��������� = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d2BeginInsertFragment - ������ ������� ��������� (�� ������� ������ �����)
		public bool BeginInsertFragment()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginInsertFragment";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// d2LocalFragmentEdit
		public bool LocalFragmentEdit(object newDoc, bool newFrw)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> LocalFragmentEdit";
				str += "\n��� ��������� = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}

    public bool BeginChoiceProperty(int objRef, double propID)
    {
      return true;
    }

    public bool ChoiceProperty(int objRef, double propID)
    {
      return true;
    }
    public bool BeginDeleteProperty(int objRef, double propID)
    {
      return true;
    }

    public bool DeleteProperty(int objRef, double propID)
    {
      return true;
    }
  }
}
