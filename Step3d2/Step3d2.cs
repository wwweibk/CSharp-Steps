using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step3d2 - ������ � ����������� (������ ��� ������)
	public class Step3d2
	{
		private KompasObject kompas;
		private ksDocument3D doc;
		private string buf;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d2 - ������ � ����������� (������ ��� ������)";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				switch (command)
				{
					case 1 : CreateDocument3D();     break;  // �������� ���������
					case 2 : DocIterator();          break;  // �������� �� ����������
					case 3 : UseEntityCollection();  break;  // ������������� ������� ���������
					default:
						doc = (ksDocument3D)kompas.ActiveDocument3D();
						if (doc != null && doc.reference != 0)
						{
							switch (command)
							{
								case 4  : GetSetPartName();				break; // �����/�������� ��� ����������
								case 5  : FixAndStandartComponent();	break; // ������������ � ��������� ������������ �������
								case 6  : GetSetColorProperty();		break; // �������� � �������� ��������� ����� ����������
								case 7  : GetSetArrayVariable();		break; // ����� � �������� ������� ���������� ����������
								case 8  : GetSetPlacmentComponent();	break; // �������� � �������� ����� ������������ ������ � ������
								case 9  : GetSetEntity();				break; // �������� ��������� ksEntity ������� ������������ �������� �� ��������� � �������� ���������
								case 10 : CreateSketch();				break; // ������� �����
								case 11 : GetArraySketch();				break; // ��������� ������ ��������(����� �������) � ���������� ��� ��������� ksEntityCollection (IEntityCollection)
								case 12 : GetSetUserParamComponent();	break; // ���������� � �������� ��������� ������������ � ����������
							}
						}
						break;
				}
			}
		}


		// ������������ ���� ����������
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;	//�� ��������� - ������ ������
			itemType = 1;					//MENUITEM

			switch (number)
			{
				case 1:
					result = "�������� 3D ���������";
					command = 1;
					break;
				case 2:
					result = "�������� �� ����������";
					command = 2;
					break;
				case 3:
					result = "������� ���������";
					command = 3;
					break;
				case 4:
					result = "����� � �������� ��� ����������";
					command = 4;
					break;
				case 5:
					result = "������������ � ��������� ������������ �������";
					command = 5;
					break;
				case 6:
					result = "�������� � �������� ��������� ����� ����������";
					command = 6;
					break;
				case 7:
					result = "����� � �������� ������� ���������� ����������";
					command = 7;
					break;
				case 8:
					result = "�������� � �������� ������������ ������ � ������";
					command = 8;
					break;
				case 9:
					result = "�������� � �������� �������� ������� ���������";
					command = 9;
					break;
				case 10:
					result = "������� �����";
					command = 10;
					break;
				case 11:
					result = "������� ������ �������";
					command = 11;
					break;
				case 12:
					result = "���������� � �������� ��������� ������������ � ����������";
					command = 12;
					break;
				case 13:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// �����/�������� ��� ����������
		void GetSetPartName()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ������� ���������
			if (part != null)
			{
				buf = string.Format("��� ���������� {0}", part.name);
				kompas.ksMessage(buf);
				part.name = "������";
				part.Update();
			}	 
		}


		// ������������ � ��������� ������������ �������
		void FixAndStandartComponent()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("������� �������� ������ ���� �������");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// ������ ������ � ������
			if (part != null)
			{
				// �������� ��������� �������� ���������� - � ������� �������������� ������ �� ���������� - fixedComponent
				bool fix = part.fixedComponent;
				// �������� ��������� ������������ ���������� - � ������� �������������� ������ �� ���������� - standardComponent
				bool stand = part.standardComponent;
				kompas.ksMessage(fix ? "��������� ������������" : "��������� �� ������������");
				// �������� ��������� �������� ���������� - � ������� �������������� ������ �� ���������� - fixedComponent
				part.fixedComponent = !fix;
				kompas.ksMessage(stand ? "��������� �����������" : "��������� �������������");
				// �������� ��������� ������������ ���������� - � ������� �������������� ������ �� ���������� - standardComponent
				part.standardComponent = !stand;
			}
		}


		// �������� � �������� ��������� ����� ����������
		void GetSetColorProperty()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part); // ������� ���������
			if (part != null)
			{
				ksColorParam colorPr = (ksColorParam)part.ColorParam();
				if (colorPr != null)
				{
					buf = string.Format("����� ����� = {0}\n����� ���� = {1}\n�������� = {2}\n������������ = {3}\n����� = {4}\n������������ = {5}\n��������� = {6}", 
						colorPr.color, colorPr.ambient, colorPr.diffuse,
						colorPr.specularity, colorPr.shininess,
						colorPr.transparency, colorPr.emission);
					kompas.ksMessage(buf);
					colorPr.color = 5421504;
					colorPr.transparency = 0.5;
					colorPr.ambient = 0.1;
					colorPr.diffuse = 0.1;
					part.Update();
					buf = string.Format("����� ����� = {0}\n����� ���� = {1}\n�������� = {2}\n������������ = {3}\n����� = {4}\n������������ = {5}\n��������� = {6}", 
						colorPr.color, colorPr.ambient, colorPr.diffuse, 
						colorPr.specularity, colorPr.shininess, 
						colorPr.transparency, colorPr.emission);
					kompas.ksMessage(buf);
				}
			}
		}


		// ����� � �������� ������� ���������� ����������
		void GetSetArrayVariable()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("������� �������� ������ ���� �������");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// ������ ������ � ������
			if (part != null)
			{
				// ������ � �������� ������� ����������
				ksVariableCollection varCol = (ksVariableCollection)part.VariableCollection();
				if (varCol != null)
				{
					ksVariable var = (ksVariable)kompas.GetParamStruct((short)StructType2DEnum.ko_VariableParam);
					if (var == null)
						return;
					for (int i = 0; i < varCol.GetCount(); i ++)
					{
						var = (ksVariable)varCol.GetByIndex(i);
						buf = string.Format("����� ���������� {0}\n��� ���������� {1}\n�������� ���������� {2}\n����������� {3}", i, var.name, var.value, var.note);
						kompas.ksMessage(buf);
						if (i == 0)
						{
							var.note = "qwerty";
							double d = 0;
							kompas.ksReadDouble("����� ����������", 10, 0, 100, ref d);
							var.value = d;
						}
					}
			
					for (int j = 0; j < varCol.GetCount(); j ++)
					{
						// �������� ����������� ����������
						var = (ksVariable)varCol.GetByIndex(j);
						buf = string.Format("����� ���������� {0}\n��� ���������� {1}\n�������� ���������� {2}\n����������� {3}", j, var.name, var.value, var.note);
						kompas.ksMessage(buf);
					}
					part.RebuildModel();	// ������������ ������
				}
			}
		}


		// �������� � �������� ����� ������������ ������ � ������
		void GetSetPlacmentComponent()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("������� �������� ������ ���� �������");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// ������ ������ � ������
			if (part != null)
			{
				ksPlacement plac = (ksPlacement)part.GetPlacement();
				if (plac != null)
				{
					double x = 0, y = 0, z = 0;
					plac.GetOrigin(out x, out y, out z);
					buf = string.Format("x = {0}\ny = {1}\nz = {2}", x, y, z);
					kompas.ksMessage(buf);
					plac.SetOrigin(20, 20, 20);
					part.SetPlacement(plac);
					part.UpdatePlacement();
					part.Update();
				}
			}
		}


		// �������� ��������� ksEntity ������� ������������ �������� �� ��������� � �������� ���������
		void GetSetEntity()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ������� ���������
			if (part != null)
			{
				ksEntity planeXOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// 1-��������� �� ��������� XOY
				if (planeXOY != null)
				{
					kompas.ksMessage(planeXOY.name);
					planeXOY.name = "Plane";
					planeXOY.Update();
				}
			}
		}


		// �������� ������
		void CreateSketch()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ������� ���������
			if (part != null) 
			{
				ksEntity planeXOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// 1-��������� �� ��������� XOY
				ksEntity entity = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (planeXOY != null && entity != null)
				{
					ksSketchDefinition sketch = (ksSketchDefinition)entity.GetDefinition();
					if (sketch != null)
					{
						sketch.SetPlane(planeXOY);
						entity.Create();
						ksDocument2D sketchDoc = (ksDocument2D)sketch.BeginEdit(); //GetSketchEditor());
						sketchDoc.ksLineSeg(0, 0, 100, 100, 1);
						sketch.EndEdit();
					}
				}
			}
		}


		// ��������� ������ ��������(����� �������) � ���������� ��� ��������� ksEntityCollection (IEntityCollection)
		void GetArraySketch()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ������� ���������
			if (part != null)
			{
				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_sketch);
				ksEntity currentEntity;
				if (entityCollection != null)
				{
					for (int k = 0; k < entityCollection.GetCount(); k++)
					{
						currentEntity = (ksEntity)entityCollection.GetByIndex(k);
						kompas.ksMessage(currentEntity.name);
					}
				}
			}
		}


		private struct dstruct
		{
			public double a, b;
			public int c, d;
		}


		// ���������� � �������� ��������� ������������ � ����������
		void GetSetUserParamComponent() 
		{
			if (doc.IsDetail()) 
			{
				kompas.ksError("������� �������� ������ ���� �������");
				return;
			}

			ksPart part = (ksPart)doc.GetPart(0);	// ������ ������ � ������
			ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par == null || item == null || arr == null || part == null)
				return;

			par.Init();
			par.SetUserArray(arr);
			item.Init();
			item.doubleVal = 12.12;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.doubleVal = 21.21;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.intVal = 666;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.intVal = 999;
			arr.ksAddArrayItem(-1, item);

			part.SetUserParam(par);	// ��������� ���������������� ���������
			part.Update();

			buf = string.Format("������ ���������������� ��������� {0}", part.GetUserParamSize()); // ������ ���������������� ���������
			kompas.ksMessage(buf);

			ksUserParam par2 = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item2 = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr2 = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par2 == null || item2 == null || arr2 == null)
				return;

			par2.Init();
			par2.SetUserArray(arr2);
			item2.Init();
			item2.doubleVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.doubleVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.intVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.intVal = 0;
			arr2.ksAddArrayItem(-1, item2);

			part.GetUserParam(par2);	// ����� ��������������e�� ���������
  
			dstruct d;

			arr2.ksGetArrayItem(0, item2);
			d.a = item2.doubleVal;
			arr2.ksGetArrayItem(1, item2);
			d.b = item2.doubleVal;
			arr2.ksGetArrayItem(2, item2);
			d.c = item2.intVal;
			arr2.ksGetArrayItem(3, item2);
			d.d = item2.intVal;
			buf = string.Format("���������� ����������������� �������\na = {0}\nb = {1}\nc = {2}\nd = {3}", d.a, d.b, d.c, d.d);
			kompas.ksMessage(buf);	// ���������� ���������� �� ����������������� �������
		}


		// �������� ��������� 
		void CreateDocument3D()
		{
			ksDocument3D doc = (ksDocument3D)kompas.Document3D();
			if (doc.Create(false /*�������*/, true /*������*/))
			{
				doc.author = "Ethereal";				// ����� ���������
				doc.comment = "������ ��������� 3D";	// ����������� � ���������
				doc.fileName = @"c:\example.m3d";		// ��� ����� ���������
				doc.UpdateDocumentParam();				// �������� ��������� ���������
				doc.Save();								// ��������� ��������
				kompas.ksMessage("�������� �������� ��� ������ ������");
				doc.SaveAs(@"c:\example_1.m3d");		// ��������� �������� ��� ������ ������

				// ����� ���������
				buf = string.Format("����� ���������: {0}", doc.author);
				kompas.ksMessage(buf);
				// ����������� � ���������
				buf = string.Format("����������� � ���������: {0}", doc.comment);
				kompas.ksMessage(buf);
				// ��� �����
				buf = string.Format("��� �����: {0}", doc.fileName);
				kompas.ksMessage(buf);

				doc.close();	// ������� ��������
			}
		}


		// �������� ������ ���������� � �������� ��������� �� ���
		void DocIterator()
		{
			ksDocument3D doc = null;
			ksDynamicArray arrDoc = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);	// ������������ ������ ���������� �� ������ ��������
			// ���� ������ ������ ��������� ������ ������ ������ 
			if (arrDoc != null && kompas.ksChoiceFiles("*.m3d","��������� (*.m3d)|*.m3d|��� ����� (*.*)|*.*", arrDoc, false) == 1)
			{
				ksChar255 item = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255); 
				if (item != null)
				{
					// ������� ��� ����� ��������� �������������
					for (int i = 0, count = arrDoc.ksGetArrayCount(); i < count; i++)
					{
						if (arrDoc.ksGetArrayItem(i, item) == 1)
						{ 
							// ����������� ��������� ���������
							doc = (ksDocument3D)kompas.Document3D();
							doc.Open(item.str, false);	// ��������� ���� � �������� ������ 
						}
					}
				}

				// �������� ���������� �������� ������
				buf = string.Format("������� {0} ������", arrDoc.ksGetArrayCount());
				kompas.ksMessage(buf); 

				// ������� �������� �� ����������
				ksIterator iter = (ksIterator)kompas.GetIterator();
				if (iter != null)
				{
					//��� �������             
					if (iter.ksCreateIterator(ldefin2d.D3_DOCUMENT_OBJ, 0))
					{
						reference rf; // �������� ���������
						if ((rf = iter.ksMoveIterator("F")) != 0)
						{
							// ������� �������� �� ������ �������
							do
							{
								// ������� �������� �� ��������� �������
								// ����������� ��������� ���������
								doc = (ksDocument3D)kompas.ksGet3dDocumentFromRef(rf);
								if (doc != null)
								{
									// ����� ���������
									buf = string.Format("����� ���������: {0}", doc.author);
									kompas.ksMessage(buf);
									// ����������� � ���������
									buf = string.Format("����������� � ���������: {0}", doc.comment);
									kompas.ksMessage(buf);
									// ��� ����� ��������� 
									buf = string.Format("��� �����: {0}", doc.fileName);
									kompas.ksMessage(buf);
									// ��� ���������
									buf = string.Format("��� ���������: {0}", doc.IsDetail() ? "������" : "������");
									kompas.ksMessage(buf);
								}
							}
							while ((rf = iter.ksMoveIterator("N")) != 0);

							iter.ksDeleteIterator(); // ������� ��������
						}
					}
				}
			}
		}


		// ������������� ������� ���������
		void UseEntityCollection()
		{
			ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D();	// ������������� � ��������� ���������
			if (doc != null)
			{
        ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
				if (part != null) 
				{
					int count1 = 0;
					int count2 = 0;	// ���������� ������� ������������
					int count = 0;	// ���������� ���������� ������������

					// ������ ������������
					ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);      
					if (collect != null)
					{
						count = collect.GetCount();
						count1 = 0; 
						count2 = 0; 
						ksColorParam colorPr = null;	// ��������� ������� �����
						if (collect != null && count != 0)
						{
							for (int i = 0; i < count; i ++)
							{
								ksEntity ent = (ksEntity)collect.GetByIndex(i);
								colorPr = (ksColorParam)ent.ColorParam();
								// ��������� ������� �����������
								ksFaceDefinition faceDef = (ksFaceDefinition)ent.GetDefinition();
								if (faceDef != null)
								{
									// ���������� ��-��		//�������������� ��-��
									if (faceDef.IsCone() || faceDef.IsCylinder()) 
									{      
										colorPr.color = Color.FromArgb(0, 255, 255, 0).ToArgb();  
										count2 ++;	// ������� ���������� ��������
									}
									// ������� ��-��  
									if (faceDef.IsPlanar()) 
									{
										colorPr.color = Color.FromArgb(0, 0, 255, 255).ToArgb(); 
										count1 ++;	// ������� ���������� ��������
									}

									ent.Update();	// �������� ���������
								}
							}
						}
					}

					// �������� � ����������� ������
					if (count == 0) 
						kompas.ksMessage("�� ������� �� ����� �����������");
					else
					{
						buf = string.Format("������� {0} ���������� � {1} ������� ��������", count2, count1);
						kompas.ksMessage(buf);
					}

					count1 = 0;
					count2 = 0;
					// ������ �����
					ksEntityCollection collect2 = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_edge);
					count = collect2.GetCount();
					if (collect2 != null && count != 0)
					{
						for (int i = 0; i < count; i++)
						{
							ksEntity ent = (ksEntity)collect2.GetByIndex(i);
							ksEdgeDefinition edgeDef = (ksEdgeDefinition)ent.GetDefinition();
							if (edgeDef != null)
							{
								if (edgeDef.IsStraight()) 
									count1++;	// ���������� ������ �����
								else
									count2++;	// ���������� ������������� �����
							}
						}
					}

					// �������� � ����������� ������
					if (count == 0) 
						kompas.ksMessage("�� ������� �� ������ �����");
					else
					{
						buf = string.Format("������� {0} ������ � {1} ������������� �����", count1, count2);
						kompas.ksMessage(buf);
					}
				}
			}
		}


		#region COM Registration
		// ��� ������� ����������� ��� ����������� ������ ��� COM
		// ��� ��������� � ����� ������� ���������� ������ Kompas_Library,
		// ������� ������������� � ���, ��� ����� �������� ����������� ������,
		// � ����� �������� ��� InprocServer32 �� ������, � ��������� ����.
		// ��� ��� �������� ��� ����, ����� ����� ����������� ����������
		// ���������� �� ������� ActiveX.
		[ComRegisterFunction]
		public static void RegisterKompasLib(Type t)
		{
			try
			{
				RegistryKey regKey = Registry.LocalMachine;
				string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
				regKey = regKey.OpenSubKey(keyName, true);
				regKey.CreateSubKey("Kompas_Library");
				regKey = regKey.OpenSubKey("InprocServer32", true);
				regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
				regKey.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("��� ����������� ������ ��� COM-Interop ��������� ������:\n{0}", ex));
			}
		}
		
		// ��� ������� ������� ������ Kompas_Library �� �������
		[ComUnregisterFunction]
		public static void UnregisterKompasLib(Type t)
		{
			RegistryKey regKey = Registry.LocalMachine;
			string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
			RegistryKey subKey = regKey.OpenSubKey(keyName, true);
			subKey.DeleteSubKey("Kompas_Library");
			subKey.Close();
		}
		#endregion
	}
}
