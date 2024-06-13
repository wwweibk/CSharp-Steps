using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step3d3 - ��������, ��� � ���������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step3d3
	{
		private KompasObject kompas;
		private ksDocument3D doc;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d3 - ��������, ��� � ���������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				doc = (ksDocument3D)kompas.ActiveDocument3D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1 : ConstrAxisOperations();	break; // �������������� ��� ��������
						case 2 : ConstrAxis2Point();		break; // �������������� ��� �� ���� ������
						case 3 : ConstrAxisEdge();			break; // �������������� ���, ���������� ����� �����
						case 4 : CreateConstrElem();		break; // �������� ��������� ���������, ��� �� ���� ���������� � ��������� ��� ����� � ��������
						case 5 : ConstrPlane3Point();		break; // ��������� ����� ��� �������
					}
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
					result = "�������������� ��� ��������";
					command = 1;
					break;
				case 2:
					result = "�������������� ��� �� ���� ������";
					command = 2;
					break;
				case 3:
					result = "�������������� ���, ���������� ����� �����";
					command = 3;
					break;
				case 4:
					result = "��������� ���������, ��� �� ���� ����������, ��������� ��� ����� � ������ ��-��";
					command = 4;
					break;
				case 5:
					result = "��������� ����� ��� �������";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// �������������� ��� ��������
		void ConstrAxisOperations()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// ��������� ������� ������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// ������� ��������� ������� ��������� XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						entitySketch.Create();			// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						sketchEdit.ksCircle(20, 0, 10, 1);
						sketchEdit.ksLineSeg(0, 0, 0, 5, 3);
						sketchDef.EndEdit();			// ���������� �������������� ������

						ksEntity entityRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
						bool res = false;
						if (entityRotate != null)
						{
							ksBossRotatedDefinition rotateDef = (ksBossRotatedDefinition)entityRotate.GetDefinition(); // ��������� ������� �������� ��������
							if (rotateDef != null)
							{
								ksRotatedParam rotproperty = (ksRotatedParam)rotateDef.RotatedParam();
								if (rotproperty != null)
								{
									rotproperty.direction = (short)Direction_Type.dtBoth;
									rotproperty.toroidShape = false;
								}

								rotateDef.SetThinParam(true, (short)Direction_Type.dtBoth, 1, 1);	// ������ ������ � ��� �����������
								rotateDef.SetSideParam(true, 180);
								rotateDef.SetSideParam(false, 180);
								rotateDef.SetSketch(entitySketch);	// ����� �������� ��������
								res = entityRotate.Create();		// ������� ��������
							}
						}
						if (res)
						{
							// �������� �������� ��������� �� ��������
							ksEntity entityAxisOperation = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisOperation);
							if (entityAxisOperation != null)
							{
								ksAxisOperationsDefinition axisOperation = (ksAxisOperationsDefinition)entityAxisOperation.GetDefinition();
								if (axisOperation != null)
								{
									axisOperation.SetOperation(entityRotate);
									entityAxisOperation.Create();	// ������� ��������
								}
							}
							kompas.ksMessage("��� ��������");
						}
						else
							kompas.ksMessage("������ �������� ��������");
					}
				}
			}
		}


		// �������������� ��� �� ���� ������
		void ConstrAxis2Point()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// ��������� ������� ������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// ������� ��������� ������� ��������� XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						entitySketch.Create();			// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// ������ ����� ����� - �������
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 
						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// ���������� �������������� ������                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// ��������� ������� ������� �������� ������������
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition();	// ��������� ������� �������� ������������
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// ����������� ������������
								extrusionDef.SetSideParam(true /*������ �����������*/, (short)End_Type.etBlind /*������ �� �������*/, 20, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 20, 20);	// ������ ������ � ��� �����������
								extrusionDef.SetSketch(entitySketch);									// ����� �������� ������������
								entityExtr.Create();													// ������� ��������
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_vertex);

				if (entityCollection != null && entityCollection.GetCount() != 0)
				{
					// �������� ��� �� ���� ������
					ksEntity entityAxis2Point = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axis2Points); 
					if (entityAxis2Point != null)
					{
						ksAxis2PointsDefinition axis2Point = (ksAxis2PointsDefinition)entityAxis2Point.GetDefinition();
						if (axis2Point != null)
						{
							axis2Point.SetPoint(1, entityCollection.GetByIndex(0));
							axis2Point.SetPoint(2, entityCollection.GetByIndex(entityCollection.GetCount() - 1));
							entityAxis2Point.Create();
						}
					}
					kompas.ksMessage("��� ����� ��� �����");
				}
			}
		}


		// �������������� ���, ���������� ����� �����
		void ConstrAxisEdge()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// ��������� ������� ������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// ������� ��������� ������� ��������� XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						entitySketch.Create();			// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// ������ ����� ����� - �������
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 

						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// ���������� �������������� ������                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// ��������� ������� ������� �������� ������������
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); // ��������� ������� �������� ������������
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// ����������� ������������
								extrusionDef.SetSideParam(true /*������ �����������*/, (short)End_Type.etBlind /*������ �� �������*/, 20, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 20, 20);	// ������ ������ � ��� �����������
								extrusionDef.SetSketch(entitySketch);									// ����� �������� ������������
								entityExtr.Create();													// ������� ��������
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_edge);

				if (entityCollection != null && entityCollection.GetCount() > 1)
				{
					// �������� ��� ����� �����
					ksEntity entityAxisEdge = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisEdge); 
					if (entityAxisEdge != null)
					{
						ksAxisEdgeDefinition axisEdge = (ksAxisEdgeDefinition)entityAxisEdge.GetDefinition();
						if (axisEdge != null)
						{
							axisEdge.SetEdge(entityCollection.GetByIndex(0));
							entityAxisEdge.Create();
						}
					}
					kompas.ksMessage("��� ����� �����");

					// �������� ��� ��� ����� �����
					ksEntity entityAxisEdge2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisEdge); 
					if (entityAxisEdge2 != null)
					{
						ksAxisEdgeDefinition axisEdge = (ksAxisEdgeDefinition)entityAxisEdge2.GetDefinition();
						if (axisEdge != null)
						{
							axisEdge.SetEdge(entityCollection.GetByIndex(1));
							entityAxisEdge2.Create();
						}
					}
					kompas.ksMessage("������ ��� ����� �����");
				}
			}
		}


		// ��������� ����� ��� �������
		void ConstrPlane3Point()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// ��������� ������� ������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// ������� ��������� ������� ��������� XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						entitySketch.Create();			// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// ������ ����� ����� - �������
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 

						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// ���������� �������������� ������                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// ��������� ������� ������� �������� ������������
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); // ��������� ������� �������� ������������
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// ����������� ������������
								extrusionDef.SetSideParam(true /*������ �����������*/, (short)End_Type.etBlind /*������ �� �������*/, 20, 30, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10);	// ������ ������ � ��� �����������
								extrusionDef.SetSketch(entitySketch);									// ����� �������� ������������
								entityExtr.Create();													// ������� ��������
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_vertex);
				if (entityCollection.GetCount() > 2)
				{
					// ��������� ����� ��� �������
					ksEntity entityConstrPlane3Point = (ksEntity)part.NewEntity((short)Obj3dType.o3d_plane3Points); 
					if (entityConstrPlane3Point != null)
					{
						ksPlane3PointsDefinition constrPlane3Point = (ksPlane3PointsDefinition)entityConstrPlane3Point.GetDefinition();
						if (constrPlane3Point != null)
						{
							constrPlane3Point.SetPoint(1, entityCollection.GetByIndex(0));
							constrPlane3Point.SetPoint(2, entityCollection.GetByIndex(1));
							constrPlane3Point.SetPoint(3, entityCollection.GetByIndex(2));
							entityConstrPlane3Point.Create();
						}
					}
					kompas.ksMessage("��������� ����� ��� �������");
				}
			}
		}


		// �������� ��������� ���������, ��� �� ���� ���������� � ��������� ��� ����� � ��������
		void CreateConstrElem()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				ksEntity entity = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				if (entity != null)
				{
					// ��������� ������� ��������� ���������
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entity.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 150;		// ���������� �� ������� ���������
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "XOY";		// �������� ��� ���������
						basePlane.Update();			// �������� ���������
      
						offsetDef.SetPlane(basePlane);			// ������� ���������
						offsetDef.direction = false;			// ����������� �������� �� ������� ���������
						entity.name = "��������� ���������";	// ��� ��� ��������� ���������
						entity.Create();						// ������� ��������� ��������� 
				
						kompas.ksMessage("������� ��������� ��������� ���������");
     
						offsetDef.offset = 50;				// ������� ���������� �� ������� ���������

						// ������� ������ ������� ���������
						basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
						basePlane.name = "YOZ";
						basePlane.Update();					// �������� ���������

						offsetDef.direction = true;			// ������� ����������� �������� ������������ ������� ���������
						offsetDef.SetPlane(basePlane); 
						entity.Update();					// �������� ���������

						// ������� ������ ������� ���������
						basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "XOY";
				
						kompas.ksMessage("�� ����������� ���������� �������� ���");

						// ��� �� ����������� ���� ����������
						ksEntity entityAxis = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axis2Planes);
						if (entityAxis != null)
						{
							ksAxis2PlanesDefinition axis2PlanesDef = (ksAxis2PlanesDefinition)entityAxis.GetDefinition();
							if (axis2PlanesDef != null)
							{
								axis2PlanesDef.SetPlane(1, entity);			// ������� ��������� 1
								axis2PlanesDef.SetPlane(2, basePlane);		// ������� ��������� 2
								entityAxis.name = "��� �� ���� ����������";	// ��� ��� ���
								entityAxis.Create();						// ������� ���  

								kompas.ksMessage("�������� ���� �� ������� ���������� ��� ���������� ���");
						
								// ������� ������ ������� ���������
								basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);
								basePlane.name = "XOZ";
						
								axis2PlanesDef.SetPlane(2, basePlane);	// ������� ��������� 2
								entityAxis.Update();

								kompas.ksMessage("����� ��������� ��������� � ����������� ��� \n �������� ��������� ��� ����� 45");

								ksEntity entityAnglePlane = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeAngle); 
								if (entityAnglePlane != null)
								{
									// ��������� ������� ��������� ��� ����� � ������ ���������
									ksPlaneAngleDefinition planeAngleDef = (ksPlaneAngleDefinition)entityAnglePlane.GetDefinition();
									if (planeAngleDef != null)
									{
										planeAngleDef.angle = 45;			// ���� ������� � ������� ���������
										planeAngleDef.SetPlane(entity);		// ������� ���������
										planeAngleDef.SetAxis(entityAxis);	// ������� ���
										entityAnglePlane.name = "��������� ��� ����� � ������ ���������";
										entityAnglePlane.Create();			// ������� ��������� ��� ����� 

										kompas.ksMessage("������� ���� �� ������� ����������");

										planeAngleDef.SetPlane(basePlane);	// ������� ���������
										entityAnglePlane.Update();			// �������� ��������� ���������
									}
								}
							}
						}
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
