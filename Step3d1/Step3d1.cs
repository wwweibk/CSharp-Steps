
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
	// ����� Step3d1 - ������� 3D
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step3d1
	{
		private KompasObject kompas;
		private ksDocument3D doc;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d1 - ������� 3D";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				doc = (ksDocument3D)kompas.ActiveDocument3D();
				if (doc == null || doc.reference == 0)
				{
					doc = (ksDocument3D)kompas.Document3D();
					doc.Create(true, true);

					doc.author = "Ethereal";
					doc.comment = "3D Steps - Step3d1";
					doc.UpdateDocumentParam();
				}

				switch (command)
				{
					case 1 : CreateExtrusion();			break; // ������� �������� ������������
					case 2 : OperationRotated();		break; // �������� ��������
					case 3 : OperationLoft();			break; // �������� �� ��������
					case 4 : CreateFilletAndChamfer();	break; // �������� ����� � ����������
					case 5 : CreateNextOper();			break; // �������� : ��������, �����, ������� ����������, ������� �������
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
					result = "�������� ������������";
					command = 1;
					break;
				case 2:
					result = "�������� ��������";
					command = 2;
					break;
				case 3:
					result = "������� �� ��������";
					command = 3;
					break;
				case 4:
					result = "�������� ����� � ����������";
					command = 4;
					break;
				case 5:
					result = "�������� : ��������, �����, ������� ����������, ������� �������";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// ������� ��� ������� �� �������� ������
		void ClearCurrentSketch(ksDocument2D sketchEdit)
		{
			// ������� �������� � ������ ��� ������������ ������� � ������          
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter != null)
			{
				if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, 0))
				{
					reference rf;
					if ((rf = iter.ksMoveIterator("F")) != 0) 
					{
						// �������� ��������� �� ������ ������� � ������
						// � ����� �������� ��������� �� ��������� ������� � ������ ���� �� ������ �� ����������
						do
						{
							if (sketchEdit.ksExistObj(rf) == 1)
								sketchEdit.ksDeleteObj(rf);	// ���� ������ ���������� ������� ��� 
						}
						while ((rf = iter.ksMoveIterator("N")) != 0); 
					}
					iter.ksDeleteIterator();	// ������ ��������
				}
			}
		}


		// ������� ������������, ������ � ������
		void CreateExtrusion()
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
						sketchDef.angle = 45;			// ���� �������� ������
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
								extrusionDef.directionType = (short)Direction_Type.dtNormal;         // ����������� ������������
								extrusionDef.SetSideParam(true,	// ������ �����������
									(short)End_Type.etBlind,	// ������ �� �������
									200, 0, false);
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10); // ������ ������ � ��� �����������
								extrusionDef.SetSketch(entitySketch);	// ����� �������� ������������
								entityExtr.Create();					// ������� ��������

								kompas.ksMessage("������� ��������� �������� ������������");

								extrusionDef.SetSideParam(false,	// �������� �����������
									(short)End_Type.etBlind,		// ������ �� �������
									150, 0, false);
								extrusionDef.directionType = (short)Direction_Type.dtBoth; // ����������� ������������ dtBoth - � ��� �����������
								entityExtr.Update();	// �������� ���������

								kompas.ksMessage("�������� �����");

								// ����� �������������� ������
								sketchEdit = (ksDocument2D)sketchDef.BeginEdit();

								// ������� �������� � ������ ��� ������������ ������� � ������
								ClearCurrentSketch(sketchEdit);
								// ������ � ����� ����������
								sketchEdit.ksCircle(0, 0, 100, 1);

								sketchDef.EndEdit();	// ���������� �������������� ������
								entitySketch.Update();	// �������� ��������� ������
								entityExtr.Update();	// �������� ��������� �������� ������������

								kompas.ksMessage("�������� �������������");

								// �������� ����� �����
								ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
								if (entitySketch2 != null)
								{
									// ��������� ������� ������
									ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
									if (sketchDef2 != null)
									{
										sketchDef2.SetPlane(basePlane);	// ��������� ���������
										sketchDef2.angle = 45;			// �������� ����� �� 45 ����.
										entitySketch2.Create();			// �������� �����

										// ��������� ��������� ������
										ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
										sketchEdit2.ksCircle(0, 0, 150, 1);
										sketchDef2.EndEdit();	// ���������� �������������� ������

										// �������� �������������
										ksEntity entityBossExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
										if (entityBossExtr != null)
										{
											ksBossExtrusionDefinition bossExtrDef = (ksBossExtrusionDefinition)entityBossExtr.GetDefinition();
											if (bossExtrDef != null)
											{
												ksExtrusionParam extrProp = (ksExtrusionParam)bossExtrDef.ExtrusionParam(); // ��������� ��������� ���������� ������������
												ksThinParam thinProp = (ksThinParam)bossExtrDef.ThinParam();      // ��������� ��������� ���������� ������ ������
												if (extrProp != null  && thinProp != null)
												{
													bossExtrDef.SetSketch(entitySketch2); // ����� �������� ������������

													extrProp.direction = (short)Direction_Type.dtNormal;      // ����������� ������������ (������)
													extrProp.typeNormal = (short)End_Type.etBlind;      // ��� ������������ (������ �� �������)
													extrProp.depthNormal = 100;         // ������� ������������

													thinProp.thin = false;              // ��� ������ ������

													entityBossExtr.Create();                // �������� ��������
												}
											}
										}

										// �������� ����� �����
										ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
										if (entitySketch3 != null)
										{
											// ��������� ������� ������
											ksSketchDefinition sketchDef3 = (ksSketchDefinition)entitySketch3.GetDefinition();
											if (sketchDef3 != null)
											{
												sketchDef3.SetPlane(basePlane);	// ��������� ���������
												sketchDef3.angle = 45;			// �������� ����� �� 45 ����.
												entitySketch3.Create();			// �������� �����

												// ��������� ��������� ������
												ksDocument2D sketchEdit3 = (ksDocument2D)sketchDef3.BeginEdit();
												// ������ ����� ����� - �������
												sketchEdit3.ksLineSeg(50,  50, -50,  50, 1);
												sketchEdit3.ksLineSeg(50, -50, -50, -50, 1);
												sketchEdit3.ksLineSeg(50, -50,  50,  50, 1);
												sketchEdit3.ksLineSeg(-50, -50, -50,  50, 1);
												sketchDef3.EndEdit();	// ���������� �������������� ������

												// ������� �������������
												ksEntity entityCutExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
												if (entityCutExtr != null)
												{
													ksCutExtrusionDefinition cutExtrDef = (ksCutExtrusionDefinition)entityCutExtr.GetDefinition();
													if (cutExtrDef != null)
													{
														cutExtrDef.SetSketch(entitySketch3);	// ��������� ����� ��������
														cutExtrDef.directionType = (short)Direction_Type.dtNormal; //������ �����������
														cutExtrDef.SetSideParam(true, (short)End_Type.etBlind, 50, 0, false);
														cutExtrDef.SetThinParam(false, 0, 0, 0);
													}

													entityCutExtr.Create();	// �������� �������� ��������� �������������
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}


		// �������� ����� � ����������
		void CreateFilletAndChamfer()
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
						sketchDef.angle = 45;			// ���� �������� ������
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
								extrusionDef.directionType = (short)Direction_Type.dtNormal;	// ����������� ������������
								extrusionDef.SetSideParam(true /*������ �����������*/,
									(short)End_Type.etBlind /*������ �� �������*/, 100, 0, false);
								extrusionDef.SetThinParam(false, 0, 0, 0);	// ��� ������ ������
								extrusionDef.SetSketch(entitySketch);		// ����� �������� ������������
								entityExtr.Create();						// ������� ��������

								ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);
								if (collect != null  && collect.SelectByPoint(0, 0, 0) && collect.GetCount()  != 0)
								{

									kompas.ksMessage("�������� ����������");

									ksEntity entityFillet = (ksEntity)part.NewEntity((short)Obj3dType.o3d_fillet);
									if (entityFillet != null)
									{
										ksFilletDefinition filletDef = (ksFilletDefinition)entityFillet.GetDefinition();
										if (filletDef != null)
										{
											filletDef.radius = 10;		// ������ ����������
											filletDef.tangent = false;	// ���������� �� �����������
											ksEntityCollection arr = (ksEntityCollection)filletDef.array();	// ������������ ������ ��������
											if (arr != null)
											{
												for (int i = 0, count = collect.GetCount(); i < count; i++)
													arr.Add(collect.GetByIndex(i));

												entityFillet.Create();
											}
										}
									}
								}

								ksEntityCollection collect2 = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);
								if (collect2 != null  && collect2.SelectByPoint(0, 0, 100) && collect2.GetCount() != 0) 
								{

									kompas.ksMessage("�������� �����");
									ksEntity entityChamfer = (ksEntity)part.NewEntity((short)Obj3dType.o3d_chamfer);
									if (entityChamfer != null)
									{
										ksChamferDefinition ChamferDef = (ksChamferDefinition)entityChamfer.GetDefinition();
										if (ChamferDef != null)
										{
											ChamferDef.tangent = false;
											ChamferDef.SetChamferParam(true, 10, 10);
											ksEntityCollection arr = (ksEntityCollection)ChamferDef.array();	// ������������ ������ ��������
											if (arr != null)
											{
												for (int i = 0, count = collect2.GetCount(); i < count; i++)
													arr.Add(collect2.GetByIndex(i));

												entityChamfer.Create();
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}


		// �������� ��������
		void OperationRotated()
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
						if (basePlane != null)
						{
							sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
							entitySketch.Create();			// �������� �����

							// ��������� ��������� ������
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
							if (sketchEdit != null)
							{
								sketchEdit.ksArcByAngle(0, 0, 20, -90, 90, 1, 1);
								sketchEdit.ksLineSeg(0, -20, 0, 20, 3);
								sketchDef.EndEdit();	// ���������� �������������� ������
							}

							ksEntity entityRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
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
									entityRotate.Create();				// ������� ��������
								}
							}
						}
					}
				}
				kompas.ksMessage("������� �������� ��������");

				ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch2 != null)
				{
					ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
					if (sketchDef2 != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						if (basePlane != null)
						{
							sketchDef2.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
							entitySketch2.Create();			// �������� �����
							// ��������� ��������� ������
							ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
							if (sketchEdit2 != null)
							{
								sketchEdit2.ksArcByAngle(15, 0, 10, -90, 90, 1, 1);
								sketchEdit2.ksLineSeg(15, -10, 15, 10, 3);
								sketchDef2.EndEdit();	// ���������� �������������� ������
							}

							ksEntity entityBossRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
							if (entityBossRotate != null)
							{
								ksBossRotatedDefinition bossRotateDef = (ksBossRotatedDefinition)entityBossRotate.GetDefinition();
								if (bossRotateDef != null)
								{
									bossRotateDef.directionType = (short)Direction_Type.dtNormal;
									bossRotateDef.SetSideParam(true, 360);
									bossRotateDef.SetSketch(entitySketch2);		// ����� �������� ��������
									entityBossRotate.Create();					// ������� ��������
								}
							}
						}
					}
				}
				kompas.ksMessage("�������� ������������ ���������");

				ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch3 != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch3.GetDefinition();
					if (sketchDef != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						if (basePlane != null)
						{
							sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
							entitySketch3.Create();			// �������� �����
							// ��������� ��������� ������
							ksDocument2D sketchEdit3 = (ksDocument2D)sketchDef.BeginEdit();
							if (sketchEdit3 != null)
							{
								sketchEdit3.ksArcByAngle(20, 0, 20, 90, 270, 1, 1);
								sketchEdit3.ksLineSeg(20, -20, 20, 20, 3);
								sketchDef.EndEdit();	// ���������� �������������� ������
							}

							ksEntity entityCutRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutRotated);
							if (entityCutRotate != null)
							{
								ksCutRotatedDefinition cutRotateDef = (ksCutRotatedDefinition)entityCutRotate.GetDefinition();
								if (cutRotateDef != null)
								{
									cutRotateDef.directionType = (short)Direction_Type.dtNormal;
									cutRotateDef.SetSideParam(true, 90);
									cutRotateDef.SetThinParam(true, (short)Direction_Type.dtBoth, 5, 7);	// ������ ������ � ��� �����������
									cutRotateDef.SetSketch(entitySketch3);	// ����� �������� ��������
									entityCutRotate.Create();				// ������� ��������
								}
							}
						}
					}
				}
				kompas.ksMessage("�������� ��������� ���������");
			}
		}


		// �������� �� ��������
		void OperationLoft()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// ����� ���������
			if (part != null)
			{
				// �������� �����
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// ������� ��������� ������� ��������� XOY
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						entitySketch.Create();			// �������� �����
						entitySketch.hidden = false;

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						if (sketchEdit != null)
							sketchEdit.ksCircle(0, 0, 4.5, 1);
						sketchDef.EndEdit();	// ���������� �������������� ������
					}
				}

				// �������� ��������� ���������, � � ��� �����
				ksEntity entityOffsetPlane = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane != null)
				{
					// ��������� ������� ��������� ���������
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 30;							// ���������� �� ������� ���������
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "��������� ���������";			// �������� ��� ���������

						offsetDef.SetPlane(basePlane);					// ������� ���������
						entityOffsetPlane.name = "��������� ���������";	// ��� ��� ��������� ���������
						entityOffsetPlane.hidden = true;
						entityOffsetPlane.Create();						// ������� ��������� ���������

						if (entitySketch2 != null)
						{
							ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch2.GetDefinition();
							if (sketchDef != null)
							{
								sketchDef.SetPlane(entityOffsetPlane);	// ��������� ��������� XOY ������� ��� ������
								entitySketch2.Create();					// �������� �����

								// ��������� ��������� ������
								ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
								sketchEdit.ksCircle(0, 0, 8, 1);
								sketchDef.EndEdit();					// ���������� �������������� ������
							}
						}
					}
				}

				// �������� ��������� ���������, � � ��� �����
				ksEntity entityOffsetPlane2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane2 != null  && entitySketch3 != null)
				{
					// ��������� ������� ��������� ���������
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane2.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 60;	// ���������� �� ������� ���������
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "��������� ���������";	// �������� ��� ���������
        
						offsetDef.SetPlane(basePlane);						// ������� ���������
						entityOffsetPlane2.name = "��������� ���������2";	// ��� ��� ��������� ���������
						entityOffsetPlane2.hidden = true;
						entityOffsetPlane2.Create();						// ������� ��������� ��������� 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch3.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane2);			// ��������� ��������� XOY ������� ��� ������
							entitySketch3.Create();							// �������� �����

							// ��������� ��������� ������
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
							sketchEdit.ksCircle(0, 0, 1.5, 1);
							sketchDef.EndEdit();							// ���������� �������������� ������                
						}
					}
				}

				// �������� ������� �������� �� ��������
				ksEntity entityBaseLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossLoft); 
				if (entityBaseLoft != null)
				{
					ksBossLoftDefinition baseLoft = (ksBossLoftDefinition)entityBaseLoft.GetDefinition();
					if (baseLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)baseLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch);
							entCol.Add(entitySketch2);
							entCol.Add(entitySketch3);
						}
						entityBaseLoft.name = "�����";
						entityBaseLoft.SetAdvancedColor(12345678, .8, .8, .8, .8, 1, .8);
						entityBaseLoft.Create();	// ������� ��������
					}
				}
				kompas.ksMessage("������� �������� �� ��������");

				// �������� ��������� ���������, � � ��� �����
				ksEntity entitySketch7 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch7 != null)
				{
					// ��������� ������� ��������� ���������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch7.GetDefinition();
					if (sketchDef != null)
					{
						sketchDef.SetPlane(entityOffsetPlane2);	// ��������� ��������� XOY ������� ��� ������
						entitySketch7.Create();					// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						sketchEdit.ksCircle(0, 0, 1.5, 1);
						sketchDef.EndEdit();					// ���������� �������������� ������                
					}
				}

				// �������� ��������� ���������, � � ��� �����
				ksEntity entityOffsetPlane3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch4 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane3 != null  && entitySketch4 != null)
				{
					// ��������� ������� ��������� ���������
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane3.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 120;					// ���������� �� ������� ���������
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "��������� ���������";	// �������� ��� ���������
        
						offsetDef.SetPlane(basePlane);			// ������� ���������
						entityOffsetPlane3.name = "��������� ���������";	// ��� ��� ��������� ���������
						entityOffsetPlane3.hidden = true;
						entityOffsetPlane3.Create();			// ������� ��������� ��������� 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch4.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane3);	// ��������� ��������� XOY ������� ��� ������
							entitySketch4.Create();					// �������� �����

							// ��������� ��������� ������
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
							sketchEdit.ksCircle(0, 0, 1.8, 1);
							sketchDef.EndEdit();					// ���������� �������������� ������                
						}
					}
				}


				// �������� �������� ������������ �� ��������
				ksEntity entityBossLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossLoft);
				if (entityBossLoft != null)
				{
					ksBossLoftDefinition bossLoft = (ksBossLoftDefinition)entityBossLoft.GetDefinition();
					if (bossLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)bossLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch7);
							entCol.Add(entitySketch4);
						}
						entityBossLoft.name = "�����";
						entityBossLoft.SetAdvancedColor(1234567890, .8, .8, .8, .8, 1, .8);

						entityBossLoft.Create();	// ������� ��������
					}
				}
				kompas.ksMessage("�������� ������������ �� ��������");

				// �������� ����� � ��� ��������� ��������� ���������
				ksEntity entitySketch5 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch5 != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch5.GetDefinition();
					if (sketchDef != null)
					{
						sketchDef.SetPlane(entityOffsetPlane3);	// ��������� ��������� XOY ������� ��� ������
						entitySketch5.Create();					// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						ksRectangleParam recPar = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
						recPar.Init();
						if (recPar != null)
						{
							recPar.x = -1.8;
							recPar.y = -.4;
							recPar.height = .8;
							recPar.width = 3.6;
							recPar.style = 1;
						}
						sketchEdit.ksRectangle(recPar, 0); 
						sketchDef.EndEdit();	// ���������� �������������� ������                
					}
				}

				// �������� ��������� ���������, � � ��� �����
				ksEntity entityOffsetPlane4 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch6 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane4 != null  && entitySketch6 != null)
				{
					// ��������� ������� ��������� ���������
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane4.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 110;	// ���������� �� ������� ���������
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "��������� ���������";	// �������� ��� ���������
        
						offsetDef.SetPlane(basePlane);	// ������� ���������
						entityOffsetPlane4.name = "��������� ���������";	// ��� ��� ��������� ���������
						entityOffsetPlane4.hidden = true;
						entityOffsetPlane4.Create();	// ������� ��������� ��������� 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch6.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane4);	// ��������� ��������� XOY ������� ��� ������
							entitySketch6.Create();					// �������� �����

							// ��������� ��������� ������
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
							ksRectangleParam recPar = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
							recPar.Init();
							if (recPar != null)
							{
								recPar.x = -1.8;
								recPar.y = -1.8;
								recPar.height = 3.6;
								recPar.width = 3.6;
								recPar.style = 1;
							}
							sketchEdit.ksRectangle(recPar, 0); 
							sketchDef.EndEdit();	// ���������� �������������� ������                
						}
					}
				}


				// �������� �������� ��������� �� ��������
				ksEntity entityCutLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutLoft); 
				if (entityCutLoft != null)
				{
					ksCutLoftDefinition cutLoft = (ksCutLoftDefinition)entityCutLoft.GetDefinition();
					if (cutLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)cutLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch5);
							entCol.Add(entitySketch6);
						}

						cutLoft.SetThinParam(true, (short)Direction_Type.dtNormal, 3, 0);
						entityCutLoft.name = "������� �����������";
						entityCutLoft.SetAdvancedColor(1234, .8, .8, .8, .8, 1, .8);

						entityCutLoft.Create();	// ������� ��������
					}
				}

				kompas.ksMessage("�������� ��������� �� ��������");		
			}
		}


		// �������� : ��������, �����, ������� ����������, ������� �������
		void CreateNextOper()
		{
			// �������� ����� ������
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);
			if (part != null)
			{
				// �������� ����� ��� ������� ��������
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// ������� ��������� ������� ������
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// ������� ��������� ������� ��������� XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// ��������� ��������� XOY ������� ��� ������
						sketchDef.angle = 0;			// ���� �������� ������
						entitySketch.Create();			// �������� �����

						// ��������� ��������� ������
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// ������ ����� ����� - �������
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 
						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// ���������� �������������� ������                

						kompas.ksMessage("������� �������� ������������");
						// ������� �������� ������������
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// ��������� ������� ������� �������� ������������
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); 
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;	// ����������� ������������
								extrusionDef.SetSideParam(true /*������ �����������*/,
									(short)End_Type.etBlind /*������ �� �������*/, 200, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10);	// ������ ������ � ��� �����������
								extrusionDef.SetSketch(entitySketch);	// ����� �������� ������������
								entityExtr.Create();					// ������� ��������
							}

							bool update = false;	// ���� update = true, �� ��������� �������� �������� 
							if (MessageBox.Show("������� �������� \"��������\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// ������� �������� ��������
								ksEntity entShell = (ksEntity)part.NewEntity((short)Obj3dType.o3d_shellOperation);
								if (entShell != null)
								{
									// ��������� ������� ������� �������� ������������
									ksShellDefinition incDef = (ksShellDefinition)entShell.GetDefinition();
									if (incDef != null)
									{
										ksEntityCollection entCol = (ksEntityCollection)incDef.FaceArray();               // ������� ������ ������ ��� �������� ��������
										ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face); // ������ ���� ������
										if (entCol != null  && collect != null)
										{
											incDef.thickness = 8;				// �������� ��������
											incDef.thinType = true;				// ����������� �������� ������
											collect.SelectByPoint(50, 0, 0);	// ������� ����� �� ������ ��� ���������� ������� ���� ������, ���������� ����� ��� �����
											entCol.Add(collect.GetByIndex(0));	// ������� � ������ ������ ��� �������� ����� � �������� = 0 
											collect.refresh();					// ������� ������
											entShell.Create();					// ������� ��������

											if (MessageBox.Show("�������� ��������� �������� \"��������\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
											{

												incDef.thickness = 25;				// �������� ��������
												incDef.thinType = false;			// ����������� �������� ������
												collect.SelectByPoint(60, 0, 10);	// ������� ����� �� ������ ��� ���������� ������� ���� ������, ���������� ����� ��� �����
												entCol.Add(collect.GetByIndex(0));	// ������� � ������ ������ ��� �������� ����� � �������� = 0 
												collect.refresh();					// ������� ������
												entShell.Update();					// ���������� ��������
												update = true;
											}
										}
									}
								}
							}

							if (MessageBox.Show("������� �������� \"�����\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// �������� �������� �����
								ksEntity entInc = (ksEntity)part.NewEntity((short)Obj3dType.o3d_incline);
								if (entInc != null)
								{
									// ��������� ������� ������� �������� ������
									ksInclineDefinition incDef = (ksInclineDefinition)entInc.GetDefinition();
									ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face); // ������ ���� ������
									if (incDef != null  && collect != null) 
									{
										incDef.direction = true;		// ����������� ������ - ������  
										incDef.angle = 3;				// ���� ������  
										incDef.SetPlane(basePlane);		// ������� ���������  
										ksEntityCollection entColInc = (ksEntityCollection)incDef.FaceArray();  // ������� ������ ������ ��� �������� �����
										if (entColInc != null)
										{
											collect.SelectByPoint(0, update ? 85 : 60, 10);	// ������� ����� �� ������ ��� ���������� ������� ���� ������, ���������� ����� ��� �����
											entColInc.Add(collect.GetByIndex(0));			// ������� � ������ ������ ��� �������� ����� � �������� = 0 
											collect.refresh();								// ������� ������
											entInc.Create();								// ������� ��������

											if (MessageBox.Show("�������� ��������� �������� \"�����\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
											{
												incDef.direction = false;							// ����������� ������ - ������
												incDef.angle = 25;									// ���� ������  
												collect.SelectByPoint(0, update ? -85 : -60, 10);	// ������� ����� �� ������ ��� ���������� ������� ���� ������, ���������� ����� ��� �����
												entColInc.Add(collect.GetByIndex(0));				// ������� � ������ ������ ��� �������� ����� � �������� = 0 
												collect.refresh();									// ������� ������
												entInc.Update();
											}
										}
									}
								}
							}

							if (MessageBox.Show("������� �������� \"������� ����������\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// �������� �������� ������� ����������
								ksEntity entityOpr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutByPlane);
								if (entityOpr != null)
								{
									// ��������� ������� �������� ������� ����������
									ksCutByPlaneDefinition incOpr = (ksCutByPlaneDefinition)entityOpr.GetDefinition(); // ��������� ��������
									if (incOpr != null)
									{
										// ������� ��������� ������� ��������� XOZ, ������� ����� �������� ���������� �������
										ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);
										incOpr.SetPlane(basePlaneXOZ);	// ������ ��������� �������
										incOpr.direction = false;		// ����������� ������� - ��������
										entityOpr.Create();				// ������� ��������

										if (MessageBox.Show("�������� ��������� �������� \"������� ����������\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
										{
											incOpr.direction = true;	// ����������� ������� - ������
											entityOpr.Update();			// ����������� ��������
										}
									}
								}
							}

							if (MessageBox.Show("������� �������� \"������� �������\"?", "��������� ����������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// �������� ����� ��� �������� ������� �������
								ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);  
								if (entitySketch2 != null)
								{
									// ��������� ������� ������
									ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
									if (sketchDef2 != null)
									{
										// ������� ��������� ������� ��������� YOZ
										ksEntity basePlaneYOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
										sketchDef2.SetPlane(basePlaneYOZ);	// ��������� ��������� yoz ������� ��� ������
										sketchDef2.angle = 45;				// ���� �������� ������
										entitySketch2.Create();				// �������� �����

										// ��������� ��������� ������
										ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit(); 
										// ������ ����� �����
										sketchEdit2.ksArcBy3Points(-200, 45, -150, 10, -50, 10, 1);
										sketchEdit2.ksLineSeg     (-200, 45, -300, 20, 1);
										sketchEdit2.ksLineSeg     (-50, 10,   60, 10, 1);
										sketchDef2.EndEdit();	// ���������� �������������� ������
              
										// ��������� �������� ������� �������
										ksEntity entityOpr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutBySketch);
										if (entityOpr != null)
										{
											// ��������� �������� ������� �������
											ksCutBySketchDefinition incOpr = (ksCutBySketchDefinition)entityOpr.GetDefinition(); 
											if (incOpr != null)
											{
												incOpr.SetSketch(entitySketch2);	// ������ �����
												incOpr.direction = true;			// ������ ����������� �������
												entityOpr.Create();					// ������� ��������
											}
										}
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
