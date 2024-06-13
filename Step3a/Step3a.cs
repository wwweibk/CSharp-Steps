
using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6Constants;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step3a - ������� - a

	// 1.  ������                            - WorkContour
	// 2.  ����������� ����������            - TDemWork
	// 3.  ������� ����                      - DrawViewPointer
	// 4.  ������ �� �������                 - WorkStamp
	// 5.  �������                           - TableWork
	// 6.  ������������                      - DrawEquidistant
	// 7.  ������                            - DrawEllipse
	// 8.  ���������                         - DrawPolyline
	// 9.  Nurbs                             - DrawNurbs
	// 10. ������ �����                      - WorkTolerance
	// 11. ���������� �������������          - DrawSpecRough
	// 12. ������� ��������� ������� ������� - DrawInsFragment1
	// 13. ������� ���������� ���������      - DrawInsFragment2

	public class Step3a
	{
		private KompasObject kompas = null;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)]
		public string GetLibraryName()
		{
			return "Step3a - ������� - a";
		}

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject)kompas_;

			if (kompas != null)
			{
				ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1:		WorkContour(doc);		break; // ������
						case 2:		TDemWork(doc);			break; // ����������� ����������
						case 3:		DrawViewPointer(doc);	break; // ������� ����
						case 4:		WorkStamp(doc);			break; // ������ �� �������
						case 5:		TableWork(doc);			break; // �������
						case 6:		DrawEquidistant(doc);	break; // ������������
						case 7:		DrawEllipse(doc);		break; // ������
						case 8:		DrawPolyline(doc);		break; // ���������
						case 9:		DrawNurbs(doc);			break; // nurbs
						case 10:	WorkTolerance(doc);		break; // ������ �����
						case 11:	DrawSpecRough(doc);		break; // ���������� �������������
						case 12:	DrawInsFragment1(doc);	break; // ������� ��������� ������� �������
						case 13:	DrawInsFragment2(doc);	break; // ������� ���������� ���������
					}
				}
			}
		}

		[return: MarshalAs(UnmanagedType.BStr)]
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "������";
					command = 1;
					break;
				case 2:
					result = "����������� ����������";
					command = 2;
					break;
				case 3:
					result = "������� ����";
					command = 3;
					break;
				case 4:
					result = "��������� �����";
					command = 4;
					break;
				case 5:
					result = "�������";
					command = 5;
					break;
				case 6:
					result = "������������";
					command = 6;
					break;
				case 7:
					result = "������";
					command = 7;
					break;
				case 8:
					result = "���������";
					command = 8;
					break;
				case 9:
					result = "Nurbs";
					command = 9;
					break;
				case 10:
					result = "������ �����";
					command = 10;
					break;
				case 11:
					result = "����������� �������������";
					command = 11;
					break;
				case 12:
					result = "������� ��������� ������� �������";
					command = 12;
					break;
				case 13:
					result = "������� ���������� ���������";
					command = 13;
					break;
				case 14:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}

		// ��������� ������
		void WorkContour(ksDocument2D doc)
		{
			if (doc.ksContour(1) == 1)
			{
				doc.ksLineSeg(20, 30, 50, 30, 1);
				doc.ksArcByPoint(50, 20, 10, 50, 30, 50, 10, -1, 1);
				//��������� ������
				doc.ksContour(2);
				doc.ksLineSeg(50, 10, 20, 10, 1);
				doc.ksArcByPoint(20, 20, 10, 20, 10, 20, 30, -1, 1);
				doc.ksEndObj();
				reference _contour = doc.ksEndObj();

				doc.ksLightObj(_contour, 1);
				kompas.ksMessage("������");
				doc.ksLightObj(_contour, 1);
				reference g = doc.ksNewGroup(0);
				doc.ksEndGroup();
				doc.ksAddObjGroup(g, _contour);
				doc.ksMoveObj(g, 10, 10);
				kompas.ksMessage("�������� ������");
			}
		}

		// ���������� ����������� ����������
		void TDemWork(ksDocument2D doc)
		{
			ksDynamicArray pGab = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.RECT_ARR);

			//�������� ������������� � ���� ���������� �����
			ksRectParam par = (ksRectParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectParam);
			ksMathPointParam pBot = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksMathPointParam pTop = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (pGab != null && par != null && pBot != null && pTop != null)
			{
				pBot.Init();
				pTop.Init();

				pTop.x = 415;
				pTop.y = 80;
				par.SetpTop(pTop);
				pBot.x = 230;
				pBot.y = 65;
				par.SetpBot(pBot);
				pGab.ksAddArrayItem(-1, par);

				pTop.x = 230;
				pTop.y = 60;
				par.SetpTop(pTop);
				pBot.x = 45;
				pBot.y = 15;
				par.SetpBot(pBot);
				pGab.ksAddArrayItem(-1, par);

				if (doc.ksOpenTechnicalDemand(pGab, 0) == 1)
				{
					ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
					if (itemParam != null)
					{
						itemParam.Init();

						ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
						if (itemFont != null)
						{
							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "1111111";
							doc.ksTextLine(itemParam);

							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "2222222";
							doc.ksTextLine(itemParam);

							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "3333333";
							doc.ksTextLine(itemParam);

							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "4444444";
							doc.ksTextLine(itemParam);

							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "5555555";
							doc.ksTextLine(itemParam);

							itemFont.Init();
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "6666666";
							doc.ksTextLine(itemParam);
						}
					}

					doc.ksCloseTechnicalDemand();
				}
			}
		}

		// �������� �������
		void TableWork(ksDocument2D doc)
		{
			doc.ksTable();
			doc.ksLineSeg(50, 50, 90, 50, 1);
			doc.ksLineSeg(50, 40, 90, 40, 1);
			doc.ksLineSeg(50, 30, 90, 30, 1);
			doc.ksLineSeg(50, 50, 50, 30, 1);
			doc.ksLineSeg(70, 50, 70, 30, 1);
			doc.ksLineSeg(90, 50, 90, 30, 1);

			doc.ksText(52, 48, 0, 5, 1, 0, "1");
			doc.ksText(72, 48, 0, 5, 1, 0, "2");
			doc.ksText(52, 38, 0, 5, 1, 0, "3");
			doc.ksText(72, 38, 0, 5, 1, 0, "4");
			doc.ksEndObj();
		}

		// �������� ������� ����
		void DrawViewPointer(ksDocument2D doc)
		{
			ksViewPointerParam par = (ksViewPointerParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ViewPointerParam);
			if (par != null)
			{
				par.Init();
				par.x1 = 55;
				par.y1 = 50;   // ���������� �������( ������) �������
				par.x2 = 40;
				par.y2 = 50;   // ���������� �������� ����� �������
				par.xt = 40;
				par.yt = 52;   // ���������� ������
				par.type = 0;
				par.str = "A"; // �������

				reference p = doc.ksViewPointer(par); //��������� "������� ����"
				if (doc.ksExistObj(p) == 1)
					doc.ksLightObj(p, 1);
			}
		}

		// ���������� �������� �������
		void WorkStamp(ksDocument2D doc)
		{
			ksStamp stamp = (ksStamp)doc.GetStamp();
			if (stamp != null)
			{
				if (stamp.ksOpenStamp() == 1)
				{
					stamp.ksColumnNumber(2);

					ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
					if (itemParam != null)
					{
						itemParam.Init();

						ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
						if (itemFont != null)
						{
							itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
							itemParam.s = "1111111";
							doc.ksTextLine(itemParam);
						}
					}

					stamp.ksCloseStamp();
				}
			}
		}

		// ������ �����
		void WorkTolerance(ksDocument2D doc)
		{
			ksToleranceParam par = (ksToleranceParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ToleranceParam);
			ksMathPointParam parPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (par != null && parPoint != null)
			{

				par.Init();
				parPoint.Init();

				ksDynamicArray branchArr = (ksDynamicArray)par.GetBranchArr();
				ksToleranceBranch tolBran = (ksToleranceBranch)kompas.GetParamStruct((short)StructType2DEnum.ko_ToleranceBranch);
				tolBran.Init();
				if (branchArr != null && tolBran != null)
				{

					ksDynamicArray arr = (ksDynamicArray)tolBran.GetpMathPoint();
					if (arr != null)
					{
						// ��������� ��������� 1-�� ����
						parPoint.x = 40;
						parPoint.y = 10;
						arr.ksAddArrayItem(-1, parPoint);
						tolBran.arrowType = 2;
						tolBran.tCorner = 1;
						branchArr.ksAddArrayItem(-1, tolBran);

						// ��������� ��������� 2-�� ����
						arr.ksClearArray();
						parPoint.x = 100;
						parPoint.y = 50;
						arr.ksAddArrayItem(-1, parPoint);
						parPoint.x = 100;
						parPoint.y = 10;
						arr.ksAddArrayItem(-1, parPoint);
						tolBran.arrowType = 1;
						tolBran.tCorner = 5;
						branchArr.ksAddArrayItem(-1, tolBran);

						par.x = 40;
						par.y = 40;
						par.type = 0;

						// ������ ����� ��������� ������
						if (doc.ksTolerance(par) == 1)
						{
							ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
							if (itemParam != null)
							{
								itemParam.Init();

								ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
								if (itemFont != null)
								{
									doc.ksColumnNumber(1);
									itemFont.SetBitVectorValue(ldefin2d.SPECIAL_SYMBOL, true);
									itemParam.type = ldefin2d.SPECIAL;
									itemParam.iSNumb = 26;
									doc.ksTextLine(itemParam);

									itemParam.Init();
									doc.ksColumnNumber(2);
									itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
									itemParam.s = "2222";
									doc.ksTextLine(itemParam);

									doc.ksColumnNumber(3);
									itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
									itemParam.s = "2222";
									doc.ksTextLine(itemParam);

									itemParam.Init();
									doc.ksColumnNumber(11);
									itemFont.SetBitVectorValue(ldefin2d.SPECIAL_SYMBOL, true);
									itemParam.type = ldefin2d.SPECIAL;
									itemParam.iSNumb = 23;
									doc.ksTextLine(itemParam);

									itemParam.Init();
									doc.ksColumnNumber(12);
									itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
									itemParam.s = "444";
									doc.ksTextLine(itemParam);
									doc.ksColumnNumber(13);
									itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
									itemParam.s = "555";
									doc.ksTextLine(itemParam);
								}
							}
						}

						reference p = doc.ksEndObj(); //�������� ��������� �� ������ �����
						doc.ksLightObj(p, 1);
					}
					arr.ksDeleteArray();
					branchArr.ksDeleteArray();
				}
			}
		}

		// ��������� ������������
		void DrawEquidistant(ksDocument2D doc)
		{
			ksEquidistantParam par = (ksEquidistantParam)kompas.GetParamStruct((short)StructType2DEnum.ko_EquidParam);
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (par != null && info != null)
			{
				par.side = 2;     // �������, � ����� ������� ������� ������������
				// 0-����� �� �����������, 1-������ �� �����������, 2-� ���� ������
				par.cutMode = false;  // ��� ������ ����� �������
				// 0-����� ������, 1- ����� �����
				par.degState = false; // ���� ���������� ����������� ��������� ������������
				// 0-����������� �������� ���������, 1-����������� �������� ���������
				par.radRight = 5; // ������ ������������
				par.radLeft = 3;  // ������ ������������
				par.style = 1;    // ��� �����
				info.commandsString = "������� ������";
				double x = 0;
				double y = 0;
				// ������ ������
				int j;
				reference p1;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						par.geoObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(par.geoObj) == 1)
						{
							if ((p1 = doc.ksEquidistant(par)) != 0)
							{ // ��������� ������������

								doc.ksLightObj(p1, 1);
								kompas.ksMessage("������������");
								doc.ksLightObj(p1, 0);
							}
							else
								kompas.ksMessageBoxResult();
						}
						else
							kompas.ksError("������ �� ������");
					}
				} while (j != 0);
			}
		}

		// ������� ������
		void DrawEllipse(ksDocument2D doc)
		{
			ksEllipseParam par = (ksEllipseParam)kompas.GetParamStruct((short)StructType2DEnum.ko_EllipseParam);
			if (par != null)
			{
				par.Init();
				par.xc = 50;
				par.yc = 40;
				par.A = 20;
				par.B = 10;
				par.style = 1;
				reference p = doc.ksEllipse(par);
				doc.ksLightObj(p, 1);
				kompas.ksMessage("������");
				doc.ksLightObj(p, 0);
			}
		}

		// �������� ���������
		void DrawPolyline(ksDocument2D doc)
		{
			//������ �������� ��������� ����� ��������
			ksPolylineParam par = (ksPolylineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PolylineParam);
			ksMathPointParam pr = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (par != null && pr != null)
			{
				par.Init();
				pr.Init();

				ksDynamicArray arr = (ksDynamicArray)par.GetpMathPoint();
				if (arr != null)
				{
					pr.x = 10;
					pr.y = 10;
					arr.ksAddArrayItem(-1, pr);
					pr.x = 20;
					pr.y = 20;
					arr.ksAddArrayItem(-1, pr);
					pr.x = 30;
					pr.y = 10;
					arr.ksAddArrayItem(-1, pr);
					pr.x = 40;
					pr.y = 20;
					arr.ksAddArrayItem(-1, pr);

					par.style = 2;
					reference p = doc.ksPolylineByParam(par);
					doc.ksLightObj(p, 1);
					kompas.ksMessage("���������");
					doc.ksLightObj(p, 0);

					arr.ksDeleteArray();
				}
			}
		}

		// ������� Nurbs - ������
		void DrawNurbs(ksDocument2D doc)
		{
			ksNurbsPointParam par = (ksNurbsPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_NurbsPointParam);
			if (par != null)
			{
				par.Init();
				//��������� Nurbs ������  ��� ��������� ������
				doc.ksNurbs(3, false, 1);
				par.x = 0;
				par.y = 0;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				par.x = 20;
				par.y = 20;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				par.x = 50;
				par.y = 10;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				par.x = 70;
				par.y = 20;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				par.x = 100;
				par.y = 0;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				par.x = 50;
				par.y = -50;
				par.weight = 1;
				doc.ksNurbsPoint(par);
				reference p = doc.ksEndObj();
				doc.ksLightObj(p, 1);
				kompas.ksMessage("NURBS");
				doc.ksLightObj(p, 0);
			}
		}

		// ������� �������� ���������
		void DrawInsFragment1(ksDocument2D doc)
		{
			ksFragment frg = (ksFragment)doc.GetFragment();
			if (frg != null)
			{
				// ��������� �������� ��� �������
				reference pDefFrg = frg.ksFragmentDefinition(@"c:\1.frw", // ��� ����� ���������
					"frw1",      // ��� �������
					1);         // ��� ������� 1 - ������� �������
				if (pDefFrg > 0)
				{
					ksPlacementParam par = (ksPlacementParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam);
					if (par != null)
					{
						par.angle = 45;
						par.scale_ = 2;
						par.xBase = 30;
						par.yBase = 40;
						// ������� ������ "������� ���������"
						reference p = frg.ksInsertFragment(pDefFrg,     // ��������� �����������  ���������
							false,       // ��� ���������� �� ����� false - �� ���� ���� 
							par);       // ��������� ��������
						doc.ksLightObj(p, 1);
						kompas.ksMessage("������� �������� ���������");
						doc.ksLightObj(p, 0);
					}
				}
			}
		}

		// ������� ���������� ���������
		void DrawInsFragment2(ksDocument2D doc)
		{
			// ��������� �������� ��� �������
			ksFragment frg = (ksFragment)doc.GetFragment();
			if (frg != null)
			{
				reference pDefFrg;
				// ��������� ��������� ��������
				if (frg.ksLocalFragmentDefinition("local") == 1)
				{
					doc.ksLineSeg(0, 0, 10, 0, 1);
					doc.ksLineSeg(0, 0, 0, 10, 1);
					doc.ksArcByPoint(0, 0, 10, 10, 0, 0, 10, -1, 1);
					pDefFrg = frg.ksCloseLocalFragmentDefinition();
					if (pDefFrg > 0)
					{
						ksPlacementParam par = (ksPlacementParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam);
						if (par != null)
						{
							par.angle = 45;
							par.scale_ = 2;
							par.xBase = 30;
							par.yBase = 40;
							// ������� ������ "������� ���������"
							reference p = frg.ksInsertFragment(pDefFrg,     // ��������� �����������  ���������
								false,       // ��� ���������� �� ����� false - �� ���� ���� 
								par);       // ��������� ��������
							doc.ksLightObj(p, 1);
							kompas.ksMessage("������� ���������� ���������");
							doc.ksLightObj(p, 0);
						}
					}
				}
			}
		}

		void DrawSpecRough(ksDocument2D doc)
		{
			ksSpecRoughParam par = (ksSpecRoughParam)kompas.GetParamStruct((short)StructType2DEnum.ko_SpecRoughParam);
			if (par != null)
			{
				par.Init();
				par.t = true;
				par.s = "Rz40";
				par.sign = 2;
				par.style = 1;
				doc.ksSpecRough(par);
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
