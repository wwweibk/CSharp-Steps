using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6Constants;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step3 - �������

	// 1.  ������� �������� - WorkDocument
	// 2.  ����             - DrawView
	// 3.  ����             - DrawLayer
	// 4.  ������           - DrawGroup
	// 5.  ������� ������   - WorkNameGroup
	// 6.  �������          - DrawLineSeg
	// 7.  ����             - DrawArc
	// 8.  �����            - DrawLine
	// 9.  ����������       - DrawCircle
	// 10. �����            - DrawPoint
	// 11. Bezier-�������   - DrawBezier
	// 12. ���������        - DrawHatch
	// 13. �����            - DrawText

  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step3
	{
		private KompasObject kompas;
		private ksDocument2D doc;

		// ������ � ����������
		private void WorkDocument()
		{
			doc = (ksDocument2D)kompas.Document2D();
			ksDocumentParam docPar = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
			ksDocumentParam docPar1 = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

			if ((docPar != null) & (docPar1 != null))
			{
				docPar.fileName = @"c:\2.cdw";
				docPar.comment = "Create document";
				docPar.author = "User";
				docPar.regime = 0;
				docPar.type = (short)DocType.lt_DocSheetStandart;
				ksSheetPar shPar = (ksSheetPar)docPar.GetLayoutParam();
				if (shPar != null)
				{
					shPar.shtType = 1;
					shPar.layoutName = string.Empty;
					ksStandartSheet stPar = (ksStandartSheet)shPar.GetSheetParam();
					if (stPar != null)
					{
						stPar.format = 3;
						stPar.multiply = 1;
						stPar.direct = false;
					}
				}
				// ������� ��������: ����, ������ �3, ������������� �������������
				// � � ��������� ������� 1
				doc.ksCreateDocument(docPar);

				int number = 0;
				ksViewParam par = (ksViewParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
				if (par != null)
				{
					number = 2;
					par.Init();
					par.x = 10;
					par.y = 20;
					par.scale_ = 0.5;
					par.angle = 45;
					par.color = Color.FromArgb(10, 20, 10).ToArgb();
					par.state = ldefin2d.stACTIVE;
					par.name = "User view";
					// � ��������� �������� ��� � ������� 2, ��������� 0.5, ��� ����� 45 ��
					doc.ksCreateSheetView(par, ref number);

					// �������� ���� � ������� 5
					doc.ksLayer(5);

					doc.ksLineSeg(20, 10, 40, 10, 1);
					doc.ksLineSeg(40, 10, 40, 30, 1);
					doc.ksLineSeg(40, 30, 20, 30, 1);
					doc.ksLineSeg(20, 30, 20, 10, 1);

					kompas.ksMessage("����������");

					// �������� ��������� ���������
					doc.ksGetObjParam(doc.reference, docPar1, ldefin2d.ALLPARAM);
					ksSheetPar shPar1 = (ksSheetPar)docPar1.GetLayoutParam();
					if (shPar1 != null)
					{
						ksStandartSheet stPar1 = (ksStandartSheet)shPar.GetSheetParam();
						if (stPar1 != null)
						{
							short direct = 0;
							if (stPar1.direct)
								direct = 1;
							string buf = string.Format("Type = {0}, f = {1}, m = {2}, d = {3}", docPar1.type, stPar1.format, stPar1.multiply, direct);
							kompas.ksMessage(buf);
						}
					}

					kompas.ksMessage(docPar1.fileName);
					kompas.ksMessage(docPar1.comment);
					kompas.ksMessage(docPar1.author);

					// C������� ��������
					doc.ksSaveDocument(string.Empty);
					//������� ��������
					doc.ksCloseDocument();
				}
			}
		}


		// ������� ���
		private void DrawView()
		{
			ksViewParam par = (ksViewParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
			if (par != null)
			{
				int number = 5;
				par.Init();
				par.x = 10;
				par.y = 20;
				par.scale_ = 0.5;
				par.angle = 45;
				par.color = Color.FromArgb(10, 20, 10).ToArgb();
				par.state = ldefin2d.stACTIVE;
				par.name = "User view";

				reference v = doc.ksCreateSheetView(par, ref number);
				number = doc.ksGetViewNumber(v);
				string buf = string.Format("������� ���: ref = {0}, number = {1}", v, number);
				kompas.ksMessage(buf);

				reference gr = doc.ksNewGroup(0);
				doc.ksLineSeg(20, 10, 20, 30, 1);
				doc.ksLineSeg(20, 30, 40, 30, 1);
				doc.ksLineSeg(40, 30, 40, 10, 1);
				doc.ksLineSeg(40, 10, 20, 10, 1);
				int res = doc.ksEndGroup();

				doc.ksAddObjGroup(gr, v);
				kompas.ksMessage("�������� ��� � ������");
				kompas.ksMessageBoxResult();

				reference p = doc.ksLineSeg(10, 10, 30, 30, 0);
				doc.ksAddObjGroup(gr, p);
				kompas.ksMessage("�������� ������� � ������");
				kompas.ksMessageBoxResult();

				doc.ksRotateObj(gr, 0, 0, -45);

				// ����� ��������� ����
				par.Init();
				doc.ksGetObjParam(v, par, ldefin2d.ALLPARAM);

				buf = string.Format("x = {0:.##}, y = {1:.##}, angl = {2:.##}, name = {3}, st = {4}", par.x, par.y, par.angle, par.name, par.state);
				kompas.ksMessage(buf);

				// ������� ������� ��������� ��� (����� 0)
				doc.ksOpenView(0);
				// ��������� ����: ������ ������
				ksLtVariant vart = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				if (vart != null)
				{
					vart.Init();
					vart.intVal = ldefin2d.stREADONLY;
					doc.ksSetObjParam(v, vart, ldefin2d.VIEW_LAYER_STATE);
				}
			}
		}


		// ������� ����
		private void DrawLayer()
		{
			int n = 0;
			if (kompas.ksReadInt("������� ����� ����", 1, 0, 255, ref n) != 1)
				return;

			// ������� ����, ���� ���������� �������
			reference lay = doc.ksLayer(n);
			doc.ksMtr(20, 15, 0, 1, 1);
			doc.ksLineSeg(-10, 0, 10, 0, 1);
			doc.ksLineSeg(10, 0, 10, 20, 1);
			doc.ksLineSeg(10, 20, -10, 20, 1);
			doc.ksLineSeg(-10, 20, -10, 0, 1);
			doc.ksDeleteMtr();

			// ���������� ����
			doc.ksLightObj(lay, 1);

			// �������� ����� ���� �� ��������� � ��������� �� ������
			int n1 = doc.ksGetLayerNumber(lay);
			reference l = doc.ksGetLayerReference(n1);

			string buf = string.Format("n = {0:.##}, n1 = {1:.##}, layer = {2:.##}, l = {3:.##}", n, n1, lay, l);
			kompas.ksMessage(buf);

			// ���������� ��������� ���� � ������� �� �������
			ksLayerParam par = (ksLayerParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LayerParam);
			ksLayerParam par1 = (ksLayerParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LayerParam);
			if ((par != null) & (par1 != null))
			{
				par.Init();
				par1.Init();
				par.color = (Color.FromArgb(0, 0, 255, 0).ToArgb());
				par.state = (ldefin2d.stACTIVE);
				par.name = ("�������");
				doc.ksLayer(0);

				if (doc.ksSetObjParam(l, par, ldefin2d.ALLPARAM) != 1)
					kompas.ksMessageBoxResult();
				else
				{
					doc.ksGetObjParam(l, par1, ldefin2d.ALLPARAM);
				
					buf = string.Format("col = {0:.##}, col1 = {1:.##}, name = {2}, name1 = {3}",
						par.color,
						par1.color,
						par.name,
						par1.name);
					kompas.ksMessage(buf);
				}

				// ����� ��������� ����
				doc.ksLightObj(lay, 0);

				// �������� ��������� ����
				ksLtVariant var = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				if (var != null)
				{
					var.Init();
					var.intVal = ldefin2d.stACTIVE;
					doc.ksSetObjParam(l, var, ldefin2d.VIEW_LAYER_STATE);
				}
			}
		}


		// ������ � �������
		private void DrawGroup()
		{
			reference p1 = doc.ksLineSeg(10, 10, 20, 10, 0);
			reference p2 = doc.ksLineSeg(10, 10, 10, 20, 0);

			// ������� ��������� ������ 1
			reference gr1 = doc.ksNewGroup(0);
			doc.ksEndGroup();

			// ������� ��������� ������ 2
			reference gr2 = doc.ksNewGroup(0);
			doc.ksEndGroup();

			doc.ksAddObjGroup(gr1, p1);
			doc.ksAddObjGroup(gr1, p2);

			doc.ksAddObjGroup(gr2, p1);
			doc.ksAddObjGroup(gr2, p2);

			kompas.ksMessage("������� ������");

			doc.ksMoveObj(gr1, 10, 0);
			kompas.ksMessage("�������� ������ �� 10 ��");

			doc.ksRotateObj(gr2, 20, 10, 45);
			kompas.ksMessage("��������� ������ �� 45 ��");

			doc.ksRotateObj(gr2, 20, 10, -45);
			kompas.ksMessage("��������� ������ �� -45 ��");

			doc.ksMoveObj(gr1, -10, 0);
			kompas.ksMessage("�������� ������ �� -10 ��");

			// �������� ������ 2 (������� ����������� �� ������)
			doc.ksClearGroup(gr2, false);
			// ������  ������ 2
			doc.ksDeleteObj(gr2);

			kompas.ksMessage("���������� gr");
			doc.ksLightObj(gr1, 1);

			kompas.ksMessage("��������� gr");
			doc.ksLightObj(gr1, 0);

			kompas.ksMessage("���������� el");
			doc.ksLightObj(p1, 1);
			kompas.ksMessage("��������� el");
			doc.ksLightObj(p1, 0);

			//������  ������ 1(������� �������� ����)
			doc.ksDeleteObj(gr1);
		}


		// ������ � ����������� �������
		private void WorkNameGroup()
		{
			//������� ������ ���������  �� ����� ������ �������
			reference gr = doc.ksNewGroup(0);
			reference p = doc.ksLineSeg(20, 20, 40, 20, 1);
			doc.ksLineSeg(40, 20, 40, 40, 1);
			doc.ksLineSeg(40, 40, 20, 40, 1);
			doc.ksLineSeg(20, 40, 20, 20, 1);
			doc.ksEndGroup();

			//������� ������ ��������� � �������
			//������� ������ �������� � ��������� �
			if (doc.ksSaveGroup(gr, "group1") != 1)
				return;
			reference gr1 = doc.ksGetGroup("group1");
			if (gr1 == 0)
				return;

			reference c = doc.ksCircle(30, 30, 10, 1);
			doc.ksAddObjGroup(gr1, c);

			doc.ksLightObj(gr1, 1);
			kompas.ksMessage("�������� ������ � ������� ������");
			doc.ksLightObj(gr1, 0);

			doc.ksExcludeObjGroup(gr1, p);

			doc.ksLightObj(gr1, 1);
			kompas.ksMessage("��������� ������ �� ������� ������");
			doc.ksLightObj(gr1, 0);
		}


		// ������� �������
		private void DrawLineSeg()
		{
			string buf = string.Empty;

			// ��������� �������
			// ������ ��������� ������� ���������
			doc.ksMtr(30, 20, 45, 1, 1);

			reference p = doc.ksLineSeg(30, 20, 60, 20, 1);

			// ����� ��������� �������
			ksLineSegParam par = (ksLineSegParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LineSegParam);
			if (par != null)
			{
				int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);

				buf = string.Format("t = {0}, x1 = {1:.#}, y1 = {2:.#}, x2 = {3:.#}, y2 = {4:.#}, tl = {5}",
					t, par.x1, par.y1,
					par.x2, par.y2,
					par.style);

				kompas.ksMessage(buf);

				// �������� ��������� �������
				par.x1 = 0;
				par.y1 = 0;
				par.x2 = 30;
				par.y2 = 60;
				par.style = 2;

				if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
					kompas.ksMessage("�������� ������");
				else
					kompas.ksMessageBoxResult();
			}

			// ��������� ��������� ������� ���������
			doc.ksDeleteMtr();
		}


		// ������� ����
		private void DrawArc()
		{
			string buf = string.Empty;
			// ��������� ����
			doc.ksMtr(10, 10, 0, 1, 1);
			reference p = doc.ksArcByAngle(30, 20, 20, 45, 135, 1, 1);

			// ����� ��������� ���� �� �����
			ksArcByAngleParam par = (ksArcByAngleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ArcByAngleParam);
			ksArcByPointParam par1 = (ksArcByPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ArcByPointParam);
			if ((par != null) && (par1 != null))
			{
				int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);

				buf = string.Format("t = {0}, xc = {1:.#}, yc = {2:.#}, rad = {3:.#}, \n a1 = {4:.#}, a2 = {5:.#}, napr = {6} , tl = {7}",
					t, par.xc, par.yc,
					par.rad, par.ang1,
					par.ang2, par.dir,
					par.style);

				kompas.ksMessage(buf);

				// �������� ��������� ���� �� ������
				par1.xc = 40;
				par1.yc = 30;
				par1.rad = 10;
				par1.dir = 1;
				par1.style = 2;
				par1.x1 = 50;
				par1.y1 = 30;
				par1.x2 = 40;
				par1.y2 = 20;

				if (doc.ksSetObjParam(p, par1, 1) == 1)
					kompas.ksMessage("�������� ������");
				else
					kompas.ksMessageBoxResult();
			}
			doc.ksDeleteMtr();
		}


		// ������� ��������������� �����
		private void DrawLine()
		{
			doc.ksMtr(0, 0, 45, 1, 1);
			string buf = string.Empty;
			// ��������� ��������������� �����
			reference p = doc.ksLine(30, 20, 0);

			// ����� ��������� ��������������� �����
			ksLineParam par = (ksLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LineParam);
			if (par != null)
			{
				par.Init();
				int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);
				buf = string.Format("t = {0}, x = {1:.#}, y = {2:.#}, alf = {3:.#}",
					t, par.x,
					par.y, par.angle);

				kompas.ksMessage(buf);

				// �������� ��������� ��������������� �����
				par.x = 0;
				par.y = 0;
				par.angle = 90;

				if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
					kompas.ksMessage("�������� ������");
				else
					kompas.ksMessageBoxResult();
			}
			doc.ksDeleteMtr();
		}


		// ������� ����������
		private void DrawCircle()
		{
			doc.ksMtr(0, 0, 0, 2, 2);
			string buf = string.Empty;
			// ��������� ����������
			reference p = doc.ksCircle(30, 20, 10, 1);

			// ����� ��������� ����������
			ksCircleParam par = (ksCircleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_CircleParam);
			if (par != null)
			{
				int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);
				buf = string.Format("t = {0}, xc = {1:.#}, yc = {2:.#}, rad = {3:.#}, tl = {4}",
					t, par.xc, par.yc,
					par.rad, par.style);
				kompas.ksMessage(buf);

				// �������� ��������� ����������
				par.xc = 0;
				par.yc = 0;
				par.rad = 20;
				par.style = 2;
				if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
					kompas.ksMessage("�������� ������");
				else
					kompas.ksMessageBoxResult();
			}
			doc.ksDeleteMtr();
		}


		// ������� �����
		private void DrawPoint()
		{
			// 0-�����, 1-�������, 2-�-�����, 3-�������, 
			// 4-�����������, 5-����������, 6-������,
			// 7-������������� �������

			doc.ksMtr(10, 10, 0, 1, 1);
			string buf = string.Empty;
			// ��������� �����
			reference p = doc.ksPoint(30, 40, 0);
			doc.ksPoint(40, 40, 1);
			doc.ksPoint(50, 40, 2);
			doc.ksPoint(60, 40, 3);
			doc.ksPoint(70, 40, 4);
			doc.ksPoint(80, 40, 5);
			doc.ksPoint(90, 40, 6);
			doc.ksPoint(100, 40, 7);

			// ����� ��������� �����
			ksPointParam par = (ksPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PointParam);
			if (par != null)
			{
				int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);
				buf = string.Format("t = {0}, x = {1:.#}, y = {2:.#}, style = {3}",
					t, par.x,
					par.y, par.style);
				kompas.ksMessage(buf);

				// �������� ��������� �����
				par.x = 20;
				par.y = 30;
				par.style = 7;

				if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
					kompas.ksMessage("�������� ������");
				else
					kompas.ksMessageBoxResult();
			}
			doc.ksDeleteMtr();
		}


		// ������� Bezier ������
		private void DrawBezier()
		{
			string buf = string.Empty;
			double[] x = new double[] { 0, 20, 50, 70, 100, 50 };
			double[] y = new double[] { 0, 20, 10, 20, 0, -50 };

			// ��������� Bezier ������
			doc.ksBezier(0, 1);
			for (int i = 0; i < 5; i++)
				doc.ksPoint(x[i], y[i], 0);
			reference p = doc.ksEndObj();

			// ����� ��������� Bezier �������
			ksMathPointParam pPar = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksBezierParam par = (ksBezierParam)kompas.GetParamStruct((short)StructType2DEnum.ko_BezierParam);
			if ((pPar != null) && (par != null))
			{
				par.Init();
				ksDynamicArray arr = (ksDynamicArray)par.GetMathPointArr();
				if (arr != null)
				{
					int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);

					int count = arr.ksGetArrayCount();
					buf = string.Format("t = {0}, count = {1}, close = {2}, tl = {3}",
						t, count, par.closed, par.style);
					kompas.ksMessage(buf);

					for (int i = 0; i < count; i++)
					{
						arr.ksGetArrayItem(i, pPar);
						buf = string.Format("x[{0}] = {1:##0.#}, y[{2}] = {3:##0.#}",
							i, pPar.x, i, pPar.y);
						kompas.ksMessage(buf);
					}

					// �������� ��������� Bezier �������
					arr.ksClearArray();
					// ��������� ���� ������
					for (int i = 0; i < 6; i++)
					{
						pPar.x = x[i];
						pPar.y = y[i];
						arr.ksAddArrayItem(-1, pPar);
					}

					par.style = 2;
					par.closed = 1;

					if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
						kompas.ksMessage("�������� ������");
					else
						kompas.ksMessageBoxResult();
					arr.ksDeleteArray();
				}
			}
		}


		// ������� ���������
		private void DrawHatch()
		{
			doc.ksMtr(30, 20, 0, 0.5, 0.5);
			string buf = string.Empty;

			// ��������� �������������� �������
			doc.ksLineSeg(20, 30, 70, 30, 2);
			doc.ksLineSeg(70, 30, 70, 80, 2);
			doc.ksLineSeg(70, 80, 20, 80, 2);
			doc.ksLineSeg(20, 80, 20, 30, 2);

			if (doc.ksHatch(0, 45, 2, 0, 0, 0) == 1)
			{
				doc.ksLineSeg(20, 30, 70, 30, 2);
				doc.ksLineSeg(70, 30, 70, 80, 2);
				doc.ksLineSeg(70, 80, 20, 80, 2);
				doc.ksLineSeg(20, 80, 20, 30, 2);
				reference p = doc.ksEndObj();

				// ����� ��������� ���������
				ksHatchParam par = (ksHatchParam)kompas.GetParamStruct((short)StructType2DEnum.ko_HatchParam);
				if (par != null)
				{
					par.Init();
					int t = doc.ksGetObjParam(p, par, ldefin2d.ALLPARAM);
					buf = string.Format("t = {0}, tip = {1}, angl = {2:0.#}, shag = {3:0.#}, \n width = {4:0.#}, x0 = {5:0.#}, y0 = {6:0.#}",
						t, par.style, par.ang,
						par.step, par.width,
						par.x, par.y);
					kompas.ksMessage(buf);
					doc.ksDeleteMtr();

					doc.ksMtr(0, 0, 0, 2, 2);

					// �������� ��������� ���������
					par.x = 0.8;

					if (doc.ksSetObjParam(p, par, ldefin2d.ALLPARAM) == 1)
						kompas.ksMessage("�������� ������");
					else
						kompas.ksMessageBoxResult();
					doc.ksDeleteMtr();
				}
			}
			else
				kompas.ksMessageBoxResult();
		}


		private void PrintPar1(ksTextLineParam par2, ksTextItemParam par3, ksDynamicArray arr2)
		{
			string buf = string.Empty;
			buf = string.Format("style = {0}", par2.style);
			kompas.ksMessage(buf);

			int count = arr2.ksGetArrayCount();
			for (int j = 0; j < count; j++)
			{
				arr2.ksGetArrayItem(j, par3);
				ksTextItemFont font = (ksTextItemFont)par3.GetItemFont();
				if (font != null)
				{
					buf = string.Format("j = {0}, h = {1:.#}, s = {2} \n fontName = {3}",
						j, font.height,
						par3.s, font.fontName);
					kompas.ksMessage(buf);
				}
			}
		}


		// ������� �����
		private void DrawText()
		{
			ksParagraphParam par = (ksParagraphParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ParagraphParam);
			ksTextParam par1 = (ksTextParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextParam);
			ksTextLineParam par2 = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam par3 = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			if ((par != null) && (par1 != null) && (par2 != null) && (par3 != null))
			{
				par1.Init();
				par2.Init();
				par3.Init();
				par.Init();
				par.x = 30;
				par.y = 30;
				par.height = 25;
				par.width = 20;

				doc.ksParagraph(par);

				// 4 ������ ������� ����� � ������� � �������� ����������
				ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
				if (itemParam != null)
				{
					itemParam.Init();

					ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
					if (itemFont != null)
					{
						itemFont.Init();
						itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
						itemParam.s = "111";
						doc.ksTextLine(itemParam);

						itemFont.Init();
						itemFont.SetBitVectorValue(ldefin2d.NUMERATOR, true);
						itemFont.SetBitVectorValue(ldefin2d.ITALIC_ON, true);
						itemParam.s = "55";
						doc.ksTextLine(itemParam);

						itemFont.Init();
						itemFont.SetBitVectorValue(ldefin2d.DENOMINATOR, true);
						itemParam.s = "77";
						doc.ksTextLine(itemParam);

						itemFont.Init();
						itemFont.SetBitVectorValue(ldefin2d.END_FRACTION, true);
						itemFont.SetBitVectorValue(ldefin2d.BOLD_OFF, true);
						itemFont.SetBitVectorValue(ldefin2d.ITALIC_OFF, true);
						itemParam.s = "4444";
						doc.ksTextLine(itemParam);
					}
				}

				reference p = doc.ksEndObj();

				// � ���������� ������ ������������� ��� ������� �������������� ����� :
				ksDynamicArray arr1 = (ksDynamicArray)par1.GetTextLineArr();   // ������ �� �������
				ksDynamicArray arr2 = (ksDynamicArray)par2.GetTextItemArr();   // ������ �� ����������� ������
				if ((arr1 != null) && (arr2 != null))
				{
					// ������� ��������� 1 -�� ������ ( ������ 0 )
					doc.ksGetObjParam(p, par2, 0);

					PrintPar1(par2, par3, arr2);
					kompas.ksMessageBoxResult();

					if (kompas.ksYesNo("�������� ��������� ������ ?") == 1)
					{
						// � ������ ������ ��������� ITALIC � BOLD � ������ ����
						arr2.ksGetArrayItem(0, par3);
						ksTextItemFont font = (ksTextItemFont)par3.GetItemFont();
						if (font != null)
						{
							font.SetBitVectorValue(ldefin2d.BOLD_OFF, true);
							font.SetBitVectorValue(ldefin2d.ITALIC_OFF, true);
							int clr = Color.FromArgb(0, 0, 255, 0).ToArgb();
							font.color = clr;
							arr2.ksSetArrayItem(0, par3);
							// ������� � ������ ������ ������
							doc.ksSetObjParam(p, par2, 0);
							// ������� ��������� 1 -�� ������ ( ������ 0 )  ��� ��������
							doc.ksGetObjParam(p, par2, 0);
							PrintPar1(par2, par3, arr2);
						}
					}
				}
			}
		}


		[return: MarshalAs(UnmanagedType.BStr)]
		public string GetLibraryName()
		{
			return "Step3 - �������";
		}


		[return: MarshalAs(UnmanagedType.BStr)]
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "������� ��������";
					command = 1;
					break;
				case 2:
					result = "����";
					command = 2;
					break;
				case 3:
					result = "����";
					command = 3;
					break;
				case 4:
					result = "������";
					command = 4;
					break;
				case 5:
					result = "������� ������";
					command = 5;
					break;
				case 6:
					result = "�������";
					command = 6;
					break;
				case 7:
					result = "����";
					command = 7;
					break;
				case 8:
					result = "�����";
					command = 8;
					break;
				case 9:
					result = "����������";
					command = 9;
					break;
				case 10:
					result = "�����";
					command = 10;
					break;
				case 11:
					result = "Bezier-�������";
					command = 11;
					break;
				case 12:
					result = "���������";
					command = 12;
					break;
				case 13:
					result = "�����";
					command = 13;
					break;
				case 14:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject)kompas_;

			if (kompas == null)
				return;

			if (command == 1)
				doc = (Document2D)kompas.Document2D();
			else
				doc = (Document2D)kompas.ActiveDocument2D();

			if (doc == null)
				return;

			switch (command)
			{
				case 1: WorkDocument(); break;	// ������� ��������
				case 2: DrawView(); break;	// ����
				case 3: DrawLayer(); break;	// ����
				case 4: DrawGroup(); break;	// ������
				case 5: WorkNameGroup(); break;	// ������� ������
				case 6: DrawLineSeg(); break;	// �������
				case 7: DrawArc(); break;	// ����
				case 8: DrawLine(); break;	// �����
				case 9: DrawCircle(); break;	// ����������
				case 10: DrawPoint(); break;	// �����
				case 11: DrawBezier(); break;	// Bezier-�������
				case 12: DrawHatch(); break;	// ���������
				case 13: DrawText(); break;	// �����
			}

			kompas.ksMessageBoxResult();
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
