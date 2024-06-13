using Kompas6API5;
using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step5 - ��������������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step5
	{
		KompasObject kompas = null;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step5 - ��������������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				ksDocument2D doc = (ksDocument2D) kompas.ActiveDocument2D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1  : DrawTransform(doc);         break; //������������� ������� �� �������
						case 2  : DrawCopy(doc);              break; //����������� �������
						case 3  : DrawSymmetry(doc);          break; //��������� �������
						case 4  : EditTolerance(doc);         break; //�������� ������� �����
						case 5  : EditTable(doc);             break; //�������� �������
						case 6  : EditStamp(doc);             break; //����� ������ ���� � ������������� �����
						case 7  : GetTextTT(doc);             break; //�������� ����� ��
						case 8	: ChangeTechnicalDemand(doc); break; //�������������� TT
						case 9	: ShowInsertFragment(doc);    break; //������� ���������
						case 10 : EditFragmentLibrary(doc);   break; //������ � ����������� ����������
						case 11 : ShowInsertFragment1(doc);   break; //������� ��������� ��������
					}
				}
			}
		}


		// ������������ ���� ����������
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "T������������ �������";
					command = 1;
					break;
				case 2:
					result = "�����  �������";
					command = 2;
					break;
				case 3:
					result = "��������� �������";
					command = 3;
					break;
				case 4:
					result = "�������� � �������������� ������� �����";
					command = 4;
					break;
				case 5:
					result = "�������� � �������������� �������";
					command = 5;
					break;
				case 6:
					result = "����� ������ ���� � ������������� �����";
					command = 6;
					break;
				case 7:
					result = "�������� ����� ��";
					command = 7;
					break;
				case 8:
					result = "������������� ��";
					command = 8;
					break;
				case 9:
					result = "������� ��������� �� ���������� ����������";
					command = 9;
					break;
				case 10:
					result = "������ � ����������� ����������";
					command = 10;
					break;
				case 11:
					result = "������� ��������� ��������";
					command = 11;
					break;
				case 12:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		void DrawTransform(ksDocument2D doc)
		{
			//�������������  ������� ������
			doc.ksMtr(-30, -30, 0, 1, 1);
			reference rf = doc.ksNewGroup(0);
			doc.ksLineSeg(30, 30, 60, 30, 1);
			doc.ksLineSeg(60, 30, 60, 60, 1);
			doc.ksLineSeg(60, 60, 30, 60, 1);
			doc.ksLineSeg(30, 60, 30, 30, 1);
			doc.ksHatch(0, 45, 2, 0, 0, 0);
			doc.ksLineSeg(30, 30, 60, 30, 1);
			doc.ksLineSeg(60, 30, 60, 60, 1);
			doc.ksLineSeg(60, 60, 30, 60, 1);
			doc.ksLineSeg(30, 60, 30, 30, 1);
			doc.ksEndObj();
			doc.ksEndGroup();
			doc.ksDeleteMtr();
  
			kompas.ksMessage("������� ������� 20, 20, 45, 2");

			doc.ksMtr(20, 20, 45, 2, 2);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			kompas.ksMessageBoxResult();
			kompas.ksMessage("������ �������");

			doc.ksMtr(-20, -20, 0, 1, 1);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			doc.ksMtr(0, 0, 0, 0.5, 0.5);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			doc.ksMtr(0, 0, -45, 1, 1);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			kompas.ksMessageBoxResult();
		}

		void DrawCopy(ksDocument2D doc)
		{
			ksViewParam par = (ksViewParam) kompas.GetParamStruct((short)StructType2DEnum.ko_ViewParam);

			if (par != null) 
			{
				par.Init();
				par.x = 20;
				par.y = 60;
				par.scale_ = 1 ;
				par.color = Color.FromArgb(0, 10,20,10).ToArgb();
				par.state = ldefin2d.stACTIVE;
				par.name = "user view";

				int number = 5;
				//������� ���
				reference v = doc.ksCreateSheetView(par, ref number);
				//������� ����
				doc.ksLayer(5);

				doc.ksLineSeg(20, 10, 20, 30, 1);
				doc.ksLineSeg(20, 30, 40, 30, 1);
				doc.ksLineSeg(40, 30, 40, 10, 1);
				doc.ksLineSeg(40, 10, 20, 10, 1);

				//�������� ��� (��� ���� ����� �������� � �������� �����������)
				doc.ksCopyObj(v, 20, 60, 40, 80, 1, 0);

				kompas.ksMessageBoxResult();
			}
		}

		void DrawSymmetry(ksDocument2D doc)
		{
			reference grp = doc.ksNewGroup(0);
			doc.ksLineSeg(20, 10, 20, 30, 1);
			doc.ksLineSeg(20, 30, 40, 30, 1);
			doc.ksLineSeg(40, 30, 40, 10, 1);
			doc.ksLineSeg(40, 10, 20, 10, 1);
			doc.ksEndGroup();

			doc.ksSymmetryObj(grp, 40, 10, 40, 20, "1");

			kompas.ksMessageBoxResult();
		}

		void EditTolerance(ksDocument2D doc)
		{
			// �������������� ������� �����
			reference pObj;

			ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info == null) 
				return;
			double x = 0, y = 0;
			info.prompt = "������� ������ �����" ;
			int cursor = doc.ksCursor(info, ref x ,ref y, 0);
			if (cursor != 0) 
			{
				if (doc.ksExistObj(pObj = doc.ksFindObj(x, y, 1000000)) == 1)
				{
					//������ ��� �������
					int type = doc.ksGetObjParam(pObj, 0, 0);   //��������� �� ����������� ������
					if (type == ldefin2d.TOLERANCE_OBJ) 
					{
						int numb = 0;
						string buf = string.Empty;
						//������� ������ ����� ��� ��������������
						doc.ksOpenTolerance(pObj);

						ksToleranceParam par = (ksToleranceParam) kompas.GetParamStruct((short)StructType2DEnum.ko_ToleranceParam);
						//��������� ������� �����
						doc.ksGetObjParam( pObj,	//��������� �� ����������� ������
							par,					//��������� �� ��������� ����������
							ldefin2d.ALLPARAM);		//��� ���������� ����������

						buf = string.Format("������� ����� = {0} ����� = {1} ������������ - {2}\nx = {3:0.##} y = {4:0.##}",
							par.tBase, par.style,
							par.type != 0 ? "������������" : "��������������",
							par.x, par.y);
						kompas.ksMessage(buf);

						ksTextLineParam par1 = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);

						if (par1 == null)
							return;
				
						par1.Init();
        
						//� ����� ����� ����� ��� ������������ ������
						while (doc.ksGetToleranceColumnText(ref numb, par1) != 0) 
						{
							buf = string.Format("numb = {0}\nstyle = {1}", numb, par1.style);
							kompas.ksMessage(buf);

							ksDynamicArray arrpTextItem = (ksDynamicArray) par1.GetTextItemArr();
							ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

							if (item == null || arrpTextItem == null)
								return;
          
							item.Init();

							for (int i = 0; i < arrpTextItem.ksGetArrayCount(); i ++) 
							{
								arrpTextItem.ksGetArrayItem(i, item);
								ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont();
								if (item.type == 0)
									buf = string.Format("���������� = {0} h = {1:0.##}\ns = {2}\nfontName = {3}\n��������� = {4}",
										i, textItemFont.height, item.s, textItemFont.fontName, textItemFont.bitVector);
								else
									buf = string.Format("���������� = {0} ��� = {1} ����� ��������� = {3}", i, item.type, item.iSNumb);
								item.s = "������ �����";
								kompas.ksMessage(buf);
							}
							arrpTextItem.ksClearArray();
							arrpTextItem.ksAddArrayItem(-1, item);
							doc.ksSetToleranceColumnText(numb, par1);
							arrpTextItem.ksDeleteArray();  //������� ������ ���������
						}
						//������� ���������
						par.x =  par.x + 10 ;
						par.y = par.y + 10 ;
						doc.ksSetObjParam(pObj,	//��������� �� ����������� ������
							par,				//��������� �� ��������� ����������
							ldefin2d.ALLPARAM);	//��� ���������� ����������
						doc.ksEndObj();			//������� ������ "������ �����"
					}
					else
						kompas.ksError("��� �� ������ �����");
				}
				else
					kompas.ksError("��� �������");
			}

			kompas.ksMessageBoxResult();
		}

		void EditTable(ksDocument2D doc)
		{
			// �������������� �������
			reference pObj;

			ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info == null) 
				return;
			double x = 0, y = 0;
			info.prompt = "������� �������";
			// ����� ������� �� �������
			int cursor = doc.ksCursor(info, ref x, ref y, 0);
			if (cursor != 0) 
			{
				if(doc.ksExistObj(pObj = doc.ksFindObj(x, y, 100000)) == 1)
				{
					//������ ��� �������
					int type = doc.ksGetObjParam(pObj, 0, 0);	//��������� �� ����������� ������
					//��������� ���������� ������  - �������
					if (type == ldefin2d.TABLE_OBJ) 
					{
						int numb = 0;
						//reference p;
						string buf = string.Empty;
						//������� ������� ��� ��������������
						doc.ksOpenTable(pObj);

						ksTextParam par = (ksTextParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextParam);
						if (par == null)
							return;
						par.Init();
						//� ����� ����� ����� ��� ������������ ������
						while (doc.ksGetTableColumnText(ref numb, par)!=0) 
						{
							buf = string.Format("numb = {0}", numb);
							kompas.ksMessage(buf);


							ksDynamicArray arrpLineText = (ksDynamicArray) par.GetTextLineArr() ;
							ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
							if (itemLineText == null)
								return;
							itemLineText.Init();

							for (int i = 0; i < arrpLineText.ksGetArrayCount(); i++) 
							{
								arrpLineText.ksGetArrayItem(i, itemLineText);
								buf = string.Format("i = {0} style = {1}", i, itemLineText.style);
								kompas.ksMessage(buf);

								ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
								ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;
								if (item == null || arrpTextItem == null)
									return;
								item.Init();
          
								for (int j = 0; j < arrpTextItem.ksGetArrayCount(); j ++) 
								{
									arrpTextItem.ksGetArrayItem(j, item);
									ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;

									if (item.type == 0)
										buf = string.Format("���������� = {0} h = {1:0.##}\ns = {2}\nfontName = {3}\n��������� = {4}",
											j, textItemFont.height, item.s, textItemFont.fontName, textItemFont.bitVector);
									else
										buf = string.Format("���������� = {0} ��� = {1} ����� ��������� = {2}",
											j, item.type, item.iSNumb);
									kompas.ksMessage(buf);
								}
								arrpTextItem.ksDeleteArray();  //������� ������ ���������
							}
							//������� ������ ��������� �����
							arrpLineText.ksDeleteArray();
						}

						//����� ������ 2
						doc.ksColumnNumber(2);
						doc.ksText(0, 0, 0, 5, 1, 0, "������ ������");

						doc.ksDivideTableItem(3, true, 2);
						doc.ksColumnNumber(4);
						doc.ksText(0, 0, 0, 5, 1, 0, "4");

						doc.ksEndObj();	//������� ������ "�������"
					}
					else
						kompas.ksError("��� �� �������");
				}
				else
					kompas.ksError("��� �������");
			}

			kompas.ksMessageBoxResult();
		}

		void EditStamp(ksDocument2D doc)
		{
			ksStamp stamp = (ksStamp) doc.GetStamp() ;
			if (stamp != null && stamp.ksOpenStamp() == 1) 
			{
				int numb = 0;
				//� ����� ����� ����� ��� ������������ �����
				ksDynamicArray arr = (ksDynamicArray) stamp.ksGetStampColumnText(ref numb);
				ksTextItemParam item = null;
				while (numb != 0 && arr != null) 
				{
					string buf = string.Empty;
					buf = string.Format("numb = {0}", numb);
					kompas.ksMessage(buf);

					ksDynamicArray arrpLineText = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
					ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
					if (itemLineText == null)
						return;
					itemLineText.Init();

					for(int i = 0, count = arr.ksGetArrayCount(); i < count; i++) 
					{
						arr.ksGetArrayItem(i, itemLineText);
						buf = string.Format("i = {0} style = {1}", i, itemLineText.style);
						kompas.ksMessage(buf);

						ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
						item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

						if (item == null || arrpTextItem == null)
							return;
						item.Init();
          
						for (int j=0, count2 = arrpTextItem.ksGetArrayCount(); j < count2; j++) 
						{
							arrpTextItem.ksGetArrayItem(j, item);
							ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;
							buf = string.Format("���������� = {0} h = {1:0.##}\ns = {2}\nfontName = {3}",
								j, textItemFont.height, item.s, textItemFont.fontName);
							kompas.ksMessage(buf);
						}
						arrpTextItem.ksDeleteArray();  //������� ������ ���������
					}
					//������� ������ ��������� �����
					arrpLineText.ksDeleteArray();

					arr.ksDeleteArray();
					arr = (ksDynamicArray) stamp.ksGetStampColumnText(ref numb);
				}
				//�������  ����� 2
				doc.ksColumnNumber(2);
				item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;
				if (item == null) 
				{
					stamp.ksCloseStamp();
					return;
				}
				item.Init();
				ksTextItemFont itemFont = (ksTextItemFont) item.GetItemFont() ;
				if (item != null && itemFont != null) 
				{
					itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
					item.s = "����� 2";
					doc.ksTextLine(item);
				}
				stamp.ksCloseStamp();
			}
			else
				kompas.ksError ("����� �� ������");

			kompas.ksMessageBoxResult();
		}

		void GetTextTT(ksDocument2D doc)
		{
			//������� ��������� �� ����������� ���������
			reference pTT = doc.ksGetReferenceDocumentPart(1);
			ksTechnicalDemandParam technicalDemandParam = (ksTechnicalDemandParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TechnicalDemandParam) ;
			if (pTT != 0 && technicalDemandParam != null) 
			{
				technicalDemandParam.Init();
				//������� ��������� �������� ��
				doc.ksGetObjParam(pTT, technicalDemandParam, ldefin2d.TECHNICAL_DEMAND_PAR);
				string buf = string.Empty;
				ksDynamicArray pGab = (ksDynamicArray) technicalDemandParam.GetPGab();
				int count = pGab.ksGetArrayCount();
				buf = string.Format("����� = {0} ����� �������  TT = {1}",
					technicalDemandParam.style, count);
				kompas.ksMessage(buf);

				// �������� ������ ��������� �����
				ksDynamicArray pTextLine = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
				//��������� �� ������ �� � ������� �����
				for(int i = 0; i < count; i++) 
				{
					doc.ksGetObjParam(pTT, pTextLine, i);//ALLPARAM);
					ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam) ;
					if (itemLineText == null)
						return;
					itemLineText.Init();

					//      TextLineParam par2;
					for(int i1 = 0, count1 = pTextLine.ksGetArrayCount(); i1 < count1; i1 ++) 
					{
						pTextLine.ksGetArrayItem(i1, itemLineText);
						buf = string.Format("���������� = {0} style = {1}",
							i1, itemLineText.style);
						kompas.ksMessage(buf);

						ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
						ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;

						if (item == null || arrpTextItem == null)
							return;
						item.Init();

						for (int j = 0, count2 = arrpTextItem.ksGetArrayCount(); j < count2; j ++) 
						{
							arrpTextItem.ksGetArrayItem(j, item);
							ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;
							buf = string.Format("���������� = {0} h = {1:0.##}\ns = {2}\nfontName = {3}",
								j, textItemFont.height, item.s, textItemFont.fontName);
							kompas.ksMessage(buf);
						}
						arrpTextItem.ksDeleteArray(); //������� ������ ���������
					}
				}
				pTextLine.ksDeleteArray(); //������� ������ ��������� �����
			}

			kompas.ksMessageBoxResult();
		}

		void ShowInsertFragment(ksDocument2D doc)
		{
			string libName = string.Empty;
			int res = 0;
			//�������  ���������� ����������
			libName = kompas.ksChoiceFile("*.lfr", "��������� ����������(*.lfr)|*.lfr|��� ����� (*.*)|*.*|", true) ;
			if (libName.Length > 0)
			{ 
				string buf = string.Empty; 
				do 
				{
					//������� �������� � ���������� ����������
					ksFragment fr = (ksFragment) doc.GetFragment() ;
					// ABB K6    ksFragmentLibrary
					ksFragmentLibrary frLib = (ksFragmentLibrary) kompas.GetFragmentLibrary() ; 
					if  (fr == null && frLib == null)
						return;
					buf = frLib.ksChoiceFragmentFromLib(libName, out res) ;
					if (buf != null && res == 3) // res = 3 - ������ ��������
					{
						// ������� ��� ������� 
						string insertName = buf.Substring(buf.IndexOf('|'));
						if (insertName != null && insertName != string.Empty) 
						{
							double x = 0, y = 0;
							//���������� ��������� ������ � �������� ��� Placement
							ksPhantom rub = (ksPhantom) kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom) ;
							rub.phantom = 1;
							ksType1 type1 = (ksType1) rub.GetPhantomParam() ;
							if (rub == null || type1 == null)
								return;

							type1.Init(); rub.Init(); 

							type1.scale_ = 1;
							rub.phantom = 1;

							reference pDefFrg;
							// �������� ��������  ������ ����������
							pDefFrg = fr.ksFragmentDefinition(buf,	//��� ����� ���������
								insertName + 1,						//��� �������
								1);									//��� ������� -������������ ��� �������� ���������
																	// 0- ����� � ��������, 1-������� �������

							if(pDefFrg > 0) 
							{
								//�� ��������� ������ ������� ������� ���������, ������ �� ���������� ���������� 
								type1.gr = doc.ksNewGroup(1);
   
								ksPlacementParam par = (ksPlacementParam) kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam) ;
								if (par == null)
									return; 
								par.Init();
								par.scale_ = 1;
								reference  p = fr.ksInsertFragment(pDefFrg, false, par);

								doc.ksEndGroup();
								int j;
								do 
								{
									type1.angle = 0;
									double ang = type1.angle;
									if ((j = doc.ksPlacement(null, ref x, ref y, ref ang, rub))!=0) 
									{
										type1.angle = ang;
										doc.ksCopyObj(p,	// ��������� �� ����������� ������
											0, 0,			// ������� ����� �������
											x, y,			// ����� ���� ����������
											1, type1.angle);// ������� � ���� �������� � ��������
									}
								}
								while (j > 0); 
								doc.ksDeleteObj(type1.gr);
							}
							else
								kompas.ksError ("������ �������� �������� ������� ���������");
						}
						else
							kompas.ksError("��� ������� �� ����������");
					}
				} while(res > 0);
			}

			kompas.ksMessageBoxResult();
		}

		void ShowInsertFragment1(ksDocument2D doc)
		{
			string frwName = string.Empty;
			ksFragment fr = (ksFragment) doc.GetFragment() ;
			if (fr == null)
				return;
			//�������  ��������
			frwName = kompas.ksChoiceFile("*.frw", "���������(*.frw)|*.frw|��� ����� (*.*)|*.*|", true);
			if(frwName != null && frwName != string.Empty) 
			{ 
				double x = 0, y = 0;
				//���������� ��������� ������ � �������� ��� Placement
				ksPhantom rub = (ksPhantom) kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom);
				rub.phantom = 1;
				ksType1 type1 = (ksType1) rub.GetPhantomParam();
				if (rub == null || type1 == null)
					return;
				type1.Init();
				rub.Init(); 

				type1.scale_ = 1;
				rub.phantom = 1;

				//�� ��������� ������ ������� ������� ���������, ������ �� ���������� ����������
				ksPlacementParam par = (ksPlacementParam) kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam);
				if (par == null)
					return; 
				par.Init();
				par.scale_ = 1;

				int j;
				do 
				{
					//���� ����� �������� ��������� ����������, ������ ����� ��������� �����,
					//��� ��� ������ � ���������� ����� ������ ��������, ������� ������������, �����,
					//������� ������� � ����������. ��� ������� ����������� ������ ��� ����� �����
					//��������.
					type1.gr = fr.ksReadFragmentToGroup(frwName, false, par);
					if (type1.gr > 0) 
					{
						double ang = type1.angle;
						if ((j = doc.ksPlacement(null, ref x, ref y, ref ang, rub)) != 0) 
						{
							//�������� ������
							doc.ksMoveObj(type1.gr, x, y);
							//������������ ������
							if(Math.Abs(ang) > 0.001)
								doc.ksRotateObj(type1.gr, x, y, ang);
							//������ ������ � ������
							doc.ksStoreTmpGroup(type1.gr);
							doc.ksClearGroup(type1.gr, true);
							doc.ksDeleteObj(type1.gr);
						}
					}
					else 
					{
						if (type1.gr > 0)
							doc.ksDeleteObj(type1.gr);
						j = 0;
					}
				} while (j > 0);
			}

			kompas.ksMessageBoxResult();
		}

		void EditFragmentLibrary(ksDocument2D doc)
		{
			string libName = string.Empty;
			string buf = string.Empty;
			//�������  ���������� ����������
			libName =  kompas.ksChoiceFile("*.lfr", "��������� ����������(*.lfr)|*.lfr|��� ����� (*.*)|*.*|", true) ;
			if (libName != null && libName != string.Empty) 
			{ 
				ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo) ;
				ksFragment fr = (ksFragment) doc.GetFragment() ;
				// ��� �6 ksFragmentLibrary
				ksFragmentLibrary frLib = (ksFragmentLibrary) kompas.GetFragmentLibrary(); 
				if  (info == null || fr == null || frLib == null)
					return;
				info.Init();

				info.commandsString = "!����� �������� !������������� �������� !������� �������� ";
				int j;
				int typeEdit = 0;
				string nameFrg = string.Empty;
				do 
				{
					j = doc.ksCommandWindow(info);
					switch (j) 
					{
						case 1:  //!�����_��������
							buf =  kompas.ksReadString("������� ��� ������ ���������", "") ;
							if (buf != null && buf != string.Empty) 
							{
								nameFrg = libName;
								if (buf[0] != '|')
									nameFrg += "|";
								nameFrg += buf;
								typeEdit = 2; //��������� �� ��������������
							}
							else
								typeEdit = 0;
							break;
						case 2 : //�������������_��������
						case 3 : //�������_��������
							//������� ��� ����� ���������
							int res = 0;
							buf = frLib.ksChoiceFragmentFromLib(libName, out res);
							if (res > 0 && buf != null && buf != string.Empty && (j == 2 || j == 3))
							{
								nameFrg = buf;
								typeEdit = j; // 2- ��������� �� ��������������, 3-������� ;
							}
							else
								typeEdit = 0;
							break;
					}

					if (j > 0 && typeEdit > 0) 
					{
						if (frLib.ksFragmentLibraryOperation(nameFrg, typeEdit) == 1) 
						{
							if (typeEdit == 2) 
							{
								frLib.ksFragmentLibraryOperation(nameFrg, 4 /*�������������� ���� ������������*/);
								//����������� �������� �� ����������
								doc.ksText(0, 100,	//����� �������� ������
									0,				//���� ������� ������
									5,				//������ ������
									1,				//������� ������
									0,				//�������� ������, ������� �������� ���.-����.
									"����������� �������� �� ����������");	//������ ��������

								doc.ksLineSeg (0, 100, 110, 100, 1);
								//����������� �������� � ������������� ������ 
								//����� ������ � ���� "������" ������� "��������� �������������� ���������", 
								//������������ � ����������
								if (kompas.ksSystemControlStart("��������� �������������� ���������") != 1)
									kompas.ksStrResult();
								if (!kompas.ksSystemControlStop())
									kompas.ksStrResult();
								frLib.ksFragmentLibraryOperation(nameFrg, 0 /*������� c �����������*/);
							}
						}  
						else
							kompas.ksMessageBoxResult();
					}

				} while (j != -1);
			}
		}

		void ChangeTechnicalDemand(ksDocument2D doc)
		{
			int rf = doc.ksGetReferenceDocumentPart(1);
			ksTechnicalDemandParam par = (ksTechnicalDemandParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TechnicalDemandParam);
			if (doc.ksGetObjParam(rf, par, ldefin2d.TECHNICAL_DEMAND_PAR) != 0) 
			{
				string buf = string.Empty;
				buf = string.Format("����� ����� TT = {0}", par.strCount);
				kompas.ksMessage(buf);
	
				doc.ksOpenTechnicalDemand(par.GetPGab(), par.style);

				ksTextLineParam parLine = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
				if (parLine != null)
				{
					parLine.Init();
					for(int i = 0; i < par.strCount; i++) 
					{
						doc.ksGetObjParam(rf, parLine, ldefin2d.TT_FIRST_STR + i);
						ksTextItemParam parItem = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
						ksDynamicArray arr = (ksDynamicArray) parLine.GetTextItemArr();
						if (parItem != null && arr != null) 
						{
							parItem.Init();
							for (int j = 0, count1 = arr.ksGetArrayCount(); j < count1; j++) 
							{
								arr.ksGetArrayItem(j, parItem);
								kompas.ksMessage(parItem.s);
								parItem.s = string.Format("{0}-� ������", i + 1);  
								arr.ksSetArrayItem(j, parItem);
							}
						}
						doc.ksSetObjParam(rf, parLine, ldefin2d.TT_FIRST_STR + i);
					}
				}
				doc.ksCloseTechnicalDemand();
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
