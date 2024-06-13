using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step10 - H�������� �� ������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step10
	{
		private KompasObject kompas;
		private ksDocument2D docActive;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step10 - H�������� �� ������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				if (command == 2 || command == 3)
				{
					docActive = (ksDocument2D) kompas.ActiveDocument2D();
					if (docActive != null && docActive.reference != 0)
					{
						switch (command)
						{
							case 2:	CreateDet(docActive);		break;	// ������� ������ ������������ ������
							case 3:	CreateStandart(docActive);	break;	// ������� ����������� ������
						}
					}
					else
						kompas.ksError("�������� �� ������������� ��� \n�� �������� ������/����������");
				}
				else
				{
					switch (command)
					{
						case 1:	TypeAttrBolt();	break; // ��������  ��� �������� �����
						case 4: DecomposeSpc();	break; // k������������� ������������ �� ��������
						case 5: ShowSpc();		break; // ����������� ������������
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
					result = "������� ������ ����������� ����";
					command = 1;
					break;
				case 2:
					result = "������� ������ ��� ������� ������";
					command = 2;
					break;
				case 3:
					result = "������� ������ ��� ������� ����������� �������";
					command = 3;
					break;
				case 4:
					result = "�������������� ������������ �� ��������";
					command = 4;
					break;
				case 5:
					result = "����������� ������� ������������";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// ��������  ��� �������� �����
		void TypeAttrBolt()
		{
			ksAttributeObject attr = (ksAttributeObject)kompas.GetAttributeObject();
			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (type != null && col != null && attr != null) 
			{
				type.Init();
				col.Init(); 
				type.header = "����111";		// �������o�-����������� ����
				type.rowsCount = 1;				// ���-�� ����� � �������
				type.flagVisible = true;		// �������, ���������   � �������
				type.password = string.Empty;	// ������, ���� �� ������ ������  - �������� �� �������������������� ��������� ����
				type.key1 = 10;
				type.key2 = 20;
				type.key3 = 30;
				type.key4 = 0;
				ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
				if (arr != null) 
				{
					// ������� 1 "��� ����."
					col.header = "��� ����.";				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 1;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "����";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 2 "����������"
					col.header = "����������";				// �������o�-����������� �������
					col.type = ldefin2d.UINT_ATTR_TYPE;		// ��� ������ � ������� - ��.����
					col.key = 3;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "1";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 3 "������"
					col.header = "������";					// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "M";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 4 "�������"
					col.header = "�������";					// �������o�-����������� �������
					col.type = ldefin2d.UINT_ATTR_TYPE;		// ��� ������ � ������� - ��.����
					col.key = 3;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "12";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 5 "�����������"
					col.header = string.Empty;				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "x";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 6 "���"
					col.header = "���";						// �������o�-����������� �������
					col.type = ldefin2d.FLOAT_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 3;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "1.25";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 7 "���� �������"
					col.header = "���� �������";			// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 3;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "-6g";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 8 "�����������"
					col.header = string.Empty;				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "x";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 9 "�����"
					col.header = "�����";					// �������o�-����������� �������
					col.type = ldefin2d.UINT_ATTR_TYPE;		// ��� ������ � ������� - ��.����
					col.key = 3;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "60";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 10 "��. ���������"
					col.header = "��. ���������";			// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "58";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 11 "��������"
					col.header = "��������";				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = ".35�";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 12 "��������"
					col.header = "��������";				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = ".16";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 13 "����"
					col.header = "����";					// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "����";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 14 "�����"
					col.header = "�����";					// �������o�-����������� �������
					col.type = ldefin2d.UINT_ATTR_TYPE;		// ��� ������ � ������� - ��.����
					col.key = 2;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "7808";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 15 "�����������"
					col.header = string.Empty;				// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "-";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������� 16 "���"
					col.header = "���";						// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "70";							// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);
				}

				//��������� ��� ����������
				string nameFile = kompas.ksChoiceFile("*.lat", null, false);
				if (nameFile.Length > 0) 
				{
					//������� ��� ��������
					double  numbType = attr.ksCreateAttrType(type,	// ���������� � ���� ��������
						nameFile);									// ��� ���������� ����� ���������
					if (numbType > 1)  
					{
						string buf = string.Format("numbType = {0}", numbType);
						kompas.ksMessage(buf);
					}
					else
						kompas.ksMessageBoxResult();	// ��������� ��������� ������ ����� �������
				}
				arr.ksDeleteArray();	//������ ������ �������
			}
		}

		// ������� ������ ������������ ������
		void CreateDet(ksDocument2D doc)
		{
			ksSpecification spc = (ksSpecification)doc.GetSpecification();
			if (spc != null) 
			{
				//������� ��������� ������   1
				reference gr = doc.ksNewGroup(0);
				doc.ksLineSeg (20, 30, 70, 30, 2);
				doc.ksLineSeg (70, 30, 70, 80, 2);
				doc.ksLineSeg (70, 80, 20, 80, 2);
				doc.ksLineSeg (20, 80, 20, 30, 2);
				doc.ksEndGroup();

				//�������  ������ ������������
				reference spcObj = EditSpcObjDet(gr, doc, spc);
				if (spcObj != 0)
					DrawPosLeader(spcObj, doc, spc);
			}
		}

		// ������� ����������� ������
		void CreateStandart(ksDocument2D doc)
		{
			ksSpecification spc = (ksSpecification)doc.GetSpecification();
			if (spc != null) 
			{
				BOLT tmp = new BOLT();
				tmp.gost = 7808;
				tmp.f = 0;
				tmp.p = 1; 
				tmp.L = 55;
				tmp.l1 = 49;
				tmp.b = 46;
				tmp.h2 = 10;
				tmp.klass = 1; /*klass = B*/
				tmp.d2 = 22.5F;
				tmp.k = 2;
				tmp.dr = 20;
				//�������  ������ ������������
				reference spcObj = EditStandartSpcObj(tmp, 0, doc, spc, 0);
				if (spcObj != 0)
					DrawPosLeader(spcObj, doc, spc);
			}
		}

		// ����������� ������������
		void ShowSpc() 
		{
			ksDocument2D doc = (ksDocument2D)kompas.Document2D();
			ksSpcDocument spc = (ksSpcDocument)kompas.SpcActiveDocument();
			if (doc != null && spc != null && spc.reference != 0) 
			{
				ksSpecification specification = (ksSpecification)spc.GetSpecification();
				ksIterator iter = (ksIterator)kompas.GetIterator();
				iter.ksCreateSpcIterator(null, 0, 0);
				if (iter.reference != 0 && specification != null) 
				{
					int obj = iter.ksMoveIterator("F");
					if (obj != 0) 
					{
						do 
						{
							//������ ���������� ������� � �������� ������� ������������
							int count = specification.ksGetSpcTableColumn(null, 0, 0);
  			  
							string buf = string.Format("���-�� ������� = {0}", count);
							kompas.ksMessage(buf);
      
							// ������� �� ���� ��������
							for (int i = 1; i <= count; i++) 
							{
								// ��� �������� ������ ��������� ��� �������, ����� ���������� � ����
								ksSpcColumnParam spcColPar = (ksSpcColumnParam)kompas.GetParamStruct((short)StructType2DEnum.ko_SpcColumnParam);
								if (specification.ksGetSpcColumnType(obj,	//������ ������������
									i,										// ����� �������, ������� � 1
									spcColPar) == 1)
								{
									// ������� �����
									int columnType = spcColPar.columnType;
									int ispoln = spcColPar.ispoln;
									int blok = spcColPar.block;
									buf = specification.ksGetSpcObjectColumnText(obj, columnType, ispoln, blok);
									kompas.ksMessage(buf);
									// �� ���� �������, ������ ���������� � ����� ��������� ����� �������
									int colNumb = specification.ksGetSpcColumnNumb(obj,	//������ ������������
										spcColPar.columnType, spcColPar.ispoln, spcColPar.block);
									buf = string.Format("i = {0} colNumb = {1}", i, colNumb);
									kompas.ksMessage(buf);
								}
							}
						}
						while ((obj = iter.ksMoveIterator("N")) != 0);
					}
				}	
			}
			else 
				kompas.ksError("������������ ������ ���� �������");
		}

		// ����������� ������������
		void DecomposeSpc() 
		{
			ksSpcDocument spc = (ksSpcDocument)kompas.SpcActiveDocument();
			ksDocument2D doc = (ksDocument2D)kompas.Document2D();
			if (spc != null && doc != null) 
			{
				// ����� ������� �������� ������������
				reference pDoc = spc.reference;
				if (pDoc == 0)
					return;
				// ������ ���������� ������ ������������
				int pageCount = spc.ksGetSpcDocumentPagesCount();
				// ������ �������� ������ ����� ������������
				ksRectParam spcGabarit = (ksRectParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectParam);
				if (spcGabarit == null)
					return;
				doc.ksGetObjGabaritRect(pDoc, spcGabarit);

				// �������� ��������  
				ksDocumentParam docPar = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
				if (docPar == null)
					return;
				docPar.Init();
				docPar.regime = 0;
				docPar.type = 3;
				doc.ksCreateDocument (docPar);
				for (int i = 0; i < pageCount; i ++) 
				{
					// �������� ��������� ������ i-�� ����� ������������
					reference group = doc.ksDecomposeObj(pDoc,	//��������� �� ������
						0,										//������� ��������� 0- �������,����,������,�����:1-�������,������,�����; 2-�������, ����, ������
						0.4,									//������� �������
						Convert.ToInt16(i + 1));				// 0 - ��������� ������� � �� ���� 1- � �� �����
					if (group != 0) 
					{
						int column = i % 3;
						ksMathPointParam mathBop = (ksMathPointParam)spcGabarit.GetpBot();
						ksMathPointParam mathTop = (ksMathPointParam)spcGabarit.GetpTop();
						if (mathBop == null && mathTop == null)
							return;
						double x = (mathTop.x - mathBop.x + 5) * column;
						int row = i / 3;
						double y = (mathTop.y - mathBop.y + 5) * row;
						// �������� ������
						doc.ksMoveObj(group, x, -y);
						// ��������� � ������
						doc.ksStoreTmpGroup(group);	// ��������� ������
						doc.ksClearGroup(group, true);
						doc.ksDeleteObj(group);
					}
				}
				string buf = string.Format("��������������� {0} ������ ��", pageCount);
				kompas.ksMessage(buf);
			}
			else
				kompas.ksError("������������ ������ ���� �������");
		}

		// �������� ������ ������������ ��� ������� "������"
		reference  EditSpcObjDet(reference geom, ksDocument2D doc, ksSpecification spc)
		{
			reference spcObj = 0;
			// ���� ����������� ����� ������, �� ������� � ����� �������������� �������
			if (doc.ksEditMacroMode() == 1) 
			{
				// ������ ����� ������������ �� ���������
				spcObj = spc.ksGetSpcObjForGeom("graphic.lyt",	// ��� ���������� ����� 
					1,											// ����� ���� ������������
					0,											// ��� ����� ��������������
					1, 1);
				// ������� � ����� ��������������                               
				if (spc.ksSpcObjectEdit(spcObj) != 1)
					spcObj  = 0;
			}

			// ���� ������� ���, ������� � ����� �������� ������� ������������
			if (spcObj != 0 || spc.ksSpcObjectCreate("graphic.lyt",	// ��� ���������� ����� 
				1,													// ����� ���� ������������
				20, 0,												// ����� ������� � ����������
				0, 0) == 1) 
			{
				// ��� ��������
				ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (par == null || item == null || arr == null)
					return 0;
				par.Init();
				par.SetUserArray(arr);
				item.Init();
				item.strVal = "������";
				arr.ksAddArrayItem(-1, item);
				// ������������
				spc.ksSpcChangeValue(5 /*����� �������*/, 1, par, ldefin2d.STRING_ATTR_TYPE );

				// ��������� ���������
				if (geom != 0)
					spc.ksSpcIncludeReference(geom, 1);

				spcObj = spc.ksSpcObjectEnd();
				// ���� ������ ������������ ������ ����� ����������� ������������ �� ���� ���������� � 
				// ���������������  ������� ����� ��������� ��� Cursor � Placement
				if (spcObj != 0 && spc.ksEditWindowSpcObject(spcObj) == 1)
					return spcObj;
			}
			return 0;
		}

		// ���������� ����������� ����� �������
		// ��� ���������, ��� ������ ������������ ����
		// ������� ����� ��������� ��� Cursor � Placement
		void DrawPosLeader(reference spcObj, ksDocument2D doc, ksSpecification spc)
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			info.Init();
			bool flag = false;
			reference posLeater = 0;
			double x1 = 0, y1 = 0;
			do 
			{
				info.commandsString = "!������� !���������� ";
				info.prompt = "������� ����������� ����� �������";
				int j1 = doc.ksCursor(info, ref x1, ref y1, null);
				switch (j1) 
				{
					case 1: // ������� ����� ����� �������
						posLeater = doc.ksCreateViewObject(ldefin2d.POSLEADER_OBJ);
						flag = false;
						break;
					case 2: // ���������� ������������
						info.prompt = "������� ����� �������";
						if (doc.ksCursor(info, ref x1, ref y1, 0) == -1) 
						{
							posLeater = doc.ksFindObj(x1, y1, 100);      // �������� ������� ������-������� � ������� x,y
							if (!(posLeater != 0 && doc.ksGetObjParam(posLeater, 0, 0) == ldefin2d.POSLEADER_OBJ)) 
							{
								posLeater = 0;
								flag = true;
							}
							else
								flag = false;
							break;
						}
						else
							flag = false;
						break;
					case -1:
						posLeater = doc.ksFindObj(x1, y1, 100);      // �������� ������� ������-������� � ������� x,y
						if (!(posLeater != 0 && doc.ksGetObjParam(posLeater, 0, 0) == ldefin2d.POSLEADER_OBJ)) 
						{
							kompas.ksError("������! ������ - �� ����������� ����� �������!");
							posLeater = 0;
							flag = true;
						}
						else
							flag = false;
						break;
				}
			}
			while(flag);
  
			// ����� ������� ����, ��������� �� � ������� ������������ 
			if (posLeater != 0) 
			{
				// ������� � ����� �������������� ������� ������������
				if (spc.ksSpcObjectEdit(spcObj) == 1) 
				{
					// ��������� ����� �������
					spc.ksSpcIncludeReference(posLeater, 1);
					// ������� ������ ������������
					spc.ksSpcObjectEnd();
				}
			}
		}

		// �������� ������ ������������ ��� ������� "����������� �������"
		reference EditStandartSpcObj (BOLT tmp,  reference geom, ksDocument2D doc, ksSpecification spc, reference spcObj)
		{
			// reference spcObj = 0;
			// ���� ����������� ����� ������ , �� ������� � ����� �������������� �������
			if (doc.ksEditMacroMode() == 1) 
			{
				// ������ ����� ������������ �� ���������
				spcObj = spc.ksGetSpcObjForGeom("graphic.lyt",	// ��� ���������� �����
					1,											// ����� ���� ������������
					0,											// ��� ����� ��������������
					1, 1);
				// ������� � ����� ��������������
				if (spc.ksSpcObjectEdit(spcObj) != 1)
					spcObj  = 0;
			}

			if (spcObj!= 0 || spc.ksSpcObjectCreate("graphic.lyt",	// ��� ���������� �����
				1,													// ����� ���� ������������
				25, 0,												// ����� ������� � ���������� ���������� � ����� ������������
				313277777065.0, 0) == 1) 
			{
				// ��� �������� ��������� � ���������� ����� ��������� spc.lat
				int uBuf;
				ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (par == null || item == null || arr == null)
					return 0;
				par.Init();
				par.SetUserArray(arr);
				item.Init();
				item.strVal = "����111";
				arr.ksAddArrayItem(-1, item);

				spc.ksSpcChangeValue(5 /*����� �������*/, 1, par, ldefin2d.STRING_ATTR_TYPE);
				// ����������
				if ((tmp.f & 0x80) == 0) // ���� ���������� ���
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 2, 0); // �������� ����������
				else 
				{
					uBuf = 2;
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 2, 1);
					arr.ksClearArray();
					item.Init();
					item.intVal = uBuf;
					arr.ksAddArrayItem(-1, item);
					spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 2, par, ldefin2d.UINT_ATTR_TYPE );
				}

				// ������� �������
				arr.ksClearArray();
				item.Init();
				item.doubleVal = 40;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 4, par, ldefin2d.DOUBLE_ATTR_TYPE );

				// �������� ������ ���
				if ((tmp.f & 0x2) == 0)
				{
					// ��������� ��� � ��� �����������
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 5, 0);
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 6, 0);	// ���
				}
				else 
				{
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 5, 1);
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 6, 1);   //���
					arr.ksClearArray();
					item.Init();
					item.floatVal = tmp.p;
					arr.ksAddArrayItem(-1, item);
					spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 6, par, ldefin2d.FLOAT_ATTR_TYPE);
				}

				// �������� ���� �������
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 7, 0);

				// ������� �����
				arr.ksClearArray();
				item.Init();
				item.intVal = (int)tmp.L;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 9, par, ldefin2d.UINT_ATTR_TYPE);

				// �������� ����� ���������
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 10, 0);

				// �������� ��������
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 11, 0);

				// �������� ��������
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 12, 0);

				// ������� ����
				arr.ksClearArray();
				item.Init();
				item.floatVal = tmp.gost;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 14, par, ldefin2d.UINT_ATTR_TYPE);

				// ��������� ���������
				if (geom != 0)
					spc.ksSpcIncludeReference(geom, 1);

				spcObj = spc.ksSpcObjectEnd();
				// ���� ������ ������������ ������ ����� ����������� ������������ �� ���� ���������� � 
				// ���������������. ������� ����� ��������� ��� Cursor � Placement
				if (spcObj != 0)
					if (spc.ksEditWindowSpcObject(spcObj) != 0)
						return spcObj;
			}
			return 0;
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
