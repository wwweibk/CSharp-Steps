using Kompas6API5;

using System;
using System.Drawing;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step7 - ����� ������� ���������� �� C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step7
	{
		KompasObject kompas;
		ksDocument2D doc;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step7 - H�������� �� ������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas == null)
				return;

			if (command == 3 || command == 4)
				doc = kompas.Document2D() as ksDocument2D;
			else
				doc = kompas.ActiveDocument2D() as ksDocument2D;

			if (doc == null)
				return;
			
			switch (command)
			{
				case 1: WalkFromView();			break;	//�������� �� ����
				case 2: WalkFromMacro();		break;	//�������� �� �������������
				case 3: WalkFromDoc();			break;	//�������� �� ����������
				case 4: WalkViewDoc();			break;	//�������� �� �����
				case 5: WalkGroup();			break;	//�������� �� ������������ � ������� �������
				case 6: WalkLayer();			break;	//�������� �� �����
				case 7: WalkFromGroup();		break;	//�������� �� ������
				case 8: WalkFromDocWithAttr();	break;	//�������� �� ��������� ��������� � ������������ ���������
				case 9: WalkFromObjWithAttr();	break;	//�������� �� ��������� �������
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
					result = "X������� �� ����";
					command = 1;
					break;
				case 2:
					result = "X������� �� �������������";
					command = 2;
					break;
				case 3:
					result = "X������� �� ����������";
					command = 3;
					break;
				case 4:
					result = "X������� �� �����";
					command = 4;
					break;
				case 5:
					result = "X������� �� �������";
					command = 5;
					break;
				case 6:
					result = "X������� �� �����";
					command = 6;
					break;
				case 7:
					result = "X������� �� ������";
					command = 7;
					break;
				case 8:
					result = "X������� �� ��������� � ���������";
					command = 8;
					break;
				case 9:
					result = "X������� �� ��������� �������";
					command = 9;
					break;
				case 10:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		void UserLightObj(reference obj, string c, int count, ksDocument2D doc)
		{
			string buf = string.Empty;
			doc.ksLightObj(obj, 1);
			buf = string.Format(c, count);
			kompas.ksMessage(buf);
			doc.ksLightObj(obj, 0);
		}

		void WalkFromView()
		{
			//� ������� ��������� � ���� �������� �������� ��� �������� �� ���� ���������
			reference obj;
			int count = 0;
			string buf = string.Empty;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, 0))
			{
				if (doc.ksExistObj(obj = iter.ksMoveIterator("F")) == 1)
				{
					do
					{
						doc.ksLightObj(obj, 1);
						count ++;
						buf = string.Format("����� = {0}", count);
						kompas.ksMessage(buf);
						doc.ksLightObj(obj, 0);
					}
					while(doc.ksExistObj(obj = iter.ksMoveIterator("N")) == 1);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkFromMacro() 
		{
			//� ������� ��������� � ���� �������� �������� ��� �������� �� ��������������
			reference obj, macro;
			string s = "����� ����� = {0}";
			string s1 = "����� ������� = {0}";
			int count = 0, count1 = 0;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.MACRO_OBJ, 0))
			{
				if (doc.ksExistObj(macro = iter.ksMoveIterator("F")) == 1)
				{
					do
					{
						count ++;
						UserLightObj(macro, s, count, doc);	// ��������� ������������
						ksIterator iter2 = (ksIterator)kompas.GetIterator();	// ������� �������� ��� �������� �� �������������
						iter2.ksCreateIterator(ldefin2d.ALL_OBJ, macro);
						if (iter2.reference != 0)
						{
							if (doc.ksExistObj(obj = iter2.ksMoveIterator("F")) == 1)
							{
								do
								{
									count1 ++;
									UserLightObj(obj, s1, count1, doc); // ��������� ������ � �����
								}
								while (doc.ksExistObj(obj = iter2.ksMoveIterator("N")) == 1);
							}
						}
					}
					while (doc.ksExistObj(macro = iter.ksMoveIterator("N")) == 1);
				}
				iter.ksDeleteIterator();
			}
			if (count == 0)
				kompas.ksError("������������� � ��������� �����������");
		}

		void WalkFromDoc() 
		{
			ksDocumentParam docPar = (ksDocumentParam)kompas.GetParamStruct((int)StructType2DEnum.ko_DocumentParam);
			string buf = string.Empty;
			if (docPar != null && doc != null)
			{
				docPar.Init();
				docPar.type = (short)DocType.lt_DocSheetStandart;
		
				ksSheetPar sheetPar = (ksSheetPar)docPar.GetLayoutParam();
				if (sheetPar == null) 
					return;
				sheetPar.Init();
				sheetPar.shtType = 1;
				sheetPar.layoutName = string.Empty;

				ksStandartSheet stSheet = (ksStandartSheet)sheetPar.GetSheetParam();
				if (stSheet == null) 
					return;
				stSheet.Init();

				docPar.comment = "Create document";
				docPar.author = "Ethereal";
				docPar.regime = 0;
				sheetPar.layoutName = string.Empty;
				sheetPar.shtType = 1;
				stSheet.format = 3;
				stSheet.multiply = 1;
				stSheet.direct = false;

				docPar.fileName = "a.cdw";
				doc.ksCreateDocument(docPar);
				docPar.fileName = "b.cdw";
				doc.ksCreateDocument(docPar);
				docPar.fileName = "c.cdw";
				doc.ksCreateDocument(docPar);

				int count = 0;
				ksIterator iter = (ksIterator)kompas.GetIterator();
				if (iter == null)
					return;
				if (iter.ksCreateIterator(ldefin2d.DOCUMENT_OBJ, 0))
				{
					//������� �������� ��� �������� �� ����������
					reference pDoc = iter.ksMoveIterator("F");
					if (pDoc != 0)
					{
						do
						{
							doc.reference = pDoc;
							if (doc.ksSetObjParam(pDoc, 0, ldefin2d.DOCUMENT_STATE) == 1)
							{
								//������������ �������� pDoc
								count ++;
								ksViewParam viewPar = (ksViewParam)kompas.GetParamStruct((int)StructType2DEnum.ko_ViewParam);
								if (viewPar == null)
									return;

								viewPar.Init();

								int number = count;
								viewPar.x = 10;
								viewPar.y = 20;
								viewPar.scale_ = 1;
								viewPar.angle = 0;
								viewPar.color = Color.FromArgb(0, 10, 20, 10).ToArgb();
								viewPar.state = ldefin2d.stACTIVE;
								viewPar.name = "User view";

								doc.ksCreateSheetView(viewPar, ref number); // �������� ��� � ���������
								doc.ksLayer(count); // ������� ����
								switch (count)
								{
									case 1: doc.ksLineSeg(20, 10, 40, 10, 1);				break; //� ������ ��������� �������� �������
									case 2: doc.ksCircle (50, 50, 20, 1);					break; //�� ������ ��������� �������� ����������
									case 3: doc.ksArcByAngle(50, 50, 20, 45, 135, 1, 1);	break; //� ������� ��������� ��������  ����
								}
								buf = string.Format("�������� {0}", count);
								kompas.ksMessage(buf);
							}
							pDoc = iter.ksMoveIterator("N");
						}
						while (pDoc != 0);
					}
					iter.ksDeleteIterator();
				}
			}
		}

		void CreateDocument(ksDocument2D doc)
		{
			ksDocumentParam docPar = (ksDocumentParam) kompas.GetParamStruct((int)StructType2DEnum.ko_DocumentParam);
			if (docPar == null || doc == null) 
				return;
			docPar.Init();
			docPar.type = (short)DocType.lt_DocSheetStandart;
	
			ksSheetPar sheetPar = (ksSheetPar)docPar.GetLayoutParam();
			if (sheetPar == null) 
				return;
			sheetPar.Init();
			sheetPar.shtType = 1;
			sheetPar.layoutName = string.Empty;

			ksStandartSheet stSheet = (ksStandartSheet)sheetPar.GetSheetParam();
			if (stSheet == null) 
				return;
			stSheet.Init();

			docPar.comment = "Create document";
			docPar.author = "Ethereal";
			docPar.regime = 0;
			stSheet.format = 3;
			stSheet.multiply = 1;
			stSheet.direct = false;

			docPar.fileName = "a.cdw";
			doc.ksCreateDocument(docPar);
		}

		void WalkViewDoc()
		{
			CreateDocument(doc); // ������� ��������

			//�������� 5 �����
			for(int i = 0; i < 5; i ++)
			{
				ksViewParam viewPar = (ksViewParam)kompas.GetParamStruct((int)StructType2DEnum.ko_ViewParam);
				if (viewPar == null)
					return;

				viewPar.Init();

				int number = 0;
				viewPar.x = 10;
				viewPar.y = 20;
				viewPar.scale_ = 1;
				viewPar.angle = 0;
				viewPar.color = Color.FromArgb(10, 20, 10).ToArgb();
				viewPar.state = ldefin2d.stACTIVE;
				viewPar.name = "User view";

				doc.ksCreateSheetView(viewPar, ref number); //�������� ��� � ���������
			}

			reference pView;
			int count = 0;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.VIEW_OBJ, 0))
			{
				//�������� �������� ��� ��������� �� ����� � ���������
				pView = iter.ksMoveIterator("F");
				if (pView != 0)
				{
					do
					{
						ksLtVariant var = (ksLtVariant)kompas.GetParamStruct((int)StructType2DEnum.ko_LtVariant);
						if (var == null) 
							return;
   			  
						var.Init();

						var.intVal = ldefin2d.stCURRENT;
						if (doc.ksSetObjParam(pView, var, ldefin2d.VIEW_LAYER_STATE) == 1)
						{
							switch (count)
							{
								case 1: doc.ksLineSeg(20, 20, 40, 20, 1);				break;
								case 2: doc.ksCircle(40, 20, 30, 1);					break;
								case 3: doc.ksArcByAngle(50, 50, 20, 45, 135, 1, 1);	break;
								case 4:
									doc.ksMtr(40, 0, 0, 1, 1);
									doc.ksLineSeg(10, 10, 30, 30, 1);
									doc.ksLineSeg(30, 30, 60, 10, 1);
									doc.ksLineSeg(60, 10, 10, 10, 1);
									doc.ksDeleteMtr();
									break;
								case 5:
									doc.ksCircle(30, 30, 20, 1);
									doc.ksCircle(30, 30, 10, 1);
									doc.ksHatch(0, 45, 2, 0, 0, 0);
									doc.ksCircle(30, 30, 20, 1);
									doc.ksCircle(30, 30, 10, 1);
									doc.ksEndObj();
									break;
							}
						}
						count ++;
						pView = iter.ksMoveIterator ("N");
					}
					while (pView != 0);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkGroup() 
		{
			reference pNameGrp;
			int count = 0;
			string buf = string.Empty;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.NAME_GROUP_OBJ, 0))
			{
				// �������� �������� ��� �������� �� ������������ ������� � ���������
				pNameGrp = iter.ksMoveIterator("F");
				if (pNameGrp != 0)
				{
					do
					{
						doc.ksLightObj(pNameGrp, 1);	// ��������� ������
						count ++;
						buf = string.Format("����� = {0}", count);
						kompas.ksMessage(buf);
						doc.ksLightObj(pNameGrp, 0);	// ������ ���������
					}
					while((pNameGrp = iter.ksMoveIterator("N")) != 0);
				}
				iter.ksDeleteIterator();
			}
			//��� ������������ ������ ������� � ������ ������� �����
			doc.ksNewGroup(0);
			doc.ksCircle(30, 30, 20, 1);
			doc.ksCircle(30, 30, 10, 1);
			doc.ksHatch(0, 45, 2, 0, 0, 0);
			doc.ksCircle(30, 30, 20, 1);
			doc.ksCircle(30, 30, 10, 1);
			doc.ksEndObj();
			doc.ksEndGroup();

			//������� �������� �� ������� �������
			count = 0;
			reference pWorkGrp;
			if (iter.ksCreateIterator(ldefin2d.WORK_GROUP_OBJ, 0))
			{
				//�������� �������� ��� �������� �� ������������ ������� � ���������
				pWorkGrp = iter.ksMoveIterator("F");
				if (pWorkGrp != 0)
				{
					do
					{
						doc.ksLightObj(pWorkGrp, 1);   //��������� ������
						count ++;
						buf = string.Format("����� = {0}", count);
						kompas.ksMessage(buf);
						doc.ksLightObj(pWorkGrp, 0);   //������ ���������
					}
					while ((pWorkGrp = iter.ksMoveIterator("N")) != 0);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkLayer()
		{
			string buf = string.Empty;
			CreateDocument(doc);	// ������� ��������

			//�������� 5 �����
			for(int i = 0; i < 5; i ++)
			{
				doc.ksLayer(i);
				doc.ksCircle(30, 30, 5 + i * 10, 1);
			}

			reference pLayer;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.LAYER_OBJ, 0))
			{
				//������� �������� �� �����
				pLayer = iter.ksMoveIterator("F");
				int count = 0;
				if (pLayer != 0)
				{
					do
					{
						doc.ksLightObj(pLayer, 1);   //��������� ������
						buf = string.Format("C��� = {0}", count);
						kompas.ksMessage(buf);
						doc.ksLightObj(pLayer, 0);   //������ ���������
						count ++;
					}
					while((pLayer = iter.ksMoveIterator ("N")) != 0);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkFromGroup() 
		{
			string buf = string.Empty;
			doc.ksMtr(20, 10, 0, 1, 1);
			reference pGrp = doc.ksNewGroup(0);
			doc.ksLineSeg(10, 50, 50, 50, 1);
			doc.ksLineSeg(10, 10, 50, 10, 1);
			doc.ksLineSeg(10, 10, 10, 50, 1);
			doc.ksLineSeg(50, 10, 50, 50, 1);
			doc.ksCircle(30, 30, 20, 1);
			doc.ksCircle(30, 30, 10, 1);
			doc.ksHatch(0, 45, 2, 0, 0, 0);
			doc.ksCircle(30, 30, 20, 1);
			doc.ksCircle(30, 30, 10, 1);
			doc.ksEndObj();
			doc.ksEndGroup();
			doc.ksDeleteMtr();
  
			reference obj;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, pGrp))
			{
				//������� �������� ��� �������� �� ������
				obj = iter.ksMoveIterator("F");
				int count = 0;
				if (doc.ksExistObj(obj) == 1)
				{
					do
					{
						doc.ksLightObj(obj, 1);
						count ++;
						buf = string.Format("����� = {0}", count);
						kompas.ksMessage(buf);
						doc.ksLightObj(obj, 0);
					}
					while (doc.ksExistObj (obj = iter.ksMoveIterator("N")) == 1);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkFromDocWithAttr() 
		{
			ksAttributeObject attr = (ksAttributeObject)kompas.GetAttributeObject();
			if (attr == null)
				return;

			reference pObj = 0, pAttr = 0;
			string buf = string.Empty;
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter == null)
				return;
			if (iter.ksCreateAttrIterator(0, 10, 0, 0, 0, 0))
			{
				// �������� �������� ��� ������ �������� �� ����� 10
				pAttr = iter.ksMoveAttrIterator("F", ref pObj);
				int count = 0;
				int rowsCount = 0, columnsCount = 0;
				if (pAttr != 0 && doc.ksExistObj(pObj) != 0)
				{
					do
					{
						doc.ksLightObj(pObj, 1);
						count ++;
						// ������ ���������� ����� � �������
						if (attr.ksGetAttrTabInfo(pAttr, out rowsCount, out columnsCount) == 1)
						{
							buf = string.Format("����� = {0} rowsCount = {1} columnsCount = {2}",
								count, rowsCount, columnsCount);
							kompas.ksMessage(buf);
						}
						else
							kompas.ksMessageBoxResult();	// ��������� ���������� - ������� ��������� ������ ����� �������

						doc.ksLightObj(pObj, 0);

						pAttr = iter.ksMoveAttrIterator("N", ref pObj);
					}
					while(pAttr != 0 && doc.ksExistObj(pObj) != 0);
				}
				iter.ksDeleteIterator();
			}
		}

		void WalkFromObjWithAttr() 
		{
			double x = 0, y = 0;
			int j;
			reference pObj;
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((int)StructType2DEnum.ko_RequestInfo);
			ksAttributeObject attr = (ksAttributeObject)kompas.GetAttributeObject();
			if (info == null || attr == null) 
				return;

			info.Init();

			info.prompt = "������� ������";
			do
			{
				j = doc.ksCursor(info, ref x, ref y, 0);
				if (j != 0)
				{
					if (doc.ksExistObj(pObj = doc.ksFindObj(x, y, 1000000)) == 1)
					{
						int count = 0;
						string buf = string.Empty;
						int rowsCount = 0, columnsCount = 0;
						doc.ksLightObj(pObj, 1);

						reference pAttr;
						ksIterator iter = (ksIterator)kompas.GetIterator();
						if (iter == null)
							return;
						if (iter.ksCreateAttrIterator(pObj, 10, 0, 0, 0, 0))
						{
							// �������� �������� ��� ������ �������� �� ����� 10
							pAttr = iter.ksMoveAttrIterator("F", ref j);
							if (pAttr != 0)
							{
								do
								{
									count ++;
									// ������ ���������� ����� � �������
									if (attr.ksGetAttrTabInfo (pAttr, out rowsCount, out columnsCount) == 1)
									{
										buf = string.Format("����� = {0} rowsCount = {1} columnsCount = {2}",
											count, rowsCount, columnsCount);
										kompas.ksMessage(buf);
									}
									else
										kompas.ksMessageBoxResult();	// ��������� ���������� - ������� ��������� ������ ����� �������
									pAttr = iter.ksMoveAttrIterator("N", ref j);
								}
								while(pAttr != 0);
							}
							iter.ksDeleteIterator();
						}
						doc.ksLightObj(pObj, 0);
					}
				}
			}
			while(j != 0);
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
