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
	// ����� Step8 - ������ � ����������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step8
	{
		private KompasObject kompas;
		private ksDocument2D doc;
		private ksAttributeObject attr;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step8 - ������ � ����������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			
			if (kompas  != null)
			{
				doc = (ksDocument2D)kompas.ActiveDocument2D();
				attr = (ksAttributeObject)kompas.GetAttributeObject();
				if (doc  != null && doc.reference  != 0 && attr  != null)
				{
					switch (command)
					{
						case 1  : FuncAttrType();			break; // ������� ��� ��������
						case 2  : DelTypeAttr();			break; // �������  ��� ��������
						case 3  : ShowTypeAttr();			break; // ��������  ��� ��������
						case 4  : ChangeType();				break; // ��������  ��� ��������
						case 5  : NewAttr();				break; // ������� ������� ������������� ����
						case 6  : DelObjAttr();				break; // ������� �������
						case 7  : ReadObjAttr();			break; // ������� �������
						case 8  : ShowObjAttr();			break; // ����������� �������
						case 9  : ShowLib();				break; // ����������� ����������
						case 10 : ShowType();				break; // ����������� ���
						case 11 : WalkFromObjWithAttr();	break; // ����������� �������
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
					result = "������� ��� ��������";
					command = 1;
					break;
				case 2:
					result = "������� ��� ��������";
					command = 2;
					break;
				case 3:
					result = "�������� ��� ��������";
					command = 3;
					break;
				case 4:
					result = "�������� ��� ��������";
					command = 4;
					break;
				case 5:
					result = "������� ������� ������������� ����";
					command = 5;
					break;
				case 6:
					result = "������� �������";
					command = 6;
					break;
				case 7:
					result = "������� �������";
					command = 7;
					break;
				case 8:
					result = "����������� �������� �������";
					command = 8;
					break;
				case 9:
					result = "����������� ����������";
					command = 9;
					break;
				case 10:
					result = "����������� ���";
					command = 10;
					break;
				case 11:
					result = "����������� �������";
					command = 11;
					break;
				case 12:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		void ShowCol(ksColumnInfoParam par, int iCol, bool fl)
		{
			string buf = string.Empty;
			string s = string.Empty;;
			if (fl)
				s = "���������";

			//������� ���� ������� �� ���������
			buf = string.Format("{0} i = {1} header = {2} type = {3} def = {4} flagEnum = {5}",
				s, iCol, par.header, par.type, par.def, par.flagEnum);
			kompas.ksMessage(buf);
			if (par.type == ldefin2d.RECORD_ATTR_TYPE)
			{
				// ���������
				ksDynamicArray pCol = (ksDynamicArray)par.GetColumns();
				if (pCol  != null)
				{
					ShowColumns(pCol, true);
					pCol.ksDeleteArray();
				}
			}
		}

		void ShowColumns(ksDynamicArray pCol, bool fl)
		{
			ksColumnInfoParam par = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (par != null) 
			{
				par.Init();
				int n = pCol.ksGetArrayCount();

				for (int i = 0; i < n; i++) 
				{
					if (pCol.ksGetArrayItem(i, par) != 1)
						kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������
					else
						ShowCol(par,i, fl);
				}
			}
		}

		// �������� ���� ���������
		void FuncAttrType() 
		{
			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (type != null && col != null)
			{
				type.Init();
				col.Init();
				type.header = "double_str_long";	// �������o�-����������� ����
				type.rowsCount = 1;					// ���-�� ����� � �������
				type.flagVisible = true;            // �������, ���������   � �������
				type.password = string.Empty;       // ������, ���� �� ������ ������  - �������� �� �������������������� ��������� ����
				type.key1 = 10;
				type.key2 = 20;
				type.key3 = 30;
				type.key4 = 0;
				ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
				if (arr != null)
				{
					// ������ �������  ���������
					col.header = "double";					// �������o�-����������� �������
					col.type = ldefin2d.DOUBLE_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "123456789";					// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������ ������� ���������
					col.header = "str";						// �������o�-����������� �������
					col.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "string";						// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);

					// ������ ������� ���������
					col.header = "long";					// �������o�-����������� �������
					col.type = ldefin2d.LINT_ATTR_TYPE;		// ��� ������ � ������� - ��.����
					col.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
					col.def = "1000000";					// �������� �� ���������
					col.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� ��������
					arr.ksAddArrayItem(-1, col);
				}
				string nameFile = string.Empty;
				//��������� ��� ����������
				nameFile = kompas.ksChoiceFile("*.lat", null, false);
				//������� ��� ��������
				double numbType = attr.ksCreateAttrType(type,	// ���������� � ���� ��������
					nameFile);									// ��� ���������� ����� ���������
				if (numbType > 1)
				{
					string buf = string.Format("numbType = {0}", numbType);
					kompas.ksMessage(buf);
				}
				else
					kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������

				//������  ������ �������
				arr.ksDeleteArray();
			}
		}

		void DelTypeAttr()
		{
			double numb = 0;
			int j;
			string password = string.Empty;
			//��������� ��� ����������
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);
			do
			{
				j = kompas.ksReadDouble("������ ����� ���� ��������", 1000, 0, 1e12, ref numb);
				if (j != 0)
				{
					password = kompas.ksReadString("������ ������ ���� ��������", "");
					if (attr.ksDeleteAttrType(numb, nameFile, password) != 1)
						kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������
				}
			}
			while(j != 0);
		}

		// �������� ��� ��������
		void ShowTypeAttr()
		{
			double numb = 0;
			//��������� ��� ����������
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);

			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			if (type != null)
			{
				type.Init();
				string buf = string.Empty;
				//		attrType.columns = CreateArray(ATTR_COLUMN_ARR, 0);

				do
				{
					numb = attr.ksChoiceAttrTypes(nameFile);
					if (numb != 0)
					{
						if (attr.ksGetAttrType(numb, nameFile, type) != 1)
							kompas.ksMessageBoxResult();	// ��������� ��������� ������ ����� �������
						else
						{
							buf = string.Format("key1 = {0} key2 = {1} key3 = {2} key4 = {3}",
								type.key1, type.key2, type.key3, type.key4);
							kompas.ksMessage(buf);
							buf = string.Format("header = {0} rowsCount = {1} flagVisible = {2} password = {3}",
								type.header, type.rowsCount, type.flagVisible, type.password);
							kompas.ksMessage(buf);
							ksDynamicArray pCol = (ksDynamicArray)type.GetColumns();
							if (pCol != null)
							{
								ShowColumns(pCol, false);
								pCol.ksDeleteArray();
							}
							//					ShowColumns(attrType.columns, FALSE); //���������������� �������
						}
					}
				}
				while(numb != 0);
				//������  ������ �������
				//		DeleteArray(attrType.columns);
			}
		}

		// �������� ��� ��������
		void ChangeType()
		{
			double numb = 0;
			string password = string.Empty;
			//��������� ��� ����������
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);
			int j;

			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			if (type != null)
			{
				type.Init();
				do
				{
					j = kompas.ksReadDouble("������ ����� ���� ��������", 1000, 0, 1e12, ref numb);
					if (j  != 0)
					{
						password = kompas.ksReadString("������ ������ ���� ��������", "");
						//������� ��� ��������
						if (attr.ksGetAttrType(numb, nameFile, type)  != 1)
							kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������
						else 
						{
							type.password = password;
							ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
							ksColumnInfoParam par1 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
							ksColumnInfoParam parN = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
							if (arr != null && par1 != null && parN != null)
							{
								par1.Init();
								parN.Init();
								// ����� �������
								int n = arr.ksGetArrayCount();
								//������� ������ �������
								arr.ksGetArrayItem(0, par1);
								//������� ��������� �������
								arr.ksGetArrayItem(n-1, parN);
								//������� ������ �������
								arr.ksSetArrayItem(0, parN);
								//������� ��������� �������
								arr.ksSetArrayItem(n-1, par1);

								//�������  ��� ��������  �� �����
								double numbType = attr.ksSetAttrType(numb, nameFile, type, password);
								if (numbType > 1)
								{
									string buf = string.Empty;
									buf = string.Format("numbType = {0}", numbType);
									kompas.ksMessage(buf);
								}
								else
									kompas.ksMessageBoxResult();  // ��������� ���������� - ������� ��������� ������ ����� �������
								arr.ksDeleteArray();
							}
						}
					}
				} while(j  != 0);
			}
		}

		// �������� ������� ����,����������� �� ������� FuncTypeAttr
		void NewAttr() 
		{
			ksAttributeParam attrPar = (ksAttributeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_Attribute);
			ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam); 
			ksDynamicArray fVisibl = (ksDynamicArray)kompas.GetDynamicArray(23);
			ksDynamicArray colKeys = (ksDynamicArray)kompas.GetDynamicArray(23);
			if (attrPar  != null && usPar  != null && fVisibl  != null && colKeys  != null)
			{
				attrPar.Init();
				usPar.Init();
				attrPar.SetValues(usPar);
				attrPar.SetColumnKeys(colKeys);
				attrPar.SetFlagVisible(fVisibl);
				attrPar.key1 = 1;
				attrPar.key2 = 10;
				attrPar.key3 = 100;
				attrPar.password = "111";

				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);
				if (item  != null && arr  != null)
				{
					usPar.SetUserArray(arr);
					item.Init();
					item.doubleVal = 987654321.0;
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.strVal = "qwerty";
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.longVal = 999991;
					arr.ksAddArrayItem(-1, item);

					item.Init();
					item.uCharVal = 1;
					fVisibl.ksAddArrayItem(-1, item);
					fVisibl.ksAddArrayItem(-1, item);
					fVisibl.ksAddArrayItem(-1, item);
				}

				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.prompt = "������� ������";
					double x = 0, y = 0;
					reference pObj = 0;
					int j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) != 0)
						{
							doc.ksLightObj(pObj, 1);
							//��������� ��� ����������
							string nameFile = kompas.ksChoiceFile("*.lat", null, false);
							double numb = 0;
							j = kompas.ksReadDouble("������ ����� ���� ��������", 1000, 0, 1e12, ref numb);
							if (j != 0)
							{
								reference pAttr = attr.ksCreateAttr(pObj, attrPar, numb, nameFile);
								if (pAttr == 0)
									kompas.ksMessageBoxResult();  // ��������� ���������� - ������� ��������� ������ ����� �������
							}
							doc.ksLightObj(pObj, 0);
						}
					}
				}
			}
		}

		// ������� ������ �������  � ������� �������
		void DelObjAttr() 
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null) 
			{
				info.Init();
				info.prompt = "������� ������";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) == 1)
						{
							doc.ksLightObj(pObj, 1);
							//�������� �������� ��� �������� �� ��������� �������
							ksIterator iter = (ksIterator)kompas.GetIterator();
							if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0))
							{
								//������ �� ������ �������
								reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);
								if (pAttr != 0)
								{
									string password = kompas.ksReadString("������ ������ ���� ��������", "");
									if (attr.ksDeleteAttr(pObj, pAttr, password) != 1) 
										kompas.ksMessageBoxResult();
								}
								else 
									kompas.ksMessage("������� �� ������");
								doc.ksLightObj(pObj, 0);
							}
						}
					}
				} while (j != 0);
			}
		}

		// �������  �������  ���� double_str_long
		void ReadObjAttr() 
		{
			bool res = false;
			ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam); 
			if (usPar != null)
			{
				usPar.Init();
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);
				if (item != null && arr != null)
				{
					usPar.SetUserArray(arr);
					item.Init();
					item.doubleVal = 987654321.0;
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.strVal = "qwerty";
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.longVal = 999991;
					arr.ksAddArrayItem(-1, item);
					res = true;
				}
			}
			if (res)
			{
				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.prompt = "������� ������";
					double x = 0, y = 0;
					int j;
					do
					{
						j = doc.ksCursor(info, ref x, ref y, null);
						if (j != 0)
						{
							reference pObj = doc.ksFindObj(x, y, 1e6);
							if (doc.ksExistObj(pObj) != 0)
							{
								doc.ksLightObj(pObj, 1);
								//�������� �������� ��� �������� �� ��������� �������
								ksIterator iter = (ksIterator)kompas.GetIterator();
								if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0)) 
								{
									//������ �� ������ �������
									reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);	
									if (pAttr != 0)
									{
										kompas.ksMessage("��� � ����� ��������");
										int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
										double numb = 0;
										attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out numb);
										string buf = string.Format("k1 = {0} k2 = {1} k3 = {2} k4 = {3} numb = {4}",
											k1, k2, k3, k4, numb);
										kompas.ksMessage(buf);

										kompas.ksMessage("������ ��������");
										attr.ksGetAttrRow(pAttr, 0, 0, 0, usPar);

										kompas.ksMessage("������� ������ ��������");
										ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
										ksDynamicArray arr = (ksDynamicArray)usPar.GetUserArray();
										if (item != null && arr != null)
										{
											item.Init();
											item.doubleVal = numb;
											arr.ksSetArrayItem(0, item);
											item.Init();
											item.strVal = "1234567\nasdfgh\nzxcvb";
											arr.ksSetArrayItem(1, item);
											item.Init();
											item.longVal = 888881;
											arr.ksSetArrayItem(2, item);
											attr.ksSetAttrRow(pAttr, 0, 0, 0, usPar, "111");
										}
									}
									else 
										kompas.ksMessage("������� �� ������");
								}
								doc.ksLightObj(pObj, 0);
							}
						}
					}
					while (j != 0);
				}
			}
		}

		// �����������  �������
		void ShowObjAttr()
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null)
			{
				info.Init();
				info.prompt = "������� ������";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) != 0)
						{
							doc.ksLightObj(pObj, 1);
							attr.ksChoiceAttr(pObj);
							doc.ksLightObj(pObj, 0);
						}
					}
				}
				while (j != 0);
			}
		}

		void ShowLib()
		{
			//��������� ��� ����������
			string nameFile = kompas.ksChoiceFile("*.lat", null, false);
			double numb = attr.ksChoiceAttrTypes(nameFile);
			if (numb > 1)  
			{
				string buf = string.Format("numbType = {0}", numb);
				kompas.ksMessage(buf);
			}
		}

		void ShowType()
		{
			//��������� ��� ����������
			string nameFile = kompas.ksChoiceFile("*.lat", null, false);
			string password = string.Empty;
			double numb = 0;
			int j = kompas.ksReadDouble("������ ����� ���� ��������", 1000, 0, 1e12, ref numb);
			if (j != 0)
			{
				password = kompas.ksReadString("������ ������ ���� ��������", "");
				attr.ksViewEditAttrType(nameFile, 2, numb, password);
			}
		}

		// �������� � ������� �� ��������� � ������
		// key1 = 10 � ������ ���������� ������� � ����� ��� ������� ��������
		void WalkFromObjWithAttr()
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null)
			{
				info.Init();
				info.prompt = "������� ������";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) == 1)
						{
							//�������� �������� ��� �������� �� ��������� � ������ 10
							ksIterator iter = (ksIterator)kompas.GetIterator();
							if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0)) 
							{
								doc.ksLightObj(pObj, 1);
								//������ �� ������ �������
								reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);	
								if (pAttr != 0)
								{
									do
									{
										attr.ksViewEditAttr(pAttr, 1, string.Empty);
										pAttr = iter.ksMoveAttrIterator("N", ref pObj);
									}
									while(pAttr != 0);
								}
								doc.ksLightObj(pObj, 0);
							}
						}
					}
				}
				while (j != 0);
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
