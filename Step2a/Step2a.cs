using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6Constants;

namespace Steps.NET
{
	// ����� Step2a - ������ �������������� �����
	// 1. ������ �����                           - StrIndefiniteArray
	// 2. ������ �������������� �����            - PointIndefiniteArray
	// 3. ������ ����� ������� "�����"           - TextIndefiniteArray
	// 4. ������ ������� ���� ��������           - AttrIndefiniteArray
	// 5. ������ ���������                       - PolyLineArray
	// 6. ������ ���������� ���������������      - RectArray
	// 7. ������ �������� ������������           - UserDataArray
	// 8. ������ ����������� ������ ������������ - UserClassArray

	public class Step2a
	{
		private KompasObject kompas = null;

		[return: MarshalAs(UnmanagedType.BStr)]
		public string GetLibraryName()
		{
			return "Step2a - ������ �������������� �����";
		}

		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject)kompas_;
			if (kompas != null)
			{
				switch (command)
				{
					case 1: StrIndefiniteArray(); break;	// ������ �����
					case 2: PointIndefiniteArray(); break;	// ������ �������������� �����
					case 3: TextIndefiniteArray(); break;	// ������ ����� ������� "�����"
					case 4: AttrIndefiniteArray(); break;	// ������ ������� ���� ��������
					case 5: PolyLineArray(); break;			// ������ ���������
					case 6: RectArray(); break;				// ������ ���������� ���������������
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
					result = "������ ������� �����";
					command = 1;
					break;
				case 2:
					result = "������ �����";
					command = 2;
					break;
				case 3:
					result = "������ ����� ������";
					command = 3;
					break;
				case 4:
					result = "������ ������� ���� ��������";
					command = 4;
					break;
				case 5:
					result = "������ ���������";
					command = 5;
					break;
				case 6:
					result = "������ ���. ���������������";
					command = 6;
					break;
				case 7:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// ������ ������������ ��� �������� � ��������� ��������� ��������� ������,
		// ���������� ��   ���������� �����, ������� � ���� �������
		// ������� �� ���������, � ������� �����������
		// (� ��������, � ����������, ��������� � �. �.)
		private void TextIndefiniteArray()
		{
			string buf = string.Empty;
			ksTextLineParam par = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam par1 = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			// ������� ������ ����� ������
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
			if (par != null && par1 != null && p != null)
			{
				par.Init();
				par1.Init();
				// ������ ��������� ������ ������
				ksDynamicArray p1 = (ksDynamicArray)par.GetTextItemArr();
				if (p1 != null)
				{
					ksTextItemFont font = (ksTextItemFont)par1.GetItemFont();
					if (font != null)
					{
						// ������� ������ ������ ������
						font.height = 10;   // ������ ������
						font.ksu = 1;       // ������� ������
						font.color = 1000;  // ����
						font.bitVector = 1; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "1 ���������� 1 ������";
						// �������� 1-� ����������  � ������ ���������
						p1.ksAddArrayItem(-1, par1);

						font.height = 20;   // ������ ������
						font.ksu = 2;       // ������� ������
						font.color = 2000;  // ����
						font.bitVector = 2; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "2 ���������� 1 ������";
						// �������� 2-� ����������  � ������ ���������
						p1.ksAddArrayItem(-1, par1);
						par.style = 1;

						// 1-� ������ ������ ������� �� ���� ��������� ������� ������ ������ �
						// ������ ����� ������
						p.ksAddArrayItem(-1, par);

						// �������� ������ ���������, ����� ������������ ��� �������� ������
						// ������ ������
						p1.ksClearArray();

						// ������� ������ ������ ������
						font.height = 30;   // ������ ������
						font.ksu = 3;       // ������� ������
						font.color = 3000;  // ����
						font.bitVector = 3; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "1 ���������� 2 ������";
						// �������� 1-� ����������  � ������ ���������
						p1.ksAddArrayItem(-1, par1);

						font.height = 40;   // ������ ������
						font.ksu = 4;       // ������� ������
						font.color = 4000;  // ����
						font.bitVector = 4; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "2 ���������� 2 ������";
						// �������� 2-� ����������  � ������ ���������
						p1.ksAddArrayItem(-1, par1);

						par.style = 2;

						// 2-� ������ ������ ������� �� ���� ��������� ������� ������ ������ �
						// ������ ����� ������		 }
						p.ksAddArrayItem(-1, par);

						kompas.ksMessageBoxResult();

						int n = p.ksGetArrayCount();
						buf = string.Format(" n = {0} ", n);
						kompas.ksMessage(buf);
						// ���������� ������ ����� ������
						for (int i = 0; i < n; i++)
						{  // ���� �� ������� ������
							p.ksGetArrayItem(i, par);
							buf = string.Format("i = {0}: style = {1},", i, par.style);
							kompas.ksMessage(buf);

							int n1 = p1.ksGetArrayCount();
							for (int j = 0; j < n1; j++)
							{  // ���� �� ����������� ������ ������
								p1.ksGetArrayItem(j, par1);
								buf = string.Format("j = {0}:  h = {1:0.#}, s = {2}", j, font.height, par1.s);
								kompas.ksMessage(buf);
							}
						}

						kompas.ksMessageBoxResult(); // ��������� ��������� ������ ����� �������

						// ������� ������ ���������� � ������ ������
						p.ksGetArrayItem(0, par);
						font.height = 50;   // ������ ������
						font.ksu = 1;       // ������� ������
						font.color = 1000;  // ����
						font.bitVector = 1; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "2 ����. 1 ���.";
						p1.ksSetArrayItem(1, par1);
						par.style = 3;
						p.ksSetArrayItem(0, par);

						// ������� ������ ���������� � ������ ������
						p.ksGetArrayItem(1, par);
						font.height = 60;   // ������ ������
						font.ksu = 1;       // ������� ������
						font.color = 1000;  // ����
						font.bitVector = 1; // ������� ������ (������, �������, �������������, ��� ��������� �����(�����, ����������, ��������� ���� �����))
						par1.s = "1 ����. 2 ���.";
						p1.ksSetArrayItem(0, par1);
						par.style = 4;
						p.ksSetArrayItem(1, par);

						n = p.ksGetArrayCount();
						// ���������� ������ ����� ������
						for (int i = 0; i < n; i++)
						{  // ���� �� ������� ������
							p.ksGetArrayItem(i, par);
							buf = string.Format("i = {0}: style = {1}, ", i, par.style);
							kompas.ksMessage(buf);

							int n1 = p1.ksGetArrayCount();
							for (int j = 0; j < n1; j++)
							{  // ���� �� ����������� ������ ������
								p1.ksGetArrayItem(j, par1);
								buf = string.Format("j = {0}:  h = {1:0.#}, s = {2}", j, font.height, par1.s);
								kompas.ksMessage(buf);
							}
						}

						kompas.ksMessageBoxResult(); // ��������� ��������� ������ ����� �������

						p1.ksDeleteArray();
						p.ksDeleteArray();
					}
				}
			}
		}


		// ������ ������������ ��� �������� �������������� ����� ����  MathPointParam
		void PointIndefiniteArray()
		{
			string buf = string.Empty;
			ksMathPointParam par = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			// ������� ������
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (p != null && par != null)
			{
				// ��������� ������
				par.Init();
				par.x = 10;
				par.y = 10;
				p.ksAddArrayItem(-1, par);
				par.x = 20;
				par.y = 20;
				p.ksAddArrayItem(-1, par);
				par.x = 30;
				par.y = 30;
				p.ksAddArrayItem(-1, par);

				kompas.ksMessageBoxResult();

				// ���������� ������
				int n = p.ksGetArrayCount();
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, par);
					buf = string.Format("i = {0},  x = {1:0.#}, y = {2:0.#}",
						i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				// ������� ��������� 1-�� ��������
				par.x = 50;
				par.y = 50;
				p.ksSetArrayItem(1, par);

				// ������� ��������� 0-�� ��������
				par.x = 60;
				par.y = 60;
				p.ksSetArrayItem(0, par);

				// ���������� ������
				n = p.ksGetArrayCount();
				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, par);
					buf = string.Format("i = {0}:  x = {1:0.#}, y = {2:0.#}",
						i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				kompas.ksMessageBoxResult();

				p.ksDeleteArray();
			}
		}


		// ������ ������������ ��� �������� �����
		void StrIndefiniteArray()
		{
			string buf = string.Empty;
			// ������� ������
			ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);
			if (p != null && charS != null)
			{
				charS.Init();
				// �������� ������
				charS.str = "12345";
				p.ksAddArrayItem(-1, charS);
				charS.str = "67890";
				p.ksAddArrayItem(-1, charS);
				charS.str = "qwerty";
				p.ksAddArrayItem(-1, charS);

				kompas.ksMessageBoxResult();

				// ���������� ������
				int n = p.ksGetArrayCount();
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, charS);
					kompas.ksMessage(charS.str);
				}

				// ��������� �� ������� �� 1
				p.ksExcludeArrayItem(1);
				n = p.ksGetArrayCount();

				// ���������� ������
				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, charS);
					kompas.ksMessage(charS.str);
				}

				kompas.ksMessageBoxResult();

				p.ksDeleteArray();
			}
		}


		private void ShowColumns(ksDynamicArray pCol, bool fl)
		{
			ksColumnInfoParam par = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (par != null)
			{
				par.Init();

				string buf = string.Empty;
				string s = string.Empty;
				if (fl)
					s = "���������";

				int n = pCol.ksGetArrayCount();

				for (int i = 0; i < n; i ++)
				{
					if (pCol.ksGetArrayItem(i, par) != 1)
						kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������
					else
					{
						//������� ���� ������� �� ���������
						buf = string.Format("{0} i = {1} header = {2} type = {3} def = {4} flagEnum = {5}",
							s, i, par.header,
							par.type, par.def, par.flagEnum);
						kompas.ksMessage(buf);
						if (par.type == ldefin2d.RECORD_ATTR_TYPE)
						{
							// ���������
							pCol = (ksDynamicArray)par.GetColumns();
							if (pCol != null)
							{
								ShowColumns(pCol, true);
								//pCol.ksDeleteArray();
							}
						}
						else
						{
							if (par.flagEnum)
							{
								// ������� ������ ������������
								ksDynamicArray fEnum = (ksDynamicArray)par.GetFieldEnum();
								if (fEnum != null)
								{
									int n1 = fEnum.ksGetArrayCount();
									kompas.ksMessage("������ ������������");
									ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
									for (int i1 = 0; i1 < n1; i1 ++)
										if (fEnum.ksGetArrayItem(i1, charS) != 1)
											kompas.ksMessageBoxResult();  // ��������� ��������� ������ ����� �������
										else
											kompas.ksMessage(charS.str);
								}
							}
						}
					}
				}
			}
		}


		// ������ ������������ ��� �������� � ��������� ���� ��������,
		// ������� ����� �������� �� ���������� �������
		void AttrIndefiniteArray()
		{
			// �������� ������ �� 3 �������
			// ������ ������� ���������  int � �������������� ���������� ( 100, 200, 300 )
			// ������ ������� - ������ ������������� ���������
			// struct {
			//   double ;// ������������� �������� 123456789
			//   long   ;// ������������� �������� 1000000
			//   char   ;// ������������� �������� 10
			// }
			// ������ ������� ������ �������� ������������� �������� "text"

			ksColumnInfoParam parCol1 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parCol2 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parCol3 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parStruct = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksDynamicArray pCol = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.ATTR_COLUMN_ARR);
			ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			if (parCol1 != null && parCol2 != null && parCol3 != null
				&& parStruct != null && charS != null && pCol != null)
			{
				// ������ �������
				parCol1.Init();
				parCol1.header = "int";					// �������o�-����������� �������
				parCol1.type = ldefin2d.INT_ATTR_TYPE;	// ��� ������ � ������� - ��.����
				parCol1.key = 0;						// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
				parCol1.def = "100";					// �������� �� ���������
				parCol1.flagEnum = true;				// ���� ���������� �����, ����� �������� ���� �������� ����� ���������� �� ������� ������������� ��������

				ksDynamicArray pEnum = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);
				if (pEnum != null)
				{
					//�������� ������ ������������� �������� ��� ������ �������
					charS.Init();
					charS.str = "100";
					pEnum.ksAddArrayItem(-1, charS);
					charS.str = "200";
					pEnum.ksAddArrayItem(-1, charS);
					charS.str = "300";
					pEnum.ksAddArrayItem(-1, charS);

					parCol1.SetFieldEnum(pEnum);	// ������ �������������� ����� ������������ (������)
				}

				pCol.ksAddArrayItem(-1, parCol1);

				// ������ �������
				parCol2.Init();
				parCol2.header = "struct";					// �������o�-����������� �������
				parCol2.type = ldefin2d.RECORD_ATTR_TYPE;	// ��� ������ � ������� - ��.����
				parCol2.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
				parCol2.def = "\0";							// �������� �� ���������
				parCol2.flagEnum = false;					// ���� ���������� �����, ����� �������� ���� �������� ����� ���������� �� ������� ������������� ��������

				ksDynamicArray pStruct = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.ATTR_COLUMN_ARR);
				if (pStruct != null)
				{
					// �������� ������ ������� ��� ���������
					// ������ �������  ���������
					parStruct.Init();
					parStruct.header = "double";				// �������o�-����������� �������
					parStruct.type = ldefin2d.DOUBLE_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					parStruct.def = "123456789";				// �������� �� ���������
					pStruct.ksAddArrayItem(-1, parStruct);

					// ������ ������� ���������
					parStruct.Init();
					parStruct.header = "long ";					// �������o�-����������� �������
					parStruct.type = ldefin2d.LINT_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					parStruct.def = "1000000";					// �������� �� ���������
					pStruct.ksAddArrayItem(-1, parStruct);

					// ������ ������� ���������
					parStruct.Init();
					parStruct.header = "char";					// �������o�-����������� �������
					parStruct.type = ldefin2d.CHAR_ATTR_TYPE;	// ��� ������ � ������� - ��.����
					parStruct.def = "10";						// �������� �� ���������
					pStruct.ksAddArrayItem(-1, parStruct);

					parCol2.SetColumns(pStruct);	// ������ �������������� ����� ���������� � �������� ��� ������
				}

				pCol.ksAddArrayItem(-1, parCol2);

				// ������  �������
				parCol3.Init();
				parCol3.header = "string";					// �������o�-����������� �������
				parCol3.type = ldefin2d.STRING_ATTR_TYPE;	// ��� ������ � ������� - ��.����
				parCol3.key = 0;							// �������������� �������, ������� �������� �������� ��� ���������� � ���������� �����
				parCol3.def = "text";						// �������� �� ���������

				pCol.ksAddArrayItem(-1, parCol3);

				kompas.ksMessageBoxResult();	// ��������� ��������� ������ ����� �������

				// ���������� ������ �������
				ShowColumns(pCol, false);	// ������� ������������

				kompas.ksMessageBoxResult();	// ��������� ��������� ������ ����� �������

				// ��������  ������� ������� 2->1 1->3 3->2
				pCol.ksSetArrayItem(0, parCol2);
				pCol.ksSetArrayItem(2, parCol1);
				pCol.ksSetArrayItem(1, parCol3);

				// ���������� ������ �������
				ShowColumns(pCol, false);	// ������� ������������

				kompas.ksMessageBoxResult();	// ��������� ��������� ������ ����� �������

				pStruct.ksDeleteArray();
				pEnum.ksDeleteArray();
				pCol.ksDeleteArray();
			}
		}


		// ������ ��������� ��� ������ �������� �������������� �����
		void PolyLineArray()
		{
			string buf = string.Empty;
			ksMathPointParam par = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			//�������� ������ �����
			ksDynamicArray arrPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			//�������� ������ ���������
			ksDynamicArray arrPoly = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POLYLINE_ARR);
			if (par != null && arrPoint != null && arrPoly != null)
			{
				//�������� ������ �����
				par.x = 10;
				par.y = 10;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 100;
				par.y = 100;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 1000;
				par.y = 1000;
				arrPoint.ksAddArrayItem(-1, par);
				//������� ������ ����� � ������ ���������
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				//�������� ������ �����
				arrPoint.ksClearArray();
				par.x = 20;
				par.y = 20;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 200;
				par.y = 200;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 2000;
				par.y = 2000;
				arrPoint.ksAddArrayItem(-1, par);
				//������� ������ ����� � ������ ���������
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				//�������� ������ �����
				arrPoint.ksClearArray();
				par.x = 30;
				par.y = 30;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 300;
				par.y = 300;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 3000;
				par.y = 3000;
				arrPoint.ksAddArrayItem(-1, par);
				//������� ������ ����� � ������ ���������
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				int count = arrPoly.ksGetArrayCount();
				buf = string.Format("n = {0}", count);
				kompas.ksMessage(buf);

				arrPoly.ksGetArrayItem(1, arrPoint);
				//���������� 2-�� ������� ������� ���������
				count = arrPoint.ksGetArrayCount();
				for (int i = 0; i < count; i ++)
				{
					arrPoint.ksGetArrayItem(i, par);
					buf = string.Format("i = {0}, x = {1}, y = {2}", i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				//������� � 2 -�� �������� ������� ��������� 2-�� �������
				par.x = 50;
				par.y = 50;
				arrPoint.ksSetArrayItem(1, par);
				par.x = 500;
				par.y = 500;
				arrPoint.ksSetArrayItem(0, par);
				arrPoly.ksSetArrayItem(1, arrPoint);

				count = arrPoly.ksGetArrayCount();
				for (int i = 0; i < count; i++)
				{
					arrPoly.ksGetArrayItem(i, arrPoint);
					int n = arrPoint.ksGetArrayCount();
					for (int j = 0; j < n; j++)
					{
						arrPoint.ksGetArrayItem(j, par);
						buf = string.Format("i = {0}, j = {1}, x = {2}, y = {3}",
							i, j, par.x, par.y);
						kompas.ksMessage(buf);
					}
				}
				kompas.ksMessageBoxResult();
				arrPoint.ksDeleteArray();
				arrPoly.ksDeleteArray();
			}
		}


		// ������ �������������� ����� ���������� ���������������-(��������� RectParam)
		void RectArray()
		{
			ksRectParam par = (ksRectParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectParam); // ��������� ��������������
			ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.RECT_ARR);    // ������� ������
			ksMathPointParam pBot = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksMathPointParam pTop = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (arr != null && par != null && pBot != null && pTop != null)
			{
				//��������� ������
				pTop.x = 10;
				pTop.y = 10;
				par.SetpTop(pTop);
				pBot.x = 20;
				pBot.y = -10;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				pTop.x = 20;
				pTop.y = 50;
				par.SetpTop(pTop);
				pBot.x = 50;
				pBot.y = 10;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				pTop.x = 20;
				pTop.y = 150;
				par.SetpTop(pTop);
				pBot.x = 50;
				pBot.y = 110;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);
				kompas.ksMessageBoxResult();

				//����������� ������
				int n = arr.ksGetArrayCount();

				string buf = string.Empty;
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					arr.ksGetArrayItem(i, par);
					pBot = (ksMathPointParam)par.GetpBot();
					pTop = (ksMathPointParam)par.GetpTop();
					buf = string.Format("i = {0}, x1 = {1:0.#}, y1 = {2:0.#},\n x2 = {3:0.#}, y2 = {4:0.#}, ",
						i, pTop.x, pTop.y,
						pBot.x, pBot.y);
					kompas.ksMessage(buf);
				}

				//����������� ������
				pTop.x = -20;
				pTop.y = -50;
				par.SetpTop(pTop);
				pBot.x = 20;
				pBot.y = -10;
				par.SetpBot(pBot);
				arr.ksSetArrayItem(1, par);

				pTop.x = 0;
				pTop.y = 0;
				par.SetpTop(pTop);
				pBot.x = 10;
				pBot.y = -20;
				par.SetpBot(pBot);
				arr.ksSetArrayItem(0, par);

				pTop.x = 5;
				pTop.y = 5;
				par.SetpTop(pTop);
				pBot.x = 25;
				pBot.y = 0;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				//���������� ������
				n = arr.ksGetArrayCount();
				for (int i = 0; i < n; i ++)
				{
					arr.ksGetArrayItem(i, par);
					pBot = (ksMathPointParam)par.GetpBot();
					pTop = (ksMathPointParam)par.GetpTop();
					buf = string.Format("i = {0}, x1 = {1:0.#}, y1 = {2:0.#},\n x2 = {3:0.#}, y2 = {4:0.#},",
						i, pTop.x, pTop.y, pBot.x, pBot.y);
					kompas.ksMessage(buf);
				}
				kompas.ksMessageBoxResult();

				//������� ������
				arr.ksDeleteArray();
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
