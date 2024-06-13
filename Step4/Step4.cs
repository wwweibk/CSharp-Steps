using Kompas6API5;
using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using KAPITypes;
using Kompas6Constants;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step4 - ������ ������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step4
	{
		KompasObject kompas = null;
		ksDocument2D doc = null;
		int type = 0, flag = 0;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)]	public string GetLibraryName()
		{
			return "Step4 - ������ ������";
		}


		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject)kompas_;
			if (kompas != null)
			{
				switch (command)
				{
					case 1: DrawTxtDB();		break; // ������ � ��
					case 3: LongIntInput();		break; // ���� �������� ������
					case 4: FileNameSelect();	break; // ����� ����� �����
					case 8: WorkRelativePath();	break; // ������ � �������������� ������ ������
					default:
						doc = (ksDocument2D)kompas.ActiveDocument2D();
						if (doc != null && doc.reference > 0)
						{
							switch (command)
							{
								case 2: PlacementCursor();	break; // Placement � Cursor
								case 5: WriteSlideStep();	break; // ������ ��������� ������
								case 6: TestShowDialog();	break; // ���������� �����
								case 7: Queue();			break; // ������ ������ � ���������
								case 9: WorkSystemPath();	break; // ������ � ���������� ����������
							}
						}
						break;
				}
			}
		}


		// ������������ ���� ����������
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "������ � ��";
					command = 1;
					break;
				case 2:
					result = "Placement, Cursor";
					command = 2;
					break;
				case 3:
					result = "���� �������� ������";
					command = 3;
					break;
				case 4:
					result = "����� ����� �����";
					command = 4;
					break;
				case 5:
					result = "�������� �����";
					command = 5;
					break;
				case 6:
					result = "���������� �����";
					command = 6;
					break;
				case 7:
					result = "������ � ���������� ������� ���������";
					command = 7;
					break;
				case 8:
					result = "������������� ���� �����";
					command = 8;
					break;
				case 9:
					result = "������ � ���������� ����������";
					command = 9;
					break;
				case 10:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		#region ��������� ���������
		private struct b_
		{
			public double dr, l;
			public int f;
		}

		private struct con_
		{
			public double a, l;
		}
		#endregion

		// ������ � ��
		void DrawTxtDB()
		{
			con_ con = new con_();
			b_ b = new b_();

			reference bd, r1, r2, r3;
			int i = 1;
			string buf = string.Empty;
			ksDataBaseObject data = (ksDataBaseObject)kompas.DataBaseObject();
			ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par == null || item == null || data == null)
				return;
			par.Init();
			par.SetUserArray(arr);
			item.Init();
			item.doubleVal = 0;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.doubleVal = 0;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.intVal = 0;
			arr.ksAddArrayItem(-1, item);

			string libName = kompas.ksChoiceFile("*.loa", "���� ������(*.loa)|*.loa|��� ����� (*.*)|*.*|", true);
			if (libName != null && libName != string.Empty)
			{
				bd = data.ksCreateDB("TXT_DB"); //������� ������, ������������� ���� ������
				if (data.ksConnectDB(bd, libName) == 1)
				{ //������� ������ ���� � ������������ ����� ������(��� ���������� ����� - ��� �����)
					r1 = data.ksRelation(bd); //������� ��������� - ����� ��� ���������� ���������� �� �������
					data.ksRDouble("dr");//����� ������ �������� ��������,
					data.ksRDouble("L"); //�� ��� � ���������� ����������� ������� �������
					data.ksRInt("");
					data.ksEndRelation();

					//���������� ������ - ��������� ����������� � �����(��������� ����� �������
					//���� ������� � ����������� ������)
					data.ksDoStatement(bd, r1, "1 2 3"); //������� dr - 1, L - 2, ���������� ������� -3
					while (i > 0)
					{
						i = data.ksReadRecord(bd, r1, par); //������� ��������� ������ ���������� � ������� � ��������� b
						if (i > 0)
						{
							arr.ksGetArrayItem(0, item);
							b.dr = item.doubleVal;
							arr.ksGetArrayItem(1, item);
							b.l = item.doubleVal;
							arr.ksGetArrayItem(2, item);
							b.f = item.intVal;
							buf = string.Format("DR = {0}\nL = {1}\nF = {2}", b.dr, b.l, b.f);//item.GetDoubleVal(), item1.GetDoubleVal(), item2.GetIntVal());
							kompas.ksMessage(buf);
						}
					}
					kompas.ksMessage("end");

					i = 1;
/*			par.Init();
	    par.SetUserArray(arr);*/
					arr.ksClearArray();
					item.Init();
					item.strVal = string.Empty;
					arr.ksAddArrayItem(-1, item);
					r2 = data.ksRelation(bd); //������� ��������� - ����� ��� ���������� ���������� �� �������
					data.ksRChar("", 255, 0);
					data.ksEndRelation();

					data.ksDoStatement(bd, r2, "2");//���������� ������ - ��������� ����������� � �����(��������� ����� �������
					while (i > 0)
					{
						i = data.ksReadRecord(bd, r2, par); //������� ��������� ������ ���������� � ������� � ��������� b
						if (i > 0)
						{
							arr.ksGetArrayItem(0, item);
							buf = string.Format("L = {0}", item.strVal);
							kompas.ksMessage(buf);
						}
					}
					kompas.ksMessage("end");

					i = 1;

					arr.ksClearArray();
					item.Init();
					item.doubleVal = 0;
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.doubleVal = 0;
					arr.ksAddArrayItem(-1, item);
					r3 = data.ksRelation(bd); //������� ��������� - ����� ��� ���������� ���������� �� �������
					data.ksRDouble("");
					data.ksRDouble("L");
					data.ksEndRelation();

					data.ksDoStatement(bd, r3, "1 2");//���������� ������ - ��������� ����������� � �����(��������� ����� �������
					data.ksCondition(bd, r3, "L=100||L=150");
					while (i > 0)
					{
						i = data.ksReadRecord(bd, r3, par); //������� ��������� ������ ���������� � ������� � ��������� b
						if (i > 0)
						{
							arr.ksGetArrayItem(0, item);
							con.a = item.doubleVal;
							arr.ksGetArrayItem(1, item);
							con.l = item.doubleVal;
							buf = string.Format("dr = {0}\nL = {1}", con.a, con.l);
							kompas.ksMessage(buf);
						}
					}
					kompas.ksMessage("end");
				}
				data.ksDeleteDB(bd); //������� ������, ������������� ���� ������
			}
		}


		// ���� �������� ������
		void LongIntInput()
		{
			int h1 = 0;
			if (kompas.ksReadInt("������ ������", 10000, -100000, 100000, ref h1) == 1)
			{
				string buf = string.Format("h = {0}", h1);
				kompas.ksMessage(buf);
			}
			else
				kompas.ksMessage("�����");
		}


		// ����� ����� �����
		void FileNameSelect()
		{
			string name = kompas.ksChoiceFile("*.cdw", null, true);
			if (name != string.Empty)
				kompas.ksMessage(name);
			else
				kompas.ksMessage("�����");
		}


		// ������ � �������������� ������ ������
		void WorkRelativePath()
		{
			//��� ��������� �����
			string mainName = kompas.ksChoiceFile("*.*", "��� ����� (*.*)|*.*|", true);
			string fileName = kompas.ksChoiceFile("*.*", "��� ����� (*.*)|*.*|", true);
			if (mainName != string.Empty && mainName != null && fileName != string.Empty && fileName != null)
			{
				// ������������� ����
				string relName = kompas.ksGetRelativePathFromFullPath(mainName,	// ������ ���� � ��������� �����
					fileName); // ������ ���� � ���������� �����

				string mess = "�������� ���� : ";
				mess += mainName;
				mess += "\n";
				mess += "������ ���� : ";
				mess += fileName;
				mess += "\n";
				mess += "������������� ���� : ";
				mess += relName;
				kompas.ksMessage(mess);

				// ������ ����
				string fullName = kompas.ksGetFullPathFromRelativePath(mainName,   // ������ ���� � ��������� �����
					relName); // ������������� ���� � ���������� �����(��� ����� � �������� ������ ����� ����)
				mess = "�������� ���� : ";
				mess += mainName;
				mess += "\n";
				mess += "������������� ���� : ";
				mess += relName;
				mess += "\n";
				mess += "������ ���� : ";
				mess += fullName;
				mess += "\n";
				kompas.ksMessage(mess);
			}
		}


		// Placement � Cursor
		void PlacementCursor()
		{
			if (kompas.ksYesNo("�������� ������� CallBack?") == 1)
				DrawRectCallBack();
			else
				DrawRectNULL();
		}

		void DrawRectCallBack()
		{
			type = 1;
			ksPhantom phan = (ksPhantom)kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom);
			if (phan != null)
			{
				phan.Init();
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				if (t1 != null)
				{
					t1.Init();
					t1.scale_ = 1;
					t1.gr = doc.ksNewGroup(1);   // ��������� ������
					doc.ksCircle(0, 0, 20, 1);
					doc.ksEndGroup();

					double x = 0, y = 0, ang = 0;
					ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
					if (info != null)
					{
						info.Init();
						info.commandsString = "!������� !�����������";
						// ��������� ����� �������� ������� ��� Placement
						info.SetCallBackP("CALLBACKPROCPLACEMENT", 0, this);
						doc.ksPlacement(info, ref x, ref y, ref ang, phan);

						t1.gr = doc.ksNewGroup(1);   // ��������� ������
						doc.ksCircle(0, 0, 20, 1);
						doc.ksEndGroup();

						// ��������� ����� �������� ������� ��� Cursor
						info.SetCallBackC("CALLBACKPROCCURSOR", 0, this);
						doc.ksCursor(info, ref x, ref y, phan);
					}
				}
			}
		}

		void DrawRectNULL()
		{
			int type = 1, flag = 1;
			int j = 1;
			ksPhantom phan = (ksPhantom)kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom);
			if (phan != null)
			{
				phan.Init();
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				if (t1 != null)
				{
					t1.Init();
					t1.scale_ = 1;
					t1.gr = 0;   // ��������� ������

					double x = 0, y = 0, ang = 0;
					ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
					if (info != null)
					{
						info.Init();
						while (j != 0)
						{

							if (t1.gr != 0)
								doc.ksDeleteObj(t1.gr);

							t1.gr = doc.ksNewGroup(1); // ��������� ������
							if ((flag == 1 && j == 1) || (flag == 2 && j == 2))
								type = 3;

							switch (type)
							{
								case 1:
									doc.ksCircle(0, 0, 20, 1);
									info.commandsString = "!������� !����������� ";
									flag = 1;
									break;
								case 2:
									doc.ksLineSeg(-10, 0, 10, 0, 1);
									doc.ksLineSeg(10, 0, 0, 20, 1);
									doc.ksLineSeg(0, 20, -10, 0, 1);
									info.commandsString = "!���������� !������� ";
									flag = 2;
									break;
								case 3:
									doc.ksLineSeg(-10, 0, 10, 0, 1);
									doc.ksLineSeg(10, 0, 10, 20, 1);
									doc.ksLineSeg(10, 20, -10, 20, 1);
									doc.ksLineSeg(-10, 20, -10, 0, 1);
									info.commandsString = "!���������� !����������� ";
									flag = 0;
									break;
							}

							doc.ksEndGroup();

							j = doc.ksPlacement(info, ref x, ref y, ref ang, phan);
							switch (j)
							{
								case 1:
								case 2: type = j; break;
								case -1:	//��������� � ������
									doc.ksMoveObj(t1.gr, x, y);
									if (Math.Abs(t1.angle) > 0.001)
										doc.ksRotateObj(t1.gr, x, y, ang);
									doc.ksStoreTmpGroup(t1.gr);	// ��������� ��������� ������ � ���
									doc.ksClearGroup(t1.gr, true);
									break;
							}
						}
					}
				}
			}
		}


		// ������ ��������� ������
		void WriteSlideStep()
		{
			//������� ���� ��� ������
			string name = kompas.ksSaveFile("*.rc", null, null, false);
			if (name != string.Empty)
			{
				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.commandsString = "������� ����� �������� ������";
					// ����� �������� ������ - ������� ����� ���� ����������� �������������� ������
					double x = 0, y = 0;
					if (doc.ksCursor(info, ref x, ref y, null) != 0)
					{
						int slideID = 0;
						if (kompas.ksReadInt("������� ������������� ������", 100, 0, 32000, ref slideID) == 1)
						{
							if (kompas.ksWriteSlide(name, slideID, x, y) != 1)
								kompas.ksError("������ �������������� �����");
							doc.ksClearGroup(0, true);
						}
					}
				}
			}
		}


		// ���������� �����
		private void TestShowDialog()
		{
			FrmTest.Instance.Doc = doc;
			FrmTest.Instance.ShowDialog();
		}


		// ������ ������ � ���������
		void Queue()
		{
			kompas.ksEnableTaskAccess(0); //��������� ������ � ������
			for (int i = 0; i < 10000; i ++)
			{
				doc.ksLineSeg(10, 10 + i, 20, 10 + i, 1);
				if ((i % 100) == 0)
				{
					System.Windows.Forms.Application.DoEvents();
				}
			}
			kompas.ksEnableTaskAccess(1); //��������� ������ � ������
		}


		// ������ � ���������� ����������
		void WorkSystemPath()
		{
			string[] catalogName = new string[]{"������� ��������� ������", "������� ���������", "������� ��������� ������", "������� ������������", "INI-����" };
			// ������������  ������ ���� � ��������� �����
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null)
			{
				info.Init();
				info.title = "�������� ������ �������";
				info.commandsString = "!��������� !���������� !��������� !������������ !INI-���� ";
				info.prompt = "�������� ������ �������";
				int j;
				string buf = "user.ttt";
				string fileName = string.Empty;
				int typeCatalog = 0;
				do
				{
					j = doc.ksCommandWindow(info);
					if (j > 0)
					{
						switch (j)
						{
							case 1: typeCatalog = ldefin2d.sptSYSTEM_FILES;	break; // ������������ �������� ��������� ������
							case 2: typeCatalog = ldefin2d.sptLIBS_FILES;	break; // ������������ �������� ������ ���������
							case 3: typeCatalog = ldefin2d.sptTEMP_FILES;	break; // ������������ �������� ���������� ��������� ������
							case 4: typeCatalog = ldefin2d.sptCONFIG_FILES;	break; // ������������ �������� ���������� ������������ �������
							case 5: typeCatalog = ldefin2d.sptINI_FILE;		break; // ������������ ������� ����� INI-����� �������
						}
						// ������ ����
						fileName = kompas.ksGetFullPathFromSystemPath(buf, // ������������� ���� � �����(��� ���������� ����)
							typeCatalog); // ���� �������������� ���� ��. ksSystemPath
						string mess = "������ ���� � ����� user.ttt \n";
						mess += catalogName[j - 1];
						mess += " :\n";
						mess += fileName;
						kompas.ksMessage(mess);

						// ������������� ����
						string relName = kompas.ksGetRelativePathFromSystemPath(fileName,       // ������ ���� � �����
							typeCatalog); // ���� �������������� ���� ��. ksSystemPath
						mess = "������������� ���� � ����� \n";
						mess += fileName;
						mess += "\n";
						mess += catalogName[j - 1];
						mess += " :\n";
						mess += relName;
						kompas.ksMessage(mess);

					}
				}
				while (j > 0);
			}
		}


		// ������� �������� �����, ���������� �� Cursor
		public int CALLBACKPROCCURSOR(int comm,
			ref double x, ref double y,
			[MarshalAs(UnmanagedType.LPStruct)] object rInfo,
			[MarshalAs(UnmanagedType.LPStruct)] object rPhan,
			int dynamic)
		{
			ksRequestInfo info = (ksRequestInfo)rInfo;
			ksPhantom phan = (ksPhantom)rPhan;

			if (info != null && phan != null)
			{
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				switch (comm)
				{
					case 1:
					case 2:
						type = comm;
						break;
					case -1: // ��������� � ������
						{
							doc.ksMoveObj(t1.gr, x, y);
							if (Math.Abs(t1.angle) > 0.001)
								doc.ksRotateObj(t1.gr, x, y, t1.angle);
							doc.ksStoreTmpGroup(t1.gr);
							doc.ksClearGroup(t1.gr, true);
						}
						break;
				}

				// ������ ��� ������� ������ ���� ��������� � ����������� ��� ��������� ���� ���������
				if (t1.gr != 0)
					doc.ksDeleteObj(t1.gr);
				t1.gr = doc.ksNewGroup(1); // ��������� ������

				if ((flag == 1 && comm == 1) || (flag == 2 && comm == 2))
					type = 3;

				// ����������� �� ������ ����������� �� � ���� ��� �������
				switch (type)
				{
					case 1:
						doc.ksCircle(0, 0, 20, 1);
						info.commandsString = "!������� !����������� ";
						flag = 1;
						break;
					case 2:
						doc.ksLineSeg(-10, 0, 10, 0, 1);
						doc.ksLineSeg(10, 0, 0, 20, 1);
						doc.ksLineSeg(0, 20, -10, 0, 1);
						info.commandsString = "!���������� !������� ";
						flag = 2;
						break;
					case 3:
						doc.ksLineSeg(-10, 0, 10, 0, 1);
						doc.ksLineSeg(10, 0, 10, 20, 1);
						doc.ksLineSeg(10, 20, -10, 20, 1);
						doc.ksLineSeg(-10, 20, -10, 0, 1);
						info.commandsString = "!���������� !����������� ";
						flag = 0;
						break;
				}
				doc.ksEndGroup();
			}

			return 1;
		}


		// ������� �������� �����, ���������� �� Placement
		public int CALLBACKPROCPLACEMENT(int comm,
			ref double x, ref double y,
			ref double ang,
			[MarshalAs(UnmanagedType.LPStruct)] object rInfo,
			[MarshalAs(UnmanagedType.LPStruct)] object rPhan,
			int dynamic)
		{
			ksRequestInfo info = (ksRequestInfo)rInfo;
			ksPhantom phan = (ksPhantom)rPhan;

			if (info != null && phan != null)
			{
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				switch (comm)
				{
					case 1:
					case 2:
						type = comm;
						break;
					case -1: // ��������� � ������
						{
							doc.ksMoveObj(t1.gr, x, y);
							// � ������� �� Cursor ���� �������� � ���� ��������� �������
							if (Math.Abs(ang) > 0.001)
								doc.ksRotateObj(t1.gr, x, y, ang);
							doc.ksStoreTmpGroup(t1.gr);
							// ��������� ��������� ������ � ���
							doc.ksClearGroup(t1.gr, true);
							break;
						}
				}

				// ������ ��� ������� ������ ���� ��������� � ����������� ��� ��������� ���� ���������
				if (t1.gr > 0)
					doc.ksDeleteObj(t1.gr);
				t1.gr = doc.ksNewGroup(1); // ��������� ������

				// ����������� �� ������ ����������� �� � ���� ��� �������
				if ((flag == 1 && comm == 1) || (flag == 2 && comm == 2))
					type = 3;

				switch (type)
				{
					case 1:
						doc.ksCircle(0, 0, 20, 1);
						info.commandsString = "!������� !����������� ";
						flag = 1;
						break;
					case 2:
						doc.ksLineSeg(-10, 0, 10, 0, 1);
						doc.ksLineSeg(10, 0, 0, 20, 1);
						doc.ksLineSeg(0, 20, -10, 0, 1);
						info.commandsString = "!���������� !������� ";
						flag = 2;
						break;
					case 3:
						doc.ksLineSeg(-10, 0, 10, 0, 1);
						doc.ksLineSeg(10, 0, 10, 20, 1);
						doc.ksLineSeg(10, 20, -10, 20, 1);
						doc.ksLineSeg(-10, 20, -10, 0, 1);
						info.commandsString = "!���������� !����������� ";
						flag = 0;
						break;
				}

				doc.ksEndGroup();
			}

			return 1;
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
