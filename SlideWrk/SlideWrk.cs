using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	// ����� SlideWrk - ���� �����
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class SlideWrk
	{
		private OpenFileDialog openFileDialog = new OpenFileDialog();
		private static double cx1 = 0, cy1 = 0, cx2 = 0, cy2 = 0;
		private static int flag = 0;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "SlideWrk - ���� �����";
		}


		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Global.Instance.KompasObject = (KompasObject)kompas_;
			if (Global.Instance.KompasObject != null)
			{
				Global.Instance.Document2D = (ksDocument2D)Global.Instance.KompasObject.ActiveDocument2D();
				if (Global.Instance.Document2D != null && Global.Instance.Document2D.reference != 0)
				{
					switch (command)
					{
						case 1 : MoveSlide();			break; // �������� �����  
						case 2 : WriteSlideStep();		break; // �������� �����
						case 3 : TestShowDialog();		break; // ���������� �����
						case 4 : DecomposeSlideStep();	break; // ������� �����
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
					result = "�������� �����";
					command = 1;
					break;
				case 2:
					result = "�������� �����";
					command = 2;
					break;
				case 3:
					result = "���������� �����";
					command = 3;
					break;
				case 4:
					result = "������� ��������";
					command = 4;
					break;
				case 5:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		private void MoveSlide() 
		{
			// ������� ��� �����
			string fileNameOld = Global.Instance.KompasObject.ksChoiceFile("*.rc", null, false);
			string fileNameNew = Global.Instance.KompasObject.ksSaveFile("*.rc", null, null, false);
			string s = string.Empty;

			if (fileNameOld != null && fileNameOld != string.Empty
				&& fileNameNew != null && fileNameNew != string.Empty)
			{
				int x = 0, y = 0;
				if (Global.Instance.KompasObject.ksReadInt("������� ����� �� X", 0, -2000, 2000, ref x) != 0 
					&& Global.Instance.KompasObject.ksReadInt("������� ����� �� Y", 0, -2000, 2000, ref y) != 0) 
				{
					SlideFile file = new SlideFile(fileNameOld);
					if (!file.Error) 
					{
						for (int i = 0; i < file.Strings.Count; i ++)
						{
							string str = file.Strings[i];
							if (file.CheckCommentary(str))
								if (file.MoveElement(str, x, y))
									str += "\r\n";
						}
					}
					file.Save(fileNameNew);
					Global.Instance.KompasObject.ksMessageBoxResult();
				}
			}
		}

		private void DecomposeSlideStep() 
		{
			ksRequestInfo info = (ksRequestInfo)Global.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			ksPhantom phan = (ksPhantom)Global.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_Phantom);
			ksIterator iter = (ksIterator)Global.Instance.KompasObject.GetIterator();
			if (info != null && phan != null && iter != null)
			{
				phan.Init();
				info.Init();
				info.commandsString = "������� ������ ����� ����";
				info.dynamic = 1;
				info.SetCallBackC("CALLBACKPROCCURSOR", 0, this);
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				if (t1 != null) 
				{
					t1.Init();
					t1.scale_ = 1;
					t1.gr = Global.Instance.Document2D.ksNewGroup(1);
					Global.Instance.Document2D.ksEndGroup();

					flag = 0;
					double x = 0, y = 0;
					Global.Instance.Document2D.ksCursor(info, ref x, ref y, phan);
			
					if (Global.Instance.Document2D.ksExistObj(t1.gr) != 0) 
						Global.Instance.Document2D.ksDeleteObj(t1.gr);

					if (flag == 2) 
					{
						Global.Instance.Document2D.ksClearGroup(0, true);
						if (Global.Instance.Document2D.ksSelectGroup(0, 1, cx1, cy1, cx2, cy2) != 0) 
						{
							if (iter.ksCreateIterator(ldefin2d.SELECT_GROUP_OBJ, 0)) 
							{
								t1.gr = Global.Instance.Document2D.ksNewGroup(1);
								Global.Instance.Document2D.ksEndGroup();
								int rObj;
								int gr;
								if (Global.Instance.Document2D.ksExistObj(rObj = iter.ksMoveIterator("F")) != 0) 
								{
									do 
									{
										gr = Global.Instance.Document2D.ksDecomposeObj(rObj, 0, 0.2F, 0);
										Global.Instance.Document2D.ksAddObjGroup(t1.gr, gr);
										Global.Instance.Document2D.ksClearGroup(gr, false);
										Global.Instance.Document2D.ksDeleteObj(gr);
									}
									while (Global.Instance.Document2D.ksExistObj(rObj = iter.ksMoveIterator("N")) != 0);
								}

								info.Init();
								t1.xBase = cx1;
								t1.yBase = cy1;
								info.commandsString = "������� ������� ����� ������";
								if (Global.Instance.Document2D.ksCursor(info, ref x, ref y, phan) != 0) 
								{
									Global.Instance.Document2D.ksMtr(x - cx1, y - cy1, 0, 1, 1);
									Global.Instance.Document2D.ksTransformObj(t1.gr);
									Global.Instance.Document2D.ksDeleteMtr();
									Global.Instance.Document2D.ksStoreTmpGroup(t1.gr);
								}

								Global.Instance.Document2D.ksClearGroup(t1.gr, true);
								Global.Instance.Document2D.ksDeleteObj(t1.gr); 
							}
							Global.Instance.Document2D.ksClearGroup(0, true);
						}
						else
							Global.Instance.KompasObject.ksError("������");
					}
				}
			}
		}

		private void WriteSlideStep()
		{
			string name = Global.Instance.KompasObject.ksSaveFile("*.rc", null, null, false);
			if (name != null && name != string.Empty)
			{
				ksRequestInfo info = (ksRequestInfo)Global.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.commandsString = "������� ����� �������� ������";

					double x = 0, y = 0;
					if (Global.Instance.Document2D.ksCursor(info, ref x, ref y, null) != 0)
					{
						int slideID = 0;
						if (Global.Instance.KompasObject.ksReadInt("������� ������������� ������", 100, 0, 32000, ref slideID) != 0)
						{
							if (Global.Instance.KompasObject.ksWriteSlide(name, slideID, x, y) == 0)
								Global.Instance.KompasObject.ksError("������ �������������� �����");
							Global.Instance.Document2D.ksClearGroup(0, true);
						}
					}
				}
			}
		}

		private void TestShowDialog()
		{
			openFileDialog.Filter = "����� ������� (*.rc)|*.rc";	
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				Global.Instance.KompasObject.ksEnableTaskAccess(0);	// ������� ������ � �������
				FrmTest.Instance.Kompas = Global.Instance.KompasObject;
				FrmTest.Instance.SlideFileName = openFileDialog.FileName;
				FrmTest.Instance.ShowDialog();
				Global.Instance.KompasObject.ksEnableTaskAccess(1);	// ������� ������ � �������
			}
		}


		/// <summary>
		/// ������� �������� �����, ���������� �� Cursor
		/// </summary>
		public int CALLBACKPROCCURSOR(int comm, ref double x, ref double y, object rInfo, object rPhan, int dynamic)
		{
			ksRequestInfo info = (ksRequestInfo)rInfo;
			ksPhantom phan = (ksPhantom)rPhan;
			if (info != null && phan != null)
			{
				phan.phantom = 1;
				ksType1 t1 = (ksType1)phan.GetPhantomParam();
				if (dynamic == 0) 
				{
					//��������
					if (flag == 0) 
					{
						//����������� 1 �����
						flag ++;
						cx1 = x; cy1= y;
						info.commandsString = "������� ������ ����� ����";
					}
					else 
					{
						flag ++;
						cx2 = x; cy2 = y;
						return 0;
					}
				}
				else 
				{
					if ( flag > 0 ) 
					{
						cx2 = x; cy2 = y;
						t1.xBase = cx2; 
						t1.yBase = cy2;
				
						if (Global.Instance.Document2D.ksExistObj(t1.gr) != 0)
							Global.Instance.Document2D.ksDeleteObj(t1.gr);

						t1.gr = Global.Instance.Document2D.ksNewGroup(1);	// ��������� ������
						DrawRamka();
						Global.Instance.Document2D.ksEndGroup();
					}
				}
			}
			return 1;
		}

		static void DrawRamka()
		{
			if (Math.Abs(cy1 - cy2) > 0) 
			{
				Global.Instance.Document2D.ksLineSeg(cx1, cy1, cx1, cy2, 1);
				Global.Instance.Document2D.ksLineSeg(cx2, cy1, cx2, cy2, 1);
			}
			if(Math.Abs(cx1 - cx2) > 0) 
			{
				Global.Instance.Document2D.ksLineSeg(cx1, cy1, cx2, cy1, 1);
				Global.Instance.Document2D.ksLineSeg(cx1, cy2, cx2, cy2, 1);
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
