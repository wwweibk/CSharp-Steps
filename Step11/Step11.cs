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
	// ����� Step11 - ������ ������� ���������� ����
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step11
	{
		private KompasObject kompas;

		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public object GetLibraryName()
		{
			return "Step11 - ������ ������� ���������� ����";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				ksDocument2D doc = (ksDocument2D)kompas.Document2D();
				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (doc != null && info != null) 
				{
					info.Init();
					info.menuId = 3002;
					info.title = "������ ������";

					info.SetCallBackCm("CALLBACKPROCCOMMANDWINDOW", 0, this);
					int cmd = doc.ksCommandWindow(info);
					kompas.ksMessage(string.Format("������� ������� {0}", cmd));
				}
			}
		}


		// ������������ ���� ����������
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem1(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;	//�� ��������� - ������ ������
			itemType = 1;					//MENUITEM

			switch (number)
			{
				case 1:
					result = "������ ������� ���������� ����";
					command = 1;
					break;
				case 2:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}
	

		// HINSTANCE ������ � ���������
		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
		}


		// ������� �������� �����
		// 0 - ���������, 1 - �������, 2 - ��� - ������ collection
		public int CALLBACKPROCCOMMANDWINDOW(int comm, [In][MarshalAs(UnmanagedType.LPStruct)]object rInfo)
		{
			kompas.ksMessage(string.Format("����������� ������� {0}", comm));
			// �������� ��������� ���� � ����������� �� ����������� �������
			// ���������� ����� �������� � ������ ������ ������
			ksRequestInfo info = null;
			info = (ksRequestInfo)rInfo;
			if (info != null) 
			{
				switch (comm) 
				{
					case 1	: info.title = "1"; break;
					default	: info.title = "2"; break;
				}
			}
			// ������������ ��������� ����������, ������ �� ������� ����������
			// ����������� ������� :
			// TRUE - ����������
			// FALSE - ��������� ������ � �����
			return comm == 2213 ? 0 : 1;
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
