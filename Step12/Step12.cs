
using Kompas6API5;
using KompasAPI7;


using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6Constants;
using Kompas6Constants3D;

using reference = System.Int32;

namespace Steps.NET
{
	// ����� Step12 - ������ �������� ���������������� ������ �������
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step12
	{
		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step12 - ������ ������ ������� (C#.NET)";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Global.Kompas = (KompasObject)kompas_;
			Global.NewKompasAPI = (IApplication)Global.Kompas.ksGetApplication7();

			switch (command)
			{
				case 1 :	// ������� �������� � �����������
					Global.CreateAndSubscriptionPropertyManager(true);
					break;
				case 2 :	// ����������
					Global.ClosePropertyManager(true);	// ���������� ��������� ������ � ����� ��
					break;
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
					result = "������� �������� � �����������";
					command = 1;
					break;
				case 2:
					result = "����������";
					command = 2;
					break;
				case 3:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
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
