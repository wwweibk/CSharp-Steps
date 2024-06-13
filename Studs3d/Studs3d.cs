
using Kompas6API5;
using KompasAPI7;


using System;
using System.Resources;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;

namespace Steps.NET
{
	// ����� Studs3d - ��������� ������� �� C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Studs3d
	{
		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Studs3d - ��������� ������� �� C#";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Kompas.Instance.KompasObject = (KompasObject)kompas_;
			Kompas.Instance.Document3D = (ksDocument3D)Kompas.Instance.KompasObject.ActiveDocument3D();
			if (Kompas.Instance.KompasObject != null 
				&& Kompas.Instance.Document3D != null 
				&& !Kompas.Instance.Document3D.IsDetail())
			{
				Pin par = new Pin();
				if (par != null)
					par.Draw3D(Kompas.Instance.Document3D);
				else
				{
					ResourceManager rm = new ResourceManager(this.GetType());
					Kompas.Instance.KompasObject.ksError(rm.GetString("IDS_3DDOCERROR"));
				}	
				if (Kompas.Instance.KompasObject.ksReturnResult() == (int)ErrorType.etError10)	//  10  "������! ����������� ������"
					Kompas.Instance.KompasObject.ksResultNULL();
			}
		}


		// ������������ ���� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "�������";
					command = 1;
					break;
				case 2:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// HMODULE ���������� ������
		public IntPtr ExternalGetResourceModule()
		{
			return Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetLoadedModules()[0]);;
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
