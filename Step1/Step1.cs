using Kompas6API5;

using System;
using System.IO;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;

namespace Steps.NET
{
	// ����� Step1 - ����� ������� ���������� �� C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
 	public class Step1
	{
		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step1 - ����� ������� ���������� �� C#";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			KompasObject kompas = (KompasObject) kompas_;
			kompas.ksMessage("������!");
		}

    // ����� ����� ������  short  ��� int16
    public short GetProtectNumber()
    {
      return 111;
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
