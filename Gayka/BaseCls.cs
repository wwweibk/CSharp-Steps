
using Kompas6API5;
using KompasAPI7;

using System;
using System.IO;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Steps.NET
{
	// ����� Gayka - ��������������� �������
	public class Gayka
	{
		// ��� ����������
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Gayka - ��������������� �������";
		}
		

		// �������� ������� ����������
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Kompas.Instance.KompasObject = (KompasObject)kompas_;
			Kompas.Instance.Document2D = (ksDocument2D)Kompas.Instance.KompasObject.ActiveDocument2D();
			Kompas.Instance.KompasApp = (_Application)Kompas.Instance.KompasObject.ksGetApplication7();

			switch (command)
			{
				case 1:
					GaykaObj gayka = new GaykaObj();
					gayka.Draw();
					break;
			}
		}

    //-------------------------------------------------------------------------------
    // �������� �������������� ���������������� � ���������� �������
    // ---
    public bool FillContextPanel([In, MarshalAs(UnmanagedType.IDispatch)] object _contextPanel)  // ������ ������
    {
      bool res = false;
			IContextPanel contextPanel = (IContextPanel)_contextPanel;
      if (contextPanel != null)
      {
        contextPanel.Fill("GAYKA5915_CSharp_BAR");
        res = true;
      }
      return res;
    }



    // ������������ ���� ����������
    [return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "����� ���� 5915-70";
					command = 1;
					break;
				case 2:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// HINSTANCE ������ � ���������
		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
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
