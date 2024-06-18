using Kompas6API5;


using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using KAPITypes;
using Kompas6Constants;

namespace Step1
{

    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Step1
    {

        private KompasObject kompas;
        private ksDocument2D doc;
        private ksMathematic2D mat;
        private void DrawPointByArray(ksDynamicArray arr)
        {
            if (arr != null)
            {
                // Создать интерфейс параметров математической точки
                ksMathPointParam par = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);

                if (par != null)
                {
                    // Интерфейс создан
                    for (int i = 0; i < arr.ksGetArrayCount(); i++)
                    {
                        arr.ksGetArrayItem(i, par);
                        doc.ksPoint(par.x, par.y, 5);
                        string buf = string.Format("x = {0:.##} y = {1:.##}", par.x, par.y);
                        kompas.ksMessage(buf);
                    }
                }
            }
        }
        private void IntersectEx()
        {
            ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
            if (arr != null)
            {
                doc.ksBezier(0, 0);
                doc.ksPoint(20, 0, 0);
                doc.ksPoint(10, 20, 0);
                doc.ksPoint(20, 40, 0);
                doc.ksPoint(30, 20, 0);
                doc.ksPoint(20, 0, 0);
                int pp1 = doc.ksEndObj();

                doc.ksBezier(0, 0);
                doc.ksPoint(0, 20, 0);
                doc.ksPoint(20, 10, 0);
                doc.ksPoint(40, 20, 0);
                doc.ksPoint(20, 30, 0);
                doc.ksPoint(0, 20, 0);
                int pp2 = doc.ksEndObj();

                mat.ksIntersectCurvCurv(pp1, pp2, arr);
                DrawPointByArray(arr);
                arr.ksDeleteArray();
            }
        }
        // --------------------------------------------------------------------------------
        // Пересечение двух кривых.
        // Возвращает:
        // 1 успешное завершение
        // 0 кривые не пересекаются или совпадают
        // -1 первый объект не существует
        // -2 второй объект не существует
        // -3 кривые расположены в разных видах
        // -4 не совпадают СК определения кривых 
        // -5 первый объект не является кривой
        // -6 второй объект не является кривой
        // -7 ошибка
        // динамический массив точек пересечения array должен быть создан до передачи его в функцию.
        // ---------------------------------------------------------------------------------

        [DllImport("KompasLibrary.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int ksIntersectCurvCurv(IntPtr p1, IntPtr p2, IntPtr pArr);
        // --------------------------------------------------------------------------------
        // Пересечение двух кривых.
        // ---------------------------------------------------------------------------------
        [DllImport("KompasLibrary.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern void IntersectCurvCurvEx(IntPtr p1, IntPtr p2, ref int kp, IntPtr xp, IntPtr yp, int maxCount, int touchInclude);


        // --------------------------------------------------------------------------------
        // Пересечение двух кривых.
        // Возвращает:
        // 1 успешное завершение
        // 0 кривые не пересекаются или совпадают
        // -1 первый объект не существует
        // -2 второй объект не существует
        // -3 кривые расположены в разных видах
        // -4 не совпадают СК определения кривых (геом и анн) (?)
        // -5 первый объект не является кривой
        // -6 второй объект не является кривой
        // -7 ошибка
        // ---------------------------------------------------------------------------------
        [DllImport("KompasLibrary.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int ksIntersectCurvCurvEx(IntPtr p1, IntPtr p2, IntPtr pArr, int touchInclude);




        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "Step2 - Использованиe математических функций";
        }


        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty;
            itemType = 1; // "MENUITEM"
            switch (number)
            {
                case 1:
                    result = "Пересечь кривые";
                    command = 2;
                    break;
                case 2:
                    command = -1;
                    itemType = 3; // "ENDMENU"
                    break;
            }
            return result;
        }


        public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {
            kompas = (KompasObject)kompas_;

            if (kompas == null)
                return;

            doc = (ksDocument2D)kompas.ActiveDocument2D();

            if (doc == null)
                return;

            mat = (ksMathematic2D)kompas.GetMathematic2D();

            if (mat == null)
                return;

            switch (command)
            {
                case 1: IntersectEx(); break; // пересечь кривые
            }

            kompas.ksMessageBoxResult();
        }


        public object ExternalGetResourceModule()
        {
            return Assembly.GetExecutingAssembly().Location;
        }


        public int ExternalGetToolBarId(short barType, short index)
        {
            int result = 0;

            if (barType == 0)
            {
                result = -1;
            }
            else
            {
                switch (index)
                {
                    case 1:
                        result = 3001;
                        break;
                    case 2:
                        result = -1;
                        break;
                }
            }

            return result;
        }



        #region COM Registration
        // Эта функция выполняется при регистрации класса для COM
        // Она добавляет в ветку реестра компонента раздел Kompas_Library,
        // который сигнализирует о том, что класс является приложением Компас,
        // а также заменяет имя InprocServer32 на полное, с указанием пути.
        // Все это делается для того, чтобы иметь возможность подключить
        // библиотеку на вкладке ActiveX.
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
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }

        // Эта функция удаляет раздел Kompas_Library из реестра
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

