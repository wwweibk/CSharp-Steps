using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using KAPITypes;
using System.Resources;
using Kompas6API5;

namespace Step1 {
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
 
}
