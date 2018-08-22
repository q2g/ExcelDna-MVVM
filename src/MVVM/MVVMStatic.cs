using System;
using System.Collections.Generic;

namespace ExcelDna_MVVM.MVVM
{
    public class MVVMStatic
    {
        public static MVVMAdapter Adapter { get; set; }
        public static string RibbonXml { get; set; }


        public static List<Type> RegisteredTypes { get; set; } = new List<Type>();
        public static void RegisterType(Type typ)
        {
            RegisteredTypes.Add(typ);
        }


    }
}
