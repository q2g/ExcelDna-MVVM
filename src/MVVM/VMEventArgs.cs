namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using System;
    using System.Collections.Generic;
    #endregion

    public class VMEventArgs : EventArgs
    {
        public List<object> VMs { get; set; }
        public int HWND { get; set; }
    }
}
