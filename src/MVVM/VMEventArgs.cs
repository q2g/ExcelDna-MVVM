namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using System;
    #endregion

    public class VMEventArgs : EventArgs
    {
        public object VM { get; set; }
        public int HWND { get; set; }
    }
}
