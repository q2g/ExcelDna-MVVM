namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using System;
    #endregion

    public class VMEventArgs : EventArgs
    {
        public IVM VM { get; set; }
        public int HWND { get; set; }
    }
}
