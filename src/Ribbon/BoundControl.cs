﻿namespace ExcelDna_MVVM.Ribbon
{
    #region Usings
    using ExcelDna_MVVM.Utils;
    using System.ComponentModel;
    using System.Runtime.CompilerServices;
    #endregion

    class BoundControl
    {
        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string caller = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(caller));
        }
        #endregion

        #region IDisposable
        public void Dispose()
        {
            Binding.Dispose();
        }
        #endregion


        #region Properties & Variables
        public BindingObject Binding { get; set; }
        public BindingInfo BindingInfo { get; set; }
        public int Hwnd { get; set; }
        #endregion
    }
}
