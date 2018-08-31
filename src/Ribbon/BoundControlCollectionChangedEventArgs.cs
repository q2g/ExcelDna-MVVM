namespace ExcelDna_MVVM.Ribbon
{
    #region Usings
    using System.Collections.Specialized;
    #endregion

    public class BoundControlCollectionChangedEventArgs : NotifyCollectionChangedEventArgs
    {

        #region CTOR
        public BoundControlCollectionChangedEventArgs(NotifyCollectionChangedAction action) : base(action)
        {

        }
        #endregion

        #region Properties & Variables
        public BoundControl Control { get; set; }
        #endregion
    }
}
