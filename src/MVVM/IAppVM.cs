using ExcelDna_MVVM.GUI;

namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    #endregion

    public interface IAppVM : IVM
    {
        WindowService WindowService { get; set; }
    }
}
