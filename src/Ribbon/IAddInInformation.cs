using System.Windows.Input;

namespace ExcelDna_MVVM.Ribbon
{
    public interface IAddInInformation
    {
        string GetRibbonXML();
        string GetLocalizedString(string key);
        ICommand InvalidateRibbonCommand { get; set; }
    }
}
