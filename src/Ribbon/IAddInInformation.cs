using System.Windows.Input;

namespace ExcelDna_MVVM.Ribbon
{
    public interface IAddInInformation
    {
        string GetRibbonXML();
        object GetResource(string key);
        ICommand InvalidateRibbonCommand { get; set; }
    }
}
