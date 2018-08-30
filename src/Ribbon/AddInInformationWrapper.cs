using System.Reflection;
using System.Windows.Input;

namespace ExcelDna_MVVM.Ribbon
{
    class AddInInformationWrapper : IAddInInformation
    {
        #region ctor
        public AddInInformationWrapper(object objToWrap)
        {
            objectTowrap = objToWrap;
            miGetRibbonXML = objectTowrap.GetType().GetMethod(nameof(IAddInInformation.GetRibbonXML));
            miGetLocalizedString = objectTowrap.GetType().GetMethod(nameof(IAddInInformation.GetLocalizedString));
            piInvalidateRibbonCommand = objectTowrap.GetType().GetProperty(nameof(IAddInInformation.InvalidateRibbonCommand));
        }
        #endregion

        #region Properties & variables
        object objectTowrap;
        MethodInfo miGetRibbonXML;
        MethodInfo miGetLocalizedString;
        PropertyInfo piInvalidateRibbonCommand;
        #endregion

        #region IAddInInformation
        public ICommand InvalidateRibbonCommand
        {
            get => (ICommand)piInvalidateRibbonCommand.GetValue(objectTowrap);
            set => piInvalidateRibbonCommand.SetValue(objectTowrap, value);
        }

        public string GetLocalizedString(string key)
        {
            return (string)miGetLocalizedString.Invoke(objectTowrap, new object[] { key });
        }

        public string GetRibbonXML()
        {
            return (string)miGetRibbonXML.Invoke(objectTowrap, new object[] { });
        }
        #endregion



    }
}
