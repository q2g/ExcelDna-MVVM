using System.Reflection;
using System.Windows.Input;

namespace ExcelDna_MVVM.Ribbon
{
    class AddInInformationProxy : IAddInInformation
    {
        #region ctor
        public AddInInformationProxy(object objToWrap)
        {
            objectTowrap = objToWrap;
            //TODO: Cache this
            miGetRibbonXML = objectTowrap.GetType().GetMethod(nameof(IAddInInformation.GetRibbonXML));
            miGetResource = objectTowrap.GetType().GetMethod(nameof(IAddInInformation.GetResource));
            piInvalidateRibbonCommand = objectTowrap.GetType().GetProperty(nameof(IAddInInformation.InvalidateRibbonCommand));
        }
        #endregion

        #region Properties & variables
        object objectTowrap;
        MethodInfo miGetRibbonXML;
        MethodInfo miGetResource;
        PropertyInfo piInvalidateRibbonCommand;
        #endregion

        #region IAddInInformation
        public ICommand InvalidateRibbonCommand
        {
            get => (ICommand)piInvalidateRibbonCommand.GetValue(objectTowrap);
            set => piInvalidateRibbonCommand.SetValue(objectTowrap, value);
        }

        public object GetResource(string key)
        {
            return miGetResource.Invoke(objectTowrap, new object[] { key });
        }

        public string GetRibbonXML()
        {
            return (string)miGetRibbonXML.Invoke(objectTowrap, new object[] { });
        }


        #endregion



    }
}
