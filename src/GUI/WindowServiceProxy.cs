using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ExcelDna_MVVM.GUI
{
    class WindowServiceProxy : WindowService
    {
        #region LoggerInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region ctor
        public WindowServiceProxy(object objToWrap)
        {
            objectToWrap = objToWrap;
            var objectToWrapType = objToWrap.GetType();
            //TODO Cache this
            piGetHwnd = objectToWrapType.GetProperty(nameof(WindowService.GetHwnd));
            piRibbonHeight = objectToWrapType.GetProperty(nameof(WindowService.RibbonHeight));
            piRibbonWidth = objectToWrapType.GetProperty(nameof(WindowService.RibbonWidth));
            miShowOverlay = objectToWrapType.GetMethod(nameof(WindowService.ShowOverlay));
        }
        #endregion

        #region Properties & Variables
        private object objectToWrap;
        PropertyInfo piGetHwnd;
        PropertyInfo piRibbonHeight;
        PropertyInfo piRibbonWidth;
        MethodInfo miShowOverlay;
        public object Target { get => objectToWrap; }
        #endregion

        #region overrides
        public override Func<int> GetHwnd
        {
            get
            {
                return (Func<int>)piGetHwnd.GetValue(objectToWrap);
            }
            set
            {
                piGetHwnd.SetValue(objectToWrap, value);
            }
        }
        public override double RibbonHeight
        {
            get
            {
                return (double)piRibbonHeight.GetValue(objectToWrap);
            }
            set
            {
                piRibbonHeight.SetValue(objectToWrap, value);
            }
        }
        public override double RibbonWidth
        {
            get
            {
                return (double)piRibbonWidth.GetValue(objectToWrap);
            }
            set
            {
                piRibbonWidth.SetValue(objectToWrap, value);
            }
        }
        public override void ShowOverlay(UserControl content, double verticalOffset = 0, double horizontalOffset = 0)
        {
            miShowOverlay.Invoke(objectToWrap, new object[] { content, verticalOffset, horizontalOffset });
        }

        #endregion
    }
}
