namespace ExcelDna_MVVM.GUI
{
    #region Usings
    using NLog;
    using System;
    using System.Runtime.InteropServices;
    using System.Windows;
    using System.Windows.Forms;
    #endregion

    #region ucwfSelection
    [ComVisible(true)]
    public partial class ucwfWPFContainer : UserControl
    {
        #region Variables & Properties
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public UIElement Child
        {
            get
            {
                return WPFContainer?.Child as UIElement;
            }
            set
            {
                try
                {
                    WPFContainer.Child = value;
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            }
        }
        #endregion

        #region Constructor & Load
        public ucwfWPFContainer()
        {
            InitializeComponent();
        }
        #endregion
    }
    #endregion
}
