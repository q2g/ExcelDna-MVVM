namespace ExcelDna_MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna.Integration.CustomUI;
    using NLog;
    using System;
    #endregion

    public static class ExcelExtensions
    {
        #region LoggingInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        private static double? excelVersion = null;
        private static double ExcelVersion
        {
            get
            {
                if (excelVersion == null)
                {
                    excelVersion = ExcelDnaUtil.ExcelVersion;
                }
                return excelVersion.Value;
            }
        }

        public static int? GetHwnd(this IRibbonControl @this)
        {
            try
            {
                if (ExcelVersion >= 15.0)
                    return (@this.Context as dynamic)?.Hwnd;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"ExcelVersion: {ExcelVersion}");
            }
            return null;
        }
    }
}
