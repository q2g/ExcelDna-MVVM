namespace ExcelDna_MVVM.Utils
{
    #region Usings
    using NLog;
    using System;
    using System.IO;
    using System.Reflection;
    #endregion

    class ResourceUtils
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion
        public static string GetResourceFileContent(string resourceFileName)
        {
            string result = null;
            try
            {
                using (Stream stream = Assembly.GetCallingAssembly()
                           .GetManifestResourceStream(resourceFileName))
                {
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        result = sr.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return result;
        }
    }
}
