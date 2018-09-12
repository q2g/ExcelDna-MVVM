namespace ExcelDna_MVVM.Environment
{

    #region Usings
    using ExcelDna.Integration;
    using NLog;
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    #endregion

    class EnvironmentAdapter
    {

        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion


        public static Task QueueAction(Action action)
        {
            AutoResetEvent resetEvent = new AutoResetEvent(false);
            Task waitingTask = Task.Run(() =>
            {
                resetEvent.WaitOne();
            });
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
                resetEvent.Set();
            });
            return waitingTask;
        }
    }
}
