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
            Task waitingTask = new Task(() => { resetEvent.WaitOne(); });
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    action();
                    resetEvent.Set();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            });
            return waitingTask;
        }

        //public static Task<T> QueueAction<T>(Action<T> action, T parameter)
        //{
        //    AutoResetEvent resetEvent = new AutoResetEvent(false);
        //    Task waitingTask = new Task(() => { resetEvent.WaitOne(); });
        //    ExcelAsyncUtil.QueueAsMacro(() =>
        //    {
        //        try
        //        {
        //            action(parameter);
        //            resetEvent.Set();
        //        }
        //        catch (Exception ex)
        //        {
        //            logger.Error(ex);
        //        }
        //    });
        //    return waitingTask;
        //}
    }
}
