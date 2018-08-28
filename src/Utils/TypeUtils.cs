using NLog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDna_MVVM.Utils
{
    public class TypeUtils
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region public Functions
        public static List<Type> GetTypesImplementingInterface<T>()
        {
            List<Type> retval = new List<Type>();
            try
            {
                var assemmblys = AppDomain.CurrentDomain.GetAssemblies(); //TODO: This works only in Packed Mode
                foreach (var assembly in assemmblys)
                {
                    var vmTypes = assembly.GetTypes().Where(typ => typ.GetInterfaces().Any(inter => inter.FullName == typeof(T).FullName)).ToList();//TODO:comparing Types is working only when LoadFromBytes is set to False 
                    retval.AddRange(vmTypes);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return retval;
        }
        #endregion
    }
}
