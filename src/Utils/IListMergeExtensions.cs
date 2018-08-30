using System;
using System.Collections.Generic;

namespace ExcelDna_MVVM.Utils
{

    public static class IListMergeExtensions
    {
        public static bool Diff<T>(this IList<T> list, IEnumerable<T> listToCompare, out List<T> newItems, out List<T> removedItems)
        {
            removedItems = new List<T>();
            foreach (var item in list)
                removedItems.Add(item);

            newItems = new List<T>();

            foreach (var item in listToCompare)
            {
                if (list.Contains(item))
                    removedItems.Remove(item);
                else
                    newItems.Add(item);
            }
            return true;
        }

        public static bool MergeChanges<T>(this IList<T> MergeList, IEnumerable<T> newList, Comparison<T> comparison = null)
        {
            MergeList.Diff(newList, out var AddList, out var RemoveList);


            foreach (var item in RemoveList)
                MergeList.Remove(item);

            foreach (var item in AddList)
            {
                bool added = false;
                if (comparison != null)
                {
                    for (int i = 0; i < MergeList.Count; i++)
                    {
                        if (comparison.Invoke(item, MergeList[i]) >= 0)
                        {
                            MergeList.Insert(i, item);
                            added = true;
                            break;
                        }
                    }
                }
                if (!added) MergeList.Insert(MergeList.Count, item);
            }

            return RemoveList.Count > 0 | AddList.Count > 0;
        }

    }
}
