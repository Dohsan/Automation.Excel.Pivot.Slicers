using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Automation.Excel.Pivot.Slicers
{
    /// <summary>
    /// The main slicer flitering class
    /// Contains all methods for performing filtering on excel slicers
    /// </summary>
    public class SlicerFilters
    {
        /// <summary>
        /// Instance of Excel
        /// </summary>
        public static Application xlApp;
        /// <summary>
        /// Reference to existing Excel Workbook
        /// </summary>
        public static Workbook xlWorkBook;

        const int DISP_E_BADINDEX = unchecked((int)0x8002000B);
        const int E_INVALIDARG = unchecked((int)0x80070057);

        /// <summary>
        /// Establishes instance of open excel application
        /// </summary>
        /// <param name="targetWorkbook">Target Excel workbook to establish instance on</param>
        /// <exception cref="ArgumentException">Thrown when <paramref name="targetWorkbook"/> is null or empty</exception>
        /// <exception cref="Exception">Thrown when <paramref name="targetWorkbook"/> does not exist in the Excel workbook collection</exception>
        public void SetInstance(string targetWorkbook)
        {
            if (string.IsNullOrEmpty(targetWorkbook)) throw new ArgumentException(nameof(targetWorkbook) + " cannot be null or empty.");

            xlApp = (Application)Marshal.GetActiveObject("Excel.Application");
            Workbooks workbooks = xlApp.Workbooks;

            try
            {
                xlWorkBook = workbooks[targetWorkbook];
            }
            catch (COMException ce) when (ce.ErrorCode == DISP_E_BADINDEX) //Catch Com errors only to check
            {
                throw new Exception("The target Excel workbook does not exist", ce);
            }
        }

        /// <summary>
        /// Provides list of slicer caches from target workbook
        /// </summary>
        /// <returns>Array of slicer cache names</returns>
        public string[] GetSlicerCacheNames()
        {
            if (xlWorkBook == null) throw new NullReferenceException(nameof(xlWorkBook) + " Has not been set and cannot be null");

            var slicerCount = xlWorkBook.SlicerCaches.Count;
            var silcerCaches = new string[slicerCount];
            int i = 0;
            foreach (SlicerCache cache in xlWorkBook.SlicerCaches)
            {
                silcerCaches[i] = cache.Name;
                i++;
            }

            return silcerCaches;
        }

        /// <summary>
        /// Clears all filters on a given slicer cache
        /// </summary>
        /// <param name="slicerCache">name of the slicer cache to clear filters on</param>
        public void ClearSlicerFilter(string slicerCache)
        {
            if (xlWorkBook == null) throw new NullReferenceException(nameof(xlWorkBook) + " Has not been set and cannot be null");
            if (string.IsNullOrEmpty(slicerCache)) throw new ArgumentException(nameof(slicerCache) + " cannot be null or empty.");

            try
            {
                SlicerCache sc = xlWorkBook.SlicerCaches[slicerCache];
                if (!sc.FilterCleared)
                    sc.ClearAllFilters();
            }
            catch (ArgumentException ce) when (ce.HResult == E_INVALIDARG) //Catch Com errors only to check
            {
                throw new Exception("The " + nameof(slicerCache) + " does not exist in the target workbook", ce);
            }
        }

        /// <summary>
        /// Selects provided slicer items on slicer
        /// </summary>
        /// <param name="slicerCache">Cache object of the slicer</param>
        /// <param name="slicerSelections">items to select on the slicer</param>
        public void SelectSlicerItems(string slicerCache, string[] slicerSelections)
        {
            if (xlWorkBook == null) throw new NullReferenceException(nameof(xlWorkBook) + " Has not been set and cannot be null");
            if (string.IsNullOrEmpty(slicerCache)) throw new ArgumentException(nameof(slicerCache) + " cannot be null or empty.");

            try
            {
                SlicerCache sc = xlWorkBook.SlicerCaches[slicerCache];
                var slicerItemsToSelect = slicerSelections.ToList();
                var isOlap = sc.OLAP;

                if (isOlap)
                    SelectSlicerItemsOlap(sc, slicerItemsToSelect);
                else
                    SelectSlicerItemsNonOlap(sc, slicerItemsToSelect);
            }
            catch (ArgumentException ce) when (ce.HResult == E_INVALIDARG) //Catch Com errors only to check
            {
                throw new Exception("The " + nameof(slicerCache) + " does not exist in the target workbook", ce);
            }
        }

        private void SelectSlicerItemsOlap(SlicerCache sc, List<string> slicerItemsToSelect)
        {
            List<string> slicerItemsFound = new List<string>();

            SlicerCacheLevel sl = sc.SlicerCacheLevels[1];

            foreach (SlicerItem slicerItem in sl.SlicerItems)
            {
                if (slicerItemsToSelect.Exists(item => item == slicerItem.Value))
                {
                    slicerItemsToSelect.Remove(slicerItem.Value); //once found remove item from selection list
                    slicerItemsFound.Add(slicerItem.Name); //add found value to new collection
                    if (!slicerItemsToSelect.Any()) break; //If all selection items found exit loop
                }
            }

            sc.VisibleSlicerItemsList = slicerItemsFound.ToArray(); //Assign selections
        }

        private void SelectSlicerItemsNonOlap(SlicerCache sc, List<string> slicerItemsToSelect)
        {
            foreach (SlicerItem slicerItem in sc.SlicerItems)
            {
                if (slicerItemsToSelect.Exists(item => item == slicerItem.Value))
                    slicerItem.Selected = true;
                else
                    slicerItem.Selected = false;
            }
        }
    }
}
