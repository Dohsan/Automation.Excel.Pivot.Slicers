using Automation.Excel.Pivot.Slicers;
using System;

namespace Automation.Excel.PivotConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            var slicers = new SlicerFilters();
            slicers.SetInstance("pivottest.xlsx");
            //slicers.ClearSlicerFilter("Slicer_Subclass");
            //slicers.SelectSlicerItems("Slicer_Subclass", new string[] { "Boxers (421215)", "BTS (431113)" });
            var chuff = slicers.GetSlicerCacheNames();
            for (var i = 0; i < chuff.Length; i++)
            {
                Console.WriteLine(i + " " + chuff[i]);
            }
            Console.WriteLine("Done...");
            Console.ReadKey();
        }
    }
}
