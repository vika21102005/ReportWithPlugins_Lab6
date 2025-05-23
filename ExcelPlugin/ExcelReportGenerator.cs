using System;
using System.Runtime.InteropServices;
using PluginBase;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPlugin
{
    public class ExcelReportGenerator : IReportGenerator
    {
        public string Name => "Excel";

        public void GenerateReport(string[] items, string outputPath)
        {
            var app = new Excel.Application();
            var wb = app.Workbooks.Add();
            var ws = (Excel.Worksheet)wb.Worksheets[1];
            ws.Cells[1, 1] = "Звіт у Excel";
            for (int i = 0; i < items.Length; i++)
                ws.Cells[i + 2, 1] = items[i];

            wb.SaveAs(outputPath);
            wb.Close();
            app.Quit();

            // COM cleanup
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
