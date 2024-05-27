using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportPlugin
{
    public class ExcelReport
    {
        public void GenerateReport(string[] data)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

            for (int i = 0; i < data.Length; i++)
            {
                xlSheet.Cells[i + 1, 1] = data[i];
            }

            string filePath = @"E:\2024@II\Net\Practice3\Reports\ExcelReport.xlsx";
            xlBook.SaveAs(filePath);
            xlBook.Close();
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("Excel report generated successfully.");
            Console.ReadLine();
        }
    }
}