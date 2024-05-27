using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordReportPlugin
{
    public class WordReport
    {
        public void GenerateReport(string text)
        {
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();
            Word.Paragraph p = doc.Paragraphs.Add();
            p.Range.Text = text;

            string filePath = @"E:\2024@II\Net\Practice3\Reports\WordReport.docx";
            doc.SaveAs2(filePath);
            doc.Close();
            wdApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(p);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("Word report generated successfully.");
            Console.ReadLine();
        }

        public void GenerateReportWithTable(string[,] data)
        {
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();

            Word.Table table = doc.Tables.Add(doc.Range(0, 0), data.GetLength(0), data.GetLength(1));
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    table.Cell(i + 1, j + 1).Range.Text = data[i, j];
                }
            }

            string filePath = @"E:\2024@II\Net\Practice3\Reports\WordReport.docx";
            doc.SaveAs2(filePath);
            doc.Close();
            wdApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("Word report with table generated successfully.");
            Console.ReadLine();
        }
    }
}