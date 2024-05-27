using System;
using System.Reflection;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string wordPluginPath = "E:\\2024@II\\Net\\Practice3\\WordReportPlugin\\bin\\Debug\\WordReportPlugin.dll";
        string excelPluginPath = "E:\\2024@II\\Net\\Practice3\\ExcelReportPlugin\\bin\\Debug\\ExcelReportPlugin.dll";

        Console.WriteLine("Select the export module (1 - Word, 2 - Excel): ");
        string choice = Console.ReadLine();

        if (choice == "1")
        {
            Assembly wordAssembly = Assembly.LoadFrom(wordPluginPath);
            dynamic wordPlugin = Activator.CreateInstance(wordAssembly.GetType("WordReportPlugin.WordReport"));

            Console.WriteLine("Select the type of Word report (1 - Text, 2 - Table): ");
            string wordChoice = Console.ReadLine();

            if (wordChoice == "1")
            {
                wordPlugin.GenerateReport("This is a sample text for Word report.");
            }
            else if (wordChoice == "2")
            {
                string[,] data = new string[,]
                {
                    { "Column1", "Column2" },
                    { "Row1 Col1", "Row1 Col2" },
                    { "Row2 Col1", "Row2 Col2" }
                };
                wordPlugin.GenerateReportWithTable(data);
            }
            else
            {
                Console.WriteLine("Invalid choice.");
            }
        }
        else if (choice == "2")
        {
            Assembly excelAssembly = Assembly.LoadFrom(excelPluginPath);
            dynamic excelPlugin = Activator.CreateInstance(excelAssembly.GetType("ExcelReportPlugin.ExcelReport"));

            string[] data = new string[] { "Row1", "Row2", "Row3" };
            excelPlugin.GenerateReport(data);
        }
        else
        {
            Console.WriteLine("Invalid choice.");
        }
    }
}