using AndreyPro.GoogleSheetsHelper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace TestConsole
{
    class Program
    {
        private static string PathToJsonFile = @"D:\Data\Projects\client_secret.json";
        private static string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";

        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();

            //TestRead();
            TestWrite();

            Console.WriteLine($"Elapsed: {sw.Elapsed}");
            Console.ReadKey();
        }

        static void TestRead()
        {
            var client = new GoogleSheetsClient(PathToJsonFile, TableId);
            var list = client.GetAsync("TestRead", new CancellationTokenSource(5000).Token).Result;
        }

        static void TestWrite()
        {
            var sheetName = "TestWrite";
            var columnStart = 0;
            var rowStart = 0;
            var request = new List<GoogleSheetUpdateRequest>();
            var r = new GoogleSheetUpdateRequest(sheetName)
            {
                ColumnStart = columnStart,
                RowStart = rowStart,
                Rows = new List<GoogleSheetRow>
                {
                    new GoogleSheetRow
                    {
                        new GoogleSheetCell("Название") { Bold = true },
                        new GoogleSheetCell(1),
                        new GoogleSheetCell(1.1),
                        new GoogleSheetCell(true),
                        new GoogleSheetCell(DateTime.Now),
                        new GoogleSheetCell(DateTime.Now) { NumberPattern = "dd/mm/yy hh:mm:ss" }
                    }
                }
            };
            
            request.Add(r);
            var client = new GoogleSheetsClient(PathToJsonFile, TableId);
            client.Update(request).Wait();
        }
    }
}
