using AndreyPro.GoogleSheetsHelper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace TestConsole
{
    class Program
    {
        private static string JsonClient0 = @"C:\MyApp\google_client.json";
        private static string JsonClient1 = @"C:\MyApp\google_client1.json";
        private static string JsonClient2 = @"C:\MyApp\google_client2.json";
        private static string JsonClient3 = @"C:\MyApp\google_client3.json";
        private static string JsonClient4 = @"C:\MyApp\google_client4.json";
        private static string JsonClient5 = @"C:\MyApp\google_client5.json";

        private static string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";
        private static string SheetName = "My Sheet";

        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();

            //CreateSheet();
            //GetSheets();
            //TestWrite();
            //TestRead();
            TestStressRead();
            //DeleteSheet();
            //TestWriteByKey();

            Console.WriteLine($"Elapsed: {sw.Elapsed}");
            Console.ReadKey();
        }

        static void CreateSheet()
        {
            var client = GetClient();
            client.AddSheetIfNotExist(SheetName).Wait();
        }

        static void GetSheets()
        {
            var client = GetClient();
            var sheets = client.GetSheets().Result;
            sheets.ToList().ForEach(x => Console.WriteLine($"Sheet: {x}"));
        }

        static void DeleteSheet()
        {
            var client = GetClient();
            client.DeleteSheet(SheetName).Wait();
        }

        static void TestRead()
        {
            var client = GetClient();
            var list = client.GetAsync(SheetName).Result;
        }

        static void TestStressRead()
        {
            int num = 0;
            void Run(string json)
            {
                try
                {
                    for (int i = 0; i < 10000; i++)
                    {
                        var client = GetClient(json);
                        var list = client.GetAsync(SheetName).Result;
                        var n = Interlocked.Increment(ref num);
                        Console.WriteLine($"{n}, {i}: {json}");
                        Thread.Sleep(1000);
                    }
                }
                catch (AggregateException ex)
                {
                    var m = ex.InnerException.ToString();
                    Console.WriteLine($"{m.Substring(0, Math.Min(m.Length, 50))}: {json}");
                }
                catch (Exception ex)
                {
                    var m = ex.ToString();
                    Console.WriteLine($"{m.Substring(0, Math.Min(m.Length, 50))}: {json}");
                }
            }

            Task.Run(() => Run(JsonClient1));
            Task.Run(() => Run(JsonClient2));
            Task.Run(() => Run(JsonClient3));
            Task.Run(() => Run(JsonClient4));
            Task.Run(() => Run(JsonClient5));

            Console.ReadKey();
        }

        static void TestWrite()
        {
            var columnStart = 0;
            var rowStart = 0;
            var requests = new List<GoogleSheetUpdateRequest>();
            var request = new GoogleSheetUpdateRequest(SheetName)
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
                        new GoogleSheetCell(DateTime.Now) { NumberFormat = "dd/mm/yy hh:mm:ss" }
                    }
                }
            };
            
            requests.Add(request);
            var client = GetClient();
            client.Update(requests).Wait();
        }

        private static void TestWriteByKey()
        {
            var items = new Dictionary<string, object[]>
            {
                { "Key1", new object[] { "Value1", 1, 1.1, null, DateTime.Now } },
                { "Key2", new object[] { "Value2", 2, 2.2, null, DateTime.Now } },
            };
            var client = GetClient();
            GoogleUtils.WriteByKey(client, "WriteByKey", 0, 1, items).Wait();
        }

        private static GoogleSheetsClient GetClient(string json = null)
        {
            return new GoogleSheetsClient(json ?? JsonClient0, TableId);
        }
    }
}
