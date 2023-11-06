using GoogleSheetsHelper;
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
        private static string JsonClient0 = @"C:\MyApp\Google\google_client.json";
        private static string JsonClient1 = @"C:\MyApp\Google\google_client1.json";
        private static string JsonClient2 = @"C:\MyApp\Google\google_client2.json";
        private static string JsonClient3 = @"C:\MyApp\Google\google_client3.json";
        private static string JsonClient4 = @"C:\MyApp\Google\google_client4.json";
        private static string JsonClient5 = @"C:\MyApp\Google\google_client5.json";

        private static string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";
        private static string SheetName = "My Sheet";

        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();

            //CreateSheet();
            //GetSheets();
            //TestWrite();
            //TestRead();
            //TestStressRead();
            //DeleteSheet();
            TestWrite2();
            //TestWriteByKey();
            //TestUpdater();

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
            var list = client.Get(SheetName).Result;
            foreach (var line in list)
            {
                foreach (var value in line)
                    Console.Write(value + "\t");
                Console.WriteLine();
            }
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
                        var list = client.Get(SheetName).Result;
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
                        new GoogleSheetCell(DateTime.Now) { DateTimeFormat = "dd-mm-yyyy" }
                    }
                }
            };
            
            requests.Add(request);
            var client = GetClient();
            client.Update(requests).Wait();
        }

        private static void TestWrite2()
        {
            var titles = new[] { "Заголовок1", "Заголовок2", "Заголовок3", "Заголовок4", null, "Время обновления" };
            var items = new List<object[]>
            {
                { new object[] { "Value1", 1, 1.1, (decimal)2.1, null, DateTime.Now } },
                { new object[] { "Value2", 2, 2.2, (decimal)2.2, null, DateTime.Now } },
            };

            using var client = GetClient();
            GoogleUtils.Write(client, "Write", items, 1, 1, 5, titles).Wait();
        }

        private static void TestWriteByKey()
        {
            var titles = new[] { "Заголовок1", "Заголовок2", "Заголовок3", "Заголовок4", null, "Время обновления" };
            var items = new Dictionary<string, object[]>
            {
                { "Key1", new object[] { "Value1", 1, 1.1, (decimal)2.1, null, DateTime.Now } },
                { "Key2", new object[] { "Value2", 2, 2.2, (decimal)2.2, null, DateTime.Now } },
            };
            using (var client = GetClient())
            {
                GoogleUtils.WriteByKey(client, "WriteByKey", items, 0, 1, titles).Wait();
            }
        }

        private static void TestUpdater()
        {
            var ct = new CancellationTokenSource();
            ct.Cancel();
            var client = GetClient();
            var updater = new GoogleSheetUpdater(client, "Updater", 0, 1);

            updater.Add("Key1", 1, DateTime.Now);
            updater.Add("Key2", 2, DateTime.Now);

            updater.Send(false, ct.Token).Wait();
            updater.Send(false).Wait();
        }

        private static GoogleSheetsClient GetClient(string json = null)
        {
            return new GoogleSheetsClient(json ?? JsonClient0, TableId);
        }
    }
}
