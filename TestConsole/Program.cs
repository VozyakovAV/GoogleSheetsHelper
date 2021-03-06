using AndreyPro.GoogleSheetsHelper;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var pathToJsonFile = @"D:\Data\Projects\client_secret.json";
            var tableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";
            var sheetName = "Test1";

            var client = new GoogleSheetsClient(pathToJsonFile, tableId);
            while (true)
            {
                var t = new CancellationTokenSource(5000);
                var sw = Stopwatch.StartNew();
                var list = client.GetAsync(sheetName, t.Token).Result;
                
                Console.WriteLine($"{DateTime.Now}, {sw.Elapsed}");
                Thread.Sleep(1000);
            }
        }
    }
}
