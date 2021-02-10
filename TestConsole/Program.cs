using AndreyPro.GoogleSheetsHelper;
using System;
using System.Diagnostics;
using System.Threading;

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
                var sw = Stopwatch.StartNew();
                var list = client.Get(sheetName);
                Console.WriteLine($"{DateTime.Now}, {sw.Elapsed}");
                Thread.Sleep(1000);
            }
        }
    }
}
