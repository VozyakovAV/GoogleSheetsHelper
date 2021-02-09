using AndreyPro.GoogleSheetsHelper;

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
            var list = client.Get(sheetName);
        }
    }
}
