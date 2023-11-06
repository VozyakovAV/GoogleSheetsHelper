using Google.Apis.Util;
using GoogleSheetsHelper;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        private const string JsonClient = @"C:\MyApp\Google\google_client.json";
        private const string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";
        private GoogleSheetsClient _client = new(JsonClient, TableId);

        [TestMethod]
        public void GetSheets()
        {
            var sheets = _client.GetSheetsAsync().Result;
            Assert.IsNotNull(sheets);
        }

        [TestMethod]
        public void AddRemoveSheet()
        {
            var sheetName = "Test";
            var sheets = _client.GetSheetsAsync().Result;
            if (sheets.Contains(sheetName))
                _client.RemoveSheetAsync(sheetName).Wait();

            _client.AddSheetAsync(sheetName).Wait();
            sheets = _client.GetSheetsAsync().Result;
            Assert.IsTrue(sheets.Any(x => x.EqualsIgnoreCase(sheetName)));

            _client.RemoveSheetAsync(sheetName).Wait();
            sheets = _client.GetSheetsAsync().Result;
            Assert.IsFalse(sheets.Any(x => x.EqualsIgnoreCase(sheetName)));
        }
    }
}