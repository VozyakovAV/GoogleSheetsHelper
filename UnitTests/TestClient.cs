using GoogleSheetsHelper;

namespace UnitTests
{
    [TestClass]
    public class TestClient
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
        public void AddSheetDubl()
        {
            var sheetName = "Test";
            var sheets = _client.GetSheetsAsync().Result;
            if (sheets.Contains(sheetName))
                _client.RemoveSheetAsync(sheetName).Wait();

            _client.AddSheetAsync(sheetName).Wait();
            _client.AddSheetAsync(sheetName).Wait();
            sheets = _client.GetSheetsAsync().Result;
            Assert.IsTrue(sheets.Any(x => x.EqualsIgnoreCase(sheetName)));
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

        [TestMethod]
        public void AddRemoveSheets()
        {
            _client.AddSheetsAsync(new[] { "Test1", "Test2", "Test3" }).Wait();
            var sheets = _client.GetSheetsAsync().Result;
            Assert.IsTrue(sheets.Any(x => x == "Test1"));
            Assert.IsTrue(sheets.Any(x => x == "Test2"));
            Assert.IsTrue(sheets.Any(x => x == "Test3"));

            _client.RemoveSheetsAsync(new[] { "Test1", "Test2", "Test3" }).Wait();
            sheets = _client.GetSheetsAsync().Result;
            Assert.IsFalse(sheets.Any(x => x == "Test1"));
            Assert.IsFalse(sheets.Any(x => x == "Test2"));
            Assert.IsFalse(sheets.Any(x => x == "Test3"));
        }

        [TestMethod]
        public void AppendData()
        {
            var sheetName = "Test";
            _client.AddSheetAsync(sheetName).Wait();
            var request = new GoogleSheetAppendRequest(sheetName)
            {
                Rows = new[]
                {
                    new GoogleSheetRow 
                    {
                        new GoogleSheetCell("abc"),
                        new GoogleSheetCell(123),
                        new GoogleSheetCell(DateTime.Now),
                    }
                }
            };
            _client.AppendDataAsync(request).Wait();
        }

        [TestMethod]
        public void UpdateData()
        {
            var sheetName = "Test";
            _client.AddSheetAsync(sheetName).Wait();
            var request = new GoogleSheetUpdateRequest(sheetName)
            {
                ColumnStart = 1,
                RowStart = 1,
                Rows = new[]
                {
                    new GoogleSheetRow
                    {
                        new GoogleSheetCell("abc"),
                        new GoogleSheetCell(123),
                        new GoogleSheetCell(DateTime.Now),
                    }
                }
            };
            _client.UpdateDataAsync(request).Wait();
        }

        [TestMethod]
        public void GetData()
        {
            var sheetName = "Test";
            _client.AddSheetAsync(sheetName).Wait();
            var request = new GoogleSheetUpdateRequest(sheetName)
            {
                ColumnStart = 1,
                RowStart = 1,
                Rows = new[]
                {
                    new GoogleSheetRow
                    {
                        new GoogleSheetCell("abc"),
                        new GoogleSheetCell(123),
                        new GoogleSheetCell(DateTime.Now),
                    }
                }
            };
            _client.UpdateDataAsync(request).Wait();

            var data = _client.GetDataAsync("Test").Result;
        }
    }
}