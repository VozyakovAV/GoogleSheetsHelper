using GoogleSheetsHelper;

namespace UnitTests
{
    [TestClass]
    public class TestUtils
    {
        private const string JsonClient = @"C:\MyApp\Google\google_client.json";
        private const string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";
        private GoogleSheetsClient _client = new(JsonClient, TableId);

        [TestMethod]
        public void UpdateData()
        {
            var sheetName = "Test";
            var values = new List<object[]>
            {
                new object[] {  1, 2, 3 },
                new object[] {  "a", "b", "c" },
                new object[] {  DateTime.Now },
            };

            GoogleUtils.Update(_client, sheetName, values).Wait();
        }

        [TestMethod]
        public void AppendData()
        {
            var sheetName = "Test";
            var values = new Dictionary<object, object[]>
            {
                { "1", new object[] { 1, 2, 3 } },
                { "4", new object[] { 1, 2, 3 } },
            };

            GoogleUtils.AppendByKey(_client, sheetName, values).Wait();
        }
    }
}
