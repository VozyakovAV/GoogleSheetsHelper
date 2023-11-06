using Google.Apis.Util;
using GoogleSheetsHelper;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        private const string JsonClient = @"C:\MyApp\Google\google_client.json";
        private const string TableId = "1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI";

        [TestInitialize]
        public void Init()
        {

        }

        [TestMethod]
        public void GetSheets()
        {
            var client = new GoogleSheetsClient(JsonClient, TableId);
            var sheets = client.GetSheetsAsync().Result;
        }
    }
}