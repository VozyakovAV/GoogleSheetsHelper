//https://www.twilio.com/blog/2017/03/google-spreadsheets-and-net-core.html?utm_source=youtube&utm_medium=video&utm_campaign=google-sheets-dotnet
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient : IDisposable
    {
        public const string DateTimeFormatDefault = "dd.MM.yyyy hh:mm:ss";

        public string FileClientJson { get; private set; }
        public string SpreadsheetId { get; private set; }
        public string DateTimeFormat { get; set; } = DateTimeFormatDefault;

        private static readonly string ApplicationName = "GoogleSheetsHelper";
        
        private Lazy<SheetsService> _service;
        private Lazy<Spreadsheet> _spreadsheet;

        public GoogleSheetsClient(string fileClientJson, string spreadsheetId)
        {
            FileClientJson = fileClientJson;
            SpreadsheetId = spreadsheetId;
            _service = new Lazy<SheetsService>(() => Init());
            ResetSpreadsheet();
        }

        private SheetsService Init()
        {
            GoogleCredential credential;
            using (var stream = new FileStream(FileClientJson, FileMode.Open, FileAccess.Read))
            {
                var scopes = new[] { SheetsService.Scope.Spreadsheets };
                credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private void ResetSpreadsheet()
        {
            _spreadsheet = new Lazy<Spreadsheet>(() => GetSpreadsheet().Result);
        }

        private async Task<Spreadsheet> GetSpreadsheet()
        {
            return await _service.Value.Spreadsheets.Get(SpreadsheetId).ExecuteAsync().ConfigureAwait(false);
        }

        private int? GetSheetId(string sheetName)
        {
            var sheet = _spreadsheet.Value.Sheets.FirstOrDefault(s => 
                s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
            {
                ResetSpreadsheet();
                sheet = _spreadsheet.Value.Sheets.FirstOrDefault(s => 
                    s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            }
            return sheet?.Properties.SheetId;
        }

        private CellData CreateCellData(GoogleSheetCell cell)
        {
            if (cell == null) return new CellData();
            var numberValue = cell.NumberValue;
            var numberFormat = cell.NumberFormat == null ? null : new NumberFormat
            {
                Type = "number",
                Pattern = cell.NumberFormat,
            };

            if (cell.DateTimeValue.HasValue)
            {
                numberValue = cell.DateTimeValue.Value.ToOADate();
                numberFormat = new NumberFormat
                {
                    Type = "number",
                    Pattern = cell.DateTimeFormat ?? DateTimeFormat
                };
            }

            var userEnteredValue = new ExtendedValue
            {
                StringValue = cell.StringValue,
                NumberValue = numberValue,
                BoolValue = cell.BoolValue,
            };

            CellFormat userEnteredFormat = null;

            if (numberFormat != null)
            {
                if (userEnteredFormat == null)
                    userEnteredFormat = new CellFormat();
                userEnteredFormat.NumberFormat = numberFormat;
            }

            if (cell.Bold.HasValue)
            {
                if (userEnteredFormat == null)
                    userEnteredFormat = new CellFormat();
                if (userEnteredFormat.TextFormat == null)
                    userEnteredFormat.TextFormat = new TextFormat();
                userEnteredFormat.TextFormat.Bold = cell.Bold;
            }

            if (cell.BackgroundColor.HasValue)
            {
                if (userEnteredFormat == null)
                    userEnteredFormat = new CellFormat();
                userEnteredFormat.BackgroundColor = cell.BackgroundColor.Value.ToGoogleColor();
            }

            if (cell.HorizontalAlignment.HasValue)
            {
                if (userEnteredFormat == null)
                    userEnteredFormat = new CellFormat();
                userEnteredFormat.HorizontalAlignment = cell.HorizontalAlignment.ToString();
            }

            var cellData = new CellData()
            {
                UserEnteredValue = userEnteredValue,
                UserEnteredFormat = userEnteredFormat,
            };
            return cellData;
        }

        public void Dispose()
        {
            if (_service.IsValueCreated)
                _service.Value.Dispose();
        }
    }
}
