namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient : IDisposable
    {
        public const string ApplicationName = "GoogleSheetsHelper";
        public const string DateTimeFormatDefault = "dd.MM.yyyy hh:mm:ss";
        public string FileClientJson { get; }
        public string SpreadsheetId { get; }
        public string DateTimeFormat { get; set; }
        
        private readonly SheetsService _service;

        public GoogleSheetsClient(string fileClientJson, string spreadsheetId)
        {
            FileClientJson = fileClientJson;
            SpreadsheetId = spreadsheetId;
            DateTimeFormat = DateTimeFormatDefault;
            _service = GetService();
        }

        private SheetsService GetService()
        {
            using var stream = new FileStream(FileClientJson, FileMode.Open, FileAccess.Read);
            var scopes = new[] { SheetsService.Scope.Spreadsheets };
            var credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private CellData CreateCellData(GoogleSheetCell cell)
        {
            if (cell == null) 
                return new CellData();
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
                userEnteredFormat ??= new CellFormat();
                userEnteredFormat.NumberFormat = numberFormat;
            }

            if (cell.Bold.HasValue)
            {
                userEnteredFormat ??= new CellFormat();
                userEnteredFormat.TextFormat ??= new TextFormat();
                userEnteredFormat.TextFormat.Bold = cell.Bold;
            }

            if (cell.BackgroundColor.HasValue)
            {
                userEnteredFormat ??= new CellFormat();
                userEnteredFormat.BackgroundColor = cell.BackgroundColor.Value.ToGoogleColor();
            }

            if (cell.HorizontalAlignment.HasValue)
            {
                userEnteredFormat ??= new CellFormat();
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
            _service.Dispose();
        }
    }
}
