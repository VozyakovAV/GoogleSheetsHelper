namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task<IList<string>> GetSheetsAsync(CancellationToken ct = default)
        {
            var response = await _service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
            var sheets = response.Sheets
                .Where(x => x.Properties.SheetId != null)
                .Select(x => x.Properties.Title)
                .ToList();
            return sheets;
        }

        internal async Task<IList<SheetData>> GetSheetsDataAsync(CancellationToken ct = default)
        {
            var response = await _service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
            var sheets = response.Sheets
                .Where(x => x.Properties.SheetId != null)
                .Select(x => new SheetData(x.Properties.SheetId.Value, x.Properties.Title))
                .ToList();
            return sheets;
        }

        public async Task AddSheetAsync(string title, int? columnCount = null, int? rowCount = null,
            CancellationToken ct = default)
        {
            await AddSheetsAsync(new[] {title}, columnCount, rowCount, ct).ConfigureAwait(false);
        }

        public async Task AddSheetsAsync(IEnumerable<string> titles, int? columnCount = null, int? rowCount = null,
            CancellationToken ct = default)
        {
            var sheets = await GetSheetsAsync(ct).ConfigureAwait(false);
            var sheetsNoExist = titles.Where(x => !sheets.Any(y => x.EqualsIgnoreCase(y))).ToList();
            if (!sheetsNoExist.Any())
                return;

            var requestsBody = GetGoogleRequest();
            foreach (var sheet in sheetsNoExist)
            {
                requestsBody.Requests.Add(CreateAddSheetRequestAsync(sheet, columnCount, rowCount));
            }

            await _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private Request CreateAddSheetRequestAsync(string sheetName, int? columnCount = null, int? rowCount = null)
        {
            var request = new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        Title = sheetName,
                        GridProperties = new GridProperties
                        {
                            ColumnCount = columnCount,
                            RowCount = rowCount,
                        }
                    }
                }
            };
            return request;
        }

        public async Task RemoveSheetAsync(string title, CancellationToken ct = default)
        {
            var sheetId = await GetSheetIdAsync(title, ct).ConfigureAwait(false);
            if (sheetId == null)
                return;

            var request = GetGoogleRequest(new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = sheetId
                }
            });
            await _service.Spreadsheets.BatchUpdate(request, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private async Task<int?> GetSheetIdAsync(string title, CancellationToken ct = default)
        {
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var sheet = sheets.FirstOrDefault(x => x.Title.EqualsIgnoreCase(title));
            return sheet?.Id;
        }
    }
}
