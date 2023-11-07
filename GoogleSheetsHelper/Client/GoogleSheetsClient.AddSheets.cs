namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
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

        private static Request CreateAddSheetRequestAsync(string sheetName, int? columnCount = null, int? rowCount = null)
        {
            return new Request
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
        }
    }
}
