namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        /// <summary>Получить список листов</summary>
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

        /// <summary>
        /// Создать лист
        /// </summary>
        /// <param name="title">Название листа</param>
        /// <param name="columnCount">Количество колонок</param>
        /// <param name="rowCount">Количество строк</param>
        public async Task AddSheetAsync(string title, int? columnCount = null, int? rowCount = null,
            bool noAddIfExist = false, CancellationToken ct = default)
        {
            if (noAddIfExist)
            {
                var sheets = await GetSheetsAsync(ct).ConfigureAwait(false);
                if (sheets.Any(x => x.EqualsIgnoreCase(title)))
                    return;
            }

            var request = GetGoogleRequest(new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        Title = title,
                        GridProperties = new GridProperties
                        {
                            ColumnCount = columnCount,
                            RowCount = rowCount,
                        }
                    }
                }
            });
            await _service.Spreadsheets.BatchUpdate(request, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        public async Task RemoveSheetAsync(string title, CancellationToken ct = default)
        {
            var sheetId = await GetSheetIdAsync(title, ct).ConfigureAwait(false) ?? throw new ArgumentException($"Not found sheet {title}");
            var request = GetGoogleRequest(new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = sheetId
                }
            });
            await _service.Spreadsheets.BatchUpdate(request, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private static BatchUpdateSpreadsheetRequest GetGoogleRequest(params Request[] requests)
        {
            var res = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var request in requests)
                res.Requests.Add(request);
            return res;
        }

        private async Task<int?> GetSheetIdAsync(string title, CancellationToken ct = default)
        {
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var sheet = sheets.FirstOrDefault(x => x.Title.EqualsIgnoreCase(title));
            return sheet?.Id;
        }
    }
}
