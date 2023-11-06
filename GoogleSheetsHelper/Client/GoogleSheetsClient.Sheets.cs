namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        /// <summary>Получить список листов</summary>
        public async Task<IList<string>> GetSheetsAsync(CancellationToken ct = default)
        {
            await UpdateSheetsAsync(ct).ConfigureAwait(false);
            return _sheets.Select(x => x.Title).ToList();
        }

        /// <summary>
        /// Создать лист
        /// </summary>
        /// <param name="title">Название листа</param>
        /// <param name="columnCount">Количество колонок</param>
        /// <param name="rowCount">Количество строк</param>
        public async Task AddSheetAsync(string title, int? columnCount = null, int? rowCount = null,
            bool ifExistNoAdd = false, CancellationToken ct = default)
        {
            if (ifExistNoAdd)
            {
                var sheets = await GetSheetsAsync(ct).ConfigureAwait(false);
                if (sheets.Any(x => x.EqualsIgnoreCase(title)))
                    return;
            }

            var requests = new BatchUpdateSpreadsheetRequest
            {
                Requests =
                {
                    new Request
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
                    }
                }
            };
            await _service.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        public async Task DeleteSheetAsync(string title, CancellationToken ct = default)
        {
            var sheetId = await GetSheetIdAsync(title, ct).ConfigureAwait(false) ?? throw new ArgumentException($"Not found sheet {title}");
            var requests = new BatchUpdateSpreadsheetRequest
            { 
                Requests =
                {
                    new Request
                    {
                        DeleteSheet = new DeleteSheetRequest
                        {
                            SheetId = sheetId
                        }
                    }
                }
            };
            await _service.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private async Task UpdateSheetsAsync(CancellationToken ct = default)
        {
            var response = await _service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
            _sheets = response.Sheets
                .Where(x => x.Properties.SheetId != null)
                .Select(x => new Models.GoogleSheet(x.Properties.SheetId.Value, x.Properties.Title))
                .ToList();
        }

        private async Task<int?> GetSheetIdAsync(string title, CancellationToken ct = default)
        {
            var sheet = _sheets.FirstOrDefault(x => x.Title.EqualsIgnoreCase(title));
            if (sheet == null)
            {
                await UpdateSheetsAsync(ct).ConfigureAwait(false);
                sheet = _sheets.FirstOrDefault(x => x.Title.EqualsIgnoreCase(title));
            }
            return sheet?.Id;
        }
    }
}
