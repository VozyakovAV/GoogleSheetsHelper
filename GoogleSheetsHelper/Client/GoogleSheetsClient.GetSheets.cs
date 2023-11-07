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
    }
}
