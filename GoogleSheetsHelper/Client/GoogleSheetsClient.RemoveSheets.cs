namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task RemoveSheetAsync(string title, CancellationToken ct = default)
        {
            await RemoveSheetsAsync(new[] { title }, ct).ConfigureAwait(false);
        }

        public async Task RemoveSheetsAsync(IEnumerable<string> titles, CancellationToken ct = default)
        {
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var sheetsExist = titles.Distinct().Where(x => sheets.Any(y => x.EqualsIgnoreCase(y.Title))).ToList();
            if (!sheetsExist.Any())
                return;

            var requestsBody = GetGoogleRequest();
            foreach (var sheet in sheetsExist)
            {
                var sheetId = sheets.First(x => x.Title.EqualsIgnoreCase(sheet)).Id;
                requestsBody.Requests.Add(CreateRemoveSheetRequestAsync(sheetId));
            }

            await _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private static Request CreateRemoveSheetRequestAsync(int sheetId)
        {
            return new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = sheetId
                }
            };
        }
    }
}
