namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task AppendDataAsync(GoogleSheetAppendRequest data, CancellationToken ct = default)
        {
            await AppendDataAsync(data, ct).ConfigureAwait(false);
        }

        public async Task AppendAsync(IEnumerable<GoogleSheetAppendRequest> data, CancellationToken ct = default)
        {
            var sheetsName = data.Select(x => x.SheetName).Distinct().ToList();
            await AddSheetsAsync(sheetsName, ct: ct).ConfigureAwait(false);
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var requestsBody = GetGoogleRequest();

            foreach (var req in data)
            {
                var sheetId = sheets.First(x => x.Title.EqualsIgnoreCase(req.SheetName)).Id;
                requestsBody.Requests.Add(CreateAppendRequest(req, sheetId));
            }

            await _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private Request CreateAppendRequest(GoogleSheetAppendRequest r, int sheetId)
        {
            var listRowData = new List<RowData>();
            var request = new Request
            {
                AppendCells = new AppendCellsRequest
                {
                    SheetId = sheetId,
                    Rows = listRowData,
                    Fields = "*",
                },
            };

            foreach (var row in r.Rows)
            {
                var listCellData = new List<CellData>();
                
                foreach (var cell in row)
                {
                    var cellData = CreateCellData(cell);
                    listCellData.Add(cellData);
                }

                var rowData = new RowData() { Values = listCellData };
                listRowData.Add(rowData);
            }

            return request;
        }
    }
}
