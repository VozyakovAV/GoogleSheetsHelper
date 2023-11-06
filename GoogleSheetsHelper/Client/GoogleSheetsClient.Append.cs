namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task Append(IEnumerable<GoogleSheetAppendRequest> data, CancellationToken ct = default)
        {
            var requestsBody = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var req in data)
            {
                var r = await CreateAppendRequest(req, ct).ConfigureAwait(false);
                requestsBody.Requests.Add(r);
            }
            var request = _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId);
            await request.ExecuteAsync(ct).ConfigureAwait(false);
        }

        private async Task<Request> CreateAppendRequest(GoogleSheetAppendRequest r, CancellationToken ct = default)
        {
            var sheetId = await GetSheetIdAsync(r.Title, ct).ConfigureAwait(false);
            if (sheetId == null)
            {
                await AddSheetAsync(r.Title, ct: ct).ConfigureAwait(false);
                sheetId = await GetSheetIdAsync(r.Title, ct).ConfigureAwait(false);
            }

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
