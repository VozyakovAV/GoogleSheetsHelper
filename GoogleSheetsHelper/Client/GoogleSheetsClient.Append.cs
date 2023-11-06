namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task AppendAsync(IEnumerable<GoogleSheetAppendRequest> data, CancellationToken ct = default)
        {
            var sheetsName = data.Select(x => x.SheetName).Distinct().ToList();
            await AddSheetsAsync(sheetsName, ct: ct).ConfigureAwait(false);
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var requestsBody = GetGoogleRequest();

            foreach (var req in data)
            {
                req.SheetId = sheets.First(x => x.Title.EqualsIgnoreCase(req.SheetName)).Id;
                requestsBody.Requests.Add(CreateAppendRequest(req));
            }

            var request = _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId);
            await request.ExecuteAsync(ct).ConfigureAwait(false);
        }

        private Request CreateAppendRequest(GoogleSheetAppendRequest r)
        {
            var listRowData = new List<RowData>();
            var request = new Request
            {
                AppendCells = new AppendCellsRequest
                {
                    SheetId = r.SheetId,
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
