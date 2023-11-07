namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task UpdateDataAsync(GoogleSheetUpdateRequest data, CancellationToken ct = default)
        {
            await UpdateAsync(new[] { data }, ct).ConfigureAwait(false);
        }
        
        public async Task UpdateAsync(IList<GoogleSheetUpdateRequest> data, CancellationToken ct = default)
        {
            var sheetsName = data.Select(x => x.SheetName).Distinct().ToList();
            await AddSheetsAsync(sheetsName, ct: ct).ConfigureAwait(false);
            var sheets = await GetSheetsDataAsync(ct).ConfigureAwait(false);
            var requestsBody = GetGoogleRequest();

            foreach (var req in data)
            {
                var sheetId = sheets.First(x => x.Title.EqualsIgnoreCase(req.SheetName)).Id;
                requestsBody.Requests.Add(CreateUpdateRequest(req, sheetId));
            }

            await _service.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        private Request CreateUpdateRequest(GoogleSheetUpdateRequest r, int sheetId)
        {
            var gc = new GridCoordinate
            {
                ColumnIndex = r.ColumnStart,
                RowIndex = r.RowStart,
                SheetId = sheetId
            };
            var request = new Request
            {
                UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" }
            };
            var listRowData = new List<RowData>();

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

            request.UpdateCells.Rows = listRowData;
            return request;
        }
    }
}
