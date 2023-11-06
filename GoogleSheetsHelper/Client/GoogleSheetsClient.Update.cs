namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        /// <summary>Записать в Google таблицу</summary>
        public async Task Update(IList<GoogleSheetUpdateRequest> data, CancellationToken ct = default)
        {
            var requestBody = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var req in data)
            {
                var r = await CreateUpdateRequest(req, ct).ConfigureAwait(false);
                requestBody.Requests.Add(r);
            }
            var request = _service.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId);
            await request.ExecuteAsync(ct).ConfigureAwait(false);
        }

        private async Task<Request> CreateUpdateRequest(GoogleSheetUpdateRequest r, CancellationToken ct = default)
        {
            var sheetId = await GetSheetIdAsync(r.SheetName, ct);
            if (sheetId == null)
            {
                await AddSheetAsync(r.SheetName, ct: ct).ConfigureAwait(false);
                sheetId = await GetSheetIdAsync(r.SheetName, ct).ConfigureAwait(false);
            }

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
