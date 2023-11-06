using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

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
                var r = await CreateUpdateRequest(req).ConfigureAwait(false);
                requestBody.Requests.Add(r);
            }
            var request = _service.Value.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId);
            var response = await request.ExecuteAsync(ct).ConfigureAwait(false);
        }

        private async Task<Request> CreateUpdateRequest(GoogleSheetUpdateRequest r)
        {
            var sheetId = GetSheetId(r.SheetName);
            if (sheetId == null)
            {
                await AddSheet(r.SheetName).ConfigureAwait(false);
                sheetId = GetSheetId(r.SheetName);
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
