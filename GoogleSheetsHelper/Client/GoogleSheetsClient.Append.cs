using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task Append(IList<GoogleSheetAppendRequest> data, CancellationToken ct = default)
        {
            var requestsBody = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var req in data)
            {
                var r = CreateAppendRequest(req);
                requestsBody.Requests.Add(r);
            }
            var request = _service.Value.Spreadsheets.BatchUpdate(requestsBody, SpreadsheetId);
            var response = await request.ExecuteAsync(ct).ConfigureAwait(false);
        }

        private Request CreateAppendRequest(GoogleSheetAppendRequest r)
        {
            var sheetId = GetSheetId(r.SheetName);
            if (sheetId == null)
                throw new ArgumentException($"Не найдена таблица {r.SheetName}");

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
