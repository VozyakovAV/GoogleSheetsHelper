using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public void Update(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange() { Values = values };
            var updateRequest = _service.Value.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var result = updateRequest.Execute();
        }

        public void Update(IList<GoogleSheetUpdateRequest> data)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var req in data)
            {
                var r = CreateUpdateRequest(req);
                requests.Requests.Add(r);
            }
            _service.Value.Spreadsheets.BatchUpdate(requests, SpreadsheetId).Execute();
        }

        private Request CreateUpdateRequest(GoogleSheetUpdateRequest r)
        {
            var sheetId = GetSheetId(r.SheetName);
            if (sheetId == null)
                throw new ArgumentException($"Не найдена таблица {r.SheetName}");

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

                foreach (var cell in row.Cells)
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
