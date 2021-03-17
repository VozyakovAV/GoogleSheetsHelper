using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public void Append(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange() { Values = values };
            var updateRequest = _service.Value.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var result = updateRequest.Execute();
        }

        public void Append(IList<GoogleSheetAppendRequest> data)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            foreach (var req in data)
            {
                var r = CreateAppendRequest(req);
                requests.Requests.Add(r);
            }
            _service.Value.Spreadsheets.BatchUpdate(requests, SpreadsheetId).Execute();
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
