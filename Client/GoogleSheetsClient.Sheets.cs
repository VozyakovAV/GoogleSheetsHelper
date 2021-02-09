using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public IList<string> GetSheets()
        {
            var response = _service.Value.Spreadsheets.Get(SpreadsheetId).Execute();
            var result = response.Sheets.Select(x => x.Properties.Title).ToList();
            return result;
        }

        public bool AddSheetIfNotExist(string title, int? columnCount = null, int? rowCount = null)
        {
            var sheets = GetSheets();
            if (!sheets.Any(x => string.Equals(x, title, StringComparison.OrdinalIgnoreCase)))
            {
                AddSheet(title, columnCount, rowCount);
                return true;
            }
            return false;
        }

        public void AddSheet(string title, int? columnCount = null, int? rowCount = null)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            var request = new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties 
                    { 
                        Title = title,
                        GridProperties = new GridProperties
                        {
                            ColumnCount = columnCount,
                            RowCount = rowCount,
                        }
                    }
                }
            };
            requests.Requests.Add(request);
            var result = _service.Value.Spreadsheets.BatchUpdate(requests, SpreadsheetId).Execute();
        }
    }
}
