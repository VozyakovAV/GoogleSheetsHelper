using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        /// <summary>Получить список листов</summary>
        public async Task<IList<string>> GetSheets(CancellationToken ct = default)
        {
            var response = await _service.Value.Spreadsheets.Get(SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
            var result = response.Sheets.Select(x => x.Properties.Title).ToList();
            return result;
        }

        /// <summary>Создать лист если нету</summary>
        public async Task<bool> AddSheetIfNotExist(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default)
        {
            var sheets = await GetSheets(ct).ConfigureAwait(false);
            if (!sheets.Any(x => string.Equals(x, title, StringComparison.OrdinalIgnoreCase)))
            {
                await AddSheet(title, columnCount, rowCount, ct).ConfigureAwait(false);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Создать лист
        /// </summary>
        /// <param name="title">Название листа</param>
        /// <param name="columnCount">Количество колонок</param>
        /// <param name="rowCount">Количество строк</param>
        public async Task AddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default)
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
            var result = await _service.Value.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }

        public async Task DeleteSheet(string sheetName, CancellationToken ct = default)
        {
            var sheetId = GetSheetId(sheetName);
            if (sheetId == null)
                throw new ArgumentException($"Не найдена таблица {sheetName}");

            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            var request = new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = sheetId
                }
            };
            requests.Requests.Add(request);
            var result = await _service.Value.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(ct).ConfigureAwait(false);
        }
    }
}
