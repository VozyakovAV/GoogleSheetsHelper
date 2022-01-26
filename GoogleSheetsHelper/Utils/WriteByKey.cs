using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Timer = System.Timers.Timer;

namespace GoogleSheetsHelper
{
    public static class GoogleUtils
    {
        private static IDictionary<PlanWriteByKey, Dictionary<string, object[]>> _plansByTimer
            = new Dictionary<PlanWriteByKey, Dictionary<string, object[]>>();

        /// <summary>
        /// Вставить или обновить значения в гугл таблицу по ключу
        /// </summary>
        /// <param name="client">Клиент</param>
        /// <param name="sheetName">Название листа</param>
        /// <param name="columnKey">Номер колонки с ключом</param>
        /// <param name="columnStartWrite">Начальный номер колонки для вставки значений</param>
        /// <param name="items">Значения (ключ (строка), массив значений)</param>
        public static async Task WriteByKey(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite, 
            Dictionary<string, object[]> items, CancellationToken ct = default)
        {
            var data = await client.GetOrAddSheet(sheetName, ct);
            
            var requestsAppend = new List<GoogleSheetAppendRequest>();
            var requestsUpdate = new List<GoogleSheetUpdateRequest>();
            foreach (var item in items)
            {
                var row = GetRow(data, columnKey, item.Key);
                if (row == null)
                    requestsAppend.Add(CreateAppendRequest(sheetName, columnKey, item.Key, columnStartWrite, item.Value));
                else
                    requestsUpdate.AddRange(CreateUpdateRequests(sheetName, columnStartWrite, row.Value, item.Value));
            }

            if (requestsAppend.Count > 0)
                await client.Append(requestsAppend, ct);
            if (requestsUpdate.Count > 0)
                await client.Update(requestsUpdate, ct);
        }

        /// <summary>
        /// Вставить или обновить значения в гугл таблицу по ключу. Вызов произойдет через указанное время.
        /// Используется если пишут в одну таблицу несколько источников.
        /// </summary>
        /// <param name="client">Клиент</param>
        /// <param name="sheetName">Название листа</param>
        /// <param name="columnKey">Номер колонки с ключом</param>
        /// <param name="columnStartWrite">Начальный номер колонки для вставки значений</param>
        /// <param name="items">Значения (ключ (строка), массив значений)</param>
        public static void WriteByKeyWithTimer(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite, 
            Dictionary<string, object[]> items, int delayMs = 5000, Action<Exception> error = null, CancellationToken ct = default)
        {
            lock (_plansByTimer)
            {
                var plan = new PlanWriteByKey(client, sheetName, columnKey, columnStartWrite);
                if (_plansByTimer.TryGetValue(plan, out var planItems))
                {
                    foreach (var item in items)
                        planItems[item.Key] = item.Value;
                }
                else
                {
                    _plansByTimer.Add(plan, items);

                    var timer = new Timer(delayMs);
                    timer.AutoReset = false;
                    timer.Elapsed += (s, e) =>
                    {
                        try
                        {
                            lock (_plansByTimer)
                            {
                                if (_plansByTimer.TryGetValue(plan, out var itemAll))
                                {
                                    WriteByKey(client, sheetName, columnKey, columnStartWrite, itemAll, ct).Wait();
                                    _plansByTimer.Remove(plan);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            error?.Invoke(ex.GetBaseException());
                        }
                    };
                    timer.Start();
                }
            }
        }

        private static int? GetRow(IList<IList<object>> list, int column, string value)
        {
            if (list == null) return null;
            for (int i = 0; i < list.Count; i++)
            {
                var items = list[i];
                if (items.Count > column && items[column]?.ToString() == value)
                    return i;
            }
            return null;
        }

        private static GoogleSheetAppendRequest CreateAppendRequest(string sheetName, int keyColumn, string keyValue, int columnStart, object[] values)
        {
            var row = new GoogleSheetRow();
            var n = Math.Max(keyColumn, columnStart + values.Length);

            for (int i = 0; i < n; i++)
                row.Add(null);

            for (int i = 0; i < values.Length; i++)
                row[i + columnStart] = GoogleSheetCell.Create(values[i]);
            
            row[keyColumn] = GoogleSheetCell.Create(keyValue);

            var request = new GoogleSheetAppendRequest(sheetName)
            {
                Rows = { row },
            };
            return request;
        }

        private static IEnumerable<GoogleSheetUpdateRequest> CreateUpdateRequests(string sheetName, int columnStart, int rowStart, object[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                var value = values[i];
                if (value == null) continue;

                var row = new GoogleSheetRow();
                row.Add(GoogleSheetCell.Create(value));

                var request = new GoogleSheetUpdateRequest(sheetName)
                {
                    ColumnStart = columnStart + i,
                    RowStart = rowStart,
                    Rows = { row },
                };
                yield return request;
            }
        }

        private class PlanWriteByKey
        {
            public GoogleSheetsClient Client { get; }
            public string SheetName { get; }
            public int ColumnKey { get; }
            public int ColumnStartWrite { get; }

            private int _hash;

            public PlanWriteByKey(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite)
            {
                Client = client;
                SheetName = sheetName;
                ColumnKey = columnKey;
                ColumnStartWrite = columnStartWrite;
                _hash = new { Client.SpreadsheetId, SheetName, ColumnKey, ColumnStartWrite }.GetHashCode();
            }

            public override int GetHashCode() => _hash;

            public override bool Equals(object obj)
            {
                return _hash == obj?.GetHashCode();
            }
        }
    }
}
