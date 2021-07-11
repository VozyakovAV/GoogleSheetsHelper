using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Linq;
using Timer = System.Timers.Timer;

namespace AndreyPro.GoogleSheetsHelper
{
    public static class GoogleUtils
    {
        /// <summary>
        /// Вставить или обновить значения в гугл таблицу по ключу
        /// </summary>
        /// <param name="client">Клиент</param>
        /// <param name="sheetName">Название листа</param>
        /// <param name="columnKey">Номер колонки с ключом</param>
        /// <param name="columnStartWrite">Начальный номер колонки для вставки значений</param>
        /// <param name="items">Значения (ключ (строка), массив значений)</param>
        /// <returns></returns>
        public static async Task WriteByKey(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite, 
            Dictionary<string, object[]> items)
        {
            await client.AddSheetIfNotExist(sheetName);
            var data = client.GetAsync(sheetName).Result;

            var requestsAppend = new List<GoogleSheetAppendRequest>();
            var requestsUpdate = new List<GoogleSheetUpdateRequest>();
            foreach (var item in items)
            {
                var row = GetRow(data, columnKey, item.Key);
                if (row == null)
                    requestsAppend.Add(CreateAppendRequest(sheetName, columnKey, item.Key, columnStartWrite, item.Value));
                else
                    //requestsUpdate.Add(CreateUpdateRequest(sheetName, columnStartWrite, row.Value, item.Value));
                    requestsUpdate.AddRange(CreateUpdateRequests(sheetName, columnStartWrite, row.Value, item.Value));
            }

            if (requestsAppend.Count > 0)
                await client.Append(requestsAppend);
            if (requestsUpdate.Count > 0)
                await client.Update(requestsUpdate);
        }

        private static IDictionary<PlanWriteByKey, Dictionary<string, object[]>> _plansByTimer = new Dictionary<PlanWriteByKey, Dictionary<string, object[]>>();

        public static void WriteByKeyWithTimer(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite, 
            Dictionary<string, object[]> items, int intervalTimerMs, CancellationToken ct = default)
        {
            lock (_plansByTimer)
            {
                var plan = new PlanWriteByKey(client, sheetName, columnKey, columnStartWrite);
                if (_plansByTimer.TryGetValue(plan, out var p1))
                {
                    foreach (var item in items)
                        p1[item.Key] = item.Value;
                }
                else
                {
                    _plansByTimer.Add(plan, items);

                    var timer = new Timer(intervalTimerMs);
                    timer.AutoReset = false;
                    timer.Elapsed += (s, e) =>
                    {
                        lock (_plansByTimer)
                        {
                            if (_plansByTimer.TryGetValue(plan, out var itemAll))
                            {
                                WriteByKey(client, sheetName, columnKey, columnStartWrite, itemAll).Wait(ct);
                                _plansByTimer.Remove(plan);
                            }
                        }
                    };
                    timer.Start();
                }
            }
        }

        private static void T_Elapsed(object sender, ElapsedEventArgs e)
        {
            throw new NotImplementedException();
        }

        public class PlanWriteByKey
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
                _hash = new { Client, SheetName, ColumnKey, ColumnStartWrite }.GetHashCode();
            }

            public override int GetHashCode() => _hash;

            public override bool Equals(object obj)
            {
                return _hash == obj?.GetHashCode();
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

        private static GoogleSheetUpdateRequest CreateUpdateRequest(string sheetName, int columnStart, int rowStart, object[] values)
        {
            var row = new GoogleSheetRow();
            foreach (var value in values)
                row.Add(GoogleSheetCell.Create(value));

            var request = new GoogleSheetUpdateRequest(sheetName)
            {
                ColumnStart = columnStart,
                RowStart = rowStart,
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
    }
}
