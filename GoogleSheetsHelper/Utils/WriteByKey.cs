using System;
using System.Collections.Generic;
using System.Threading.Tasks;

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
                    requestsUpdate.Add(CreateUpdateRequest(sheetName, columnStartWrite, row.Value, item.Value));
            }

            if (requestsAppend.Count > 0)
                await client.Append(requestsAppend);
            if (requestsUpdate.Count > 0)
                await client.Update(requestsUpdate);
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
            {
                row.Add(GoogleSheetCell.Create(value));
            }

            var request = new GoogleSheetUpdateRequest(sheetName)
            {
                ColumnStart = columnStart,
                RowStart = rowStart,
                Rows = { row },
            };
            return request;
        }
    }
}
