using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AndreyPro.GoogleSheetsHelper
{
    public static class GoogleUtils
    {
        public static async Task WriteByKey(GoogleSheetsClient client, string sheetName, IList<WriteItem> items)
        {
            await client.AddSheetIfNotExist(sheetName);
            var data = client.GetAsync(sheetName).Result;

            var requestsAppend = new List<GoogleSheetAppendRequest>();
            var requestsUpdate = new List<GoogleSheetUpdateRequest>();
            foreach (var item in items)
            {
                var row = GetRow(data, item.KeyColumn, item.KeyValue);
                if (row == null)
                    requestsAppend.Add(CreateAppendRequest(sheetName, item.KeyColumn, item.KeyValue, item.WriteColumn, item.WriteItems));
                else
                    requestsUpdate.Add(CreateUpdateRequest(sheetName, item.WriteColumn, row.Value, item.WriteItems));
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

    public class WriteItem
    {
        public int KeyColumn { get; set; }
        public string KeyValue { get; set; }

        public int WriteColumn { get; set; }
        public object[] WriteItems { get; set; }

        public WriteItem(int keyColumn, string keyValue, int writeColumn, object[] writeItems)
        {
            KeyColumn = keyColumn;
            KeyValue = keyValue;
            WriteColumn = writeColumn;
            WriteItems = writeItems;
        }
    }
}
