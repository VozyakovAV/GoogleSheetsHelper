namespace GoogleSheetsHelper
{
    public static partial class GoogleUtils
    {
        /// <summary>
        /// Вставить или обновить значения в гугл таблицу по ключу
        /// </summary>
        /// <param name="client">Клиент</param>
        /// <param name="sheetName">Название листа</param>
        /// <param name="columnKey">Номер колонки с ключом</param>
        /// <param name="columnStart">Начальный номер колонки для вставки значений</param>
        /// <param name="values">Значения (ключ (строка), массив значений)</param>
        public static async Task WriteByKey(
            GoogleSheetsClient client, 
            string sheetName, 
            IDictionary<string, object[]> values, 
            int columnKey, 
            int columnStart, 
            string[] titles = null, 
            CancellationToken ct = default)
        {
            if (values == null || values.Count == 0)
                return;

            // Получаем данные
            var data = await client.GetOrAddSheet(sheetName, ct).ConfigureAwait(false);
            var requestsAppend = new List<GoogleSheetAppendRequest>();
            var requestsUpdate = new List<GoogleSheetUpdateRequest>();

            // Если нет таблицы, то сначала вставляем заголовки
            if (data == null && titles != null)
            {
                requestsAppend.Add(CreateAppendRequest(sheetName, columnKey, "Key", columnStart, titles));
            }

            // Вставляем контент
            foreach (var item in values)
            {
                var row = GetRowByValue(data, columnKey, item.Key);
                if (row == null)
                    requestsAppend.Add(CreateAppendRequest(sheetName, columnKey, item.Key, columnStart, item.Value));
                else
                    requestsUpdate.AddRange(CreateUpdateRequests(sheetName, columnStart, row.Value, item.Value));
            }

            // Отправляем данные
            if (requestsAppend.Count > 0)
                await client.AppendAsync(requestsAppend, ct).ConfigureAwait(false);
            if (requestsUpdate.Count > 0)
                await client.UpdateAsync(requestsUpdate, ct).ConfigureAwait(false);
        }

        private static int? GetRowByValue(IList<IList<object>> list, int column, string value)
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

        private static GoogleSheetAppendRequest CreateAppendRequest(string sheetName, int columnKey, string valueKey, 
            int columnStart, object[] values)
        {
            var row = new GoogleSheetRow();
            var countColumns = Math.Max(columnKey, columnStart + values.Length);

            for (int i = 0; i < countColumns; i++)
                row.Add(null);

            for (int i = 0; i < values.Length; i++)
                row[i + columnStart] = GoogleSheetCell.Create(values[i]);
            
            row[columnKey] = GoogleSheetCell.Create(valueKey);

            var request = new GoogleSheetAppendRequest(sheetName)
            {
                Rows = { row },
            };
            return request;
        }

        private static IEnumerable<GoogleSheetUpdateRequest> CreateUpdateRequests(string sheetName, int columnStart, 
            int rowStart, object[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                var value = values[i];
                if (value == null) continue;

                var row = new GoogleSheetRow
                {
                    GoogleSheetCell.Create(value)
                };

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
