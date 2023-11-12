namespace GoogleSheetsHelper
{
    public static partial class GoogleUtils
    {
        /// <summary>
        /// Вставить или обновить значения в гугл таблицу по ключу
        /// </summary>
        public static async Task Update(
            GoogleSheetsClient client, 
            string sheetName, 
            IList<object[]> values, 
            int startRow = 0,
            int startColumn = 0, 
            string[] titles = null,
            CancellationToken ct = default)
        {
            if (values == null || values.Count == 0)
                return;

            var requestsAppend = new List<GoogleSheetUpdateRequest>();
            var curRow = startRow;
            var emptyValues = values[0].Select(x => "").ToArray();

            // Если нет таблицы, то сначала вставляем заголовки
            if (titles != null)
            {
                requestsAppend.AddRange(CreateUpdateRequests2(sheetName, curRow++, startColumn, titles));
            }

            // Вставляем контент
            foreach (var item in values)
            {
                requestsAppend.AddRange(CreateUpdateRequests2(sheetName, curRow++, startColumn, item));
            }

            // Отправляем данные
            if (requestsAppend.Count > 0)
            {
                // вставляем новые данные
                await client.UpdateAsync(requestsAppend, ct).ConfigureAwait(false);

                // удаляем лишние строки
                var r1 = startRow + values.Count + 1;
                var r2 = r1 + 10;
                await client.ClearDataAsync($"{sheetName}!{r1}:{r2}", ct).ConfigureAwait(false);
            }
        }

        private static IEnumerable<GoogleSheetUpdateRequest> CreateUpdateRequests2(string sheetName, int rowStart,
            int columnStart, object[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                var value = values[i];
                if (value == null)
                    continue;

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
