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
        public static async Task AppendByKey(
            GoogleSheetsClient client, 
            string sheetName, 
            IDictionary<object, object[]> values, 
            int columnKey = 0, 
            int columnStart = 0, 
            string[] titles = null, 
            CancellationToken ct = default)
        {
            if (values == null || values.Count == 0)
                return;

            // Получаем данные
            var data = await client.GetOrAddSheet(sheetName, ct).ConfigureAwait(false);
            var requestsAppend = new List<GoogleSheetAppendRequest>();

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
                {
                    requestsAppend.Add(CreateAppendRequest(sheetName, columnKey, item.Key, columnStart, item.Value));
                }
            }

            // Отправляем данные
            if (requestsAppend.Count > 0)
            {
                await client.AppendAsync(requestsAppend, ct).ConfigureAwait(false);
            }
        }
    }
}
