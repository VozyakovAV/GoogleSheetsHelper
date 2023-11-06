using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsHelper
{
    /// <summary>
    /// Позволяет обновлять гугл таблицу из разных потоков.
    /// Собирает данные через метод Add.
    /// Пересылает все данные в гугл таблицу через метод Send и очищает внутреннюю коллекцию данных.
    /// Использует GoogleUtils.WriteByKey.
    /// </summary>
    public class GoogleSheetUpdater
    {
        private GoogleSheetsClient _client;
        private string _sheetName;
        private int _columnKey;
        private int _columnStartWrite;

        private readonly Dictionary<string, object[]> _tempValues = new Dictionary<string, object[]>();
        private readonly object _lockUpdateValues = new object();
        private readonly SemaphoreSlim _lockSend = new SemaphoreSlim(1, 1);

        public GoogleSheetUpdater(GoogleSheetsClient client, string sheetName, int columnKey, int columnStartWrite)
        {
            _client = client;
            _sheetName = sheetName;
            _columnKey = columnKey;
            _columnStartWrite = columnStartWrite;
        }

        public int CountValues => _tempValues.Count;

        public void Add(Dictionary<string, object[]> items)
        {
            lock (_lockUpdateValues)
            {
                foreach (var item in items)
                    _tempValues[item.Key] = item.Value;
            }
        }

        public void Add(string key, params object[] values)
        {
            lock (_lockUpdateValues)
            {
                _tempValues[key] = values;
            }
        }

        public async Task Send(bool isThrowException = true, CancellationToken ct = default)
        {
            await _lockSend.WaitAsync().ConfigureAwait(false);
            try
            {
                var items = GetValues();
                if (items == null || items.Count == 0)
                    return;
                await GoogleUtils.WriteByKey(_client, _sheetName, items, _columnKey, _columnStartWrite, ct: ct).ConfigureAwait(false);
                RemoveValues(items);
            }
            catch
            {
                if (isThrowException)
                    throw;
            }
            finally
            {
                _lockSend.Release();
            }
        }

        private Dictionary<string, object[]> GetValues()
        {
            lock (_lockUpdateValues)
            {
                return _tempValues.ToDictionary(x => x.Key, x => x.Value);
            }
        }

        private void RemoveValues(Dictionary<string, object[]> values)
        {
            lock (_lockUpdateValues)
            {
                foreach (var value in values)
                {
                    if (_tempValues.TryGetValue(value.Key, out var v))
                    {
                        if (value.Value == v)
                        {
                            _tempValues.Remove(value.Key);
                        }
                    }
                }
            }
        }
    }
}
