using GoogleSheetsHelper;
using System.Globalization;
using System.Text;

namespace ExportToGoogleConsole
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Init();
            var cmd = ParseArg(args, "/cmd:");
            var fileGoogle = ParseArg(args, "/fileGoogle:");
            var file = ParseArg(args, "/fileSource:");
            var table = ParseArg(args, "/table:");
            var sheet = ParseArg(args, "/sheet:");
            var row = ParseArg(args, "/row:");
            var column = ParseArg(args, "/column:");

            _ = int.TryParse(row, out var rowStart);
            _ = int.TryParse(column, out var columnStart);

            if (!File.Exists(fileGoogle))
                throw new FileNotFoundException(fileGoogle);
            
            using var client = new GoogleSheetsClient(fileGoogle, table);
            var ct = new CancellationTokenSource(30000).Token;

            if (cmd.EqualsIgnoreCase("append"))
            {
                var items = GetDictionaryValues(file);
                await GoogleUtils.AppendByKey(client, sheet, items, rowStart, columnStart, ct: ct).ConfigureAwait(false);
            }
            else if (cmd.EqualsIgnoreCase("update"))
            {
                var items = GetListValues(file);
                await GoogleUtils.Update(client, sheet, items, rowStart, columnStart, ct: ct).ConfigureAwait(false);
            }
            else
            {
                throw new InvalidOperationException($"Unknown cmd: {cmd}");
            }
        }

        private static IList<object[]> GetListValues(string fileName)
        {
            var items = new List<object[]>();
            var lines = File.ReadAllLines(fileName, Encoding.GetEncoding(1251));
            foreach (var line in lines)
            {
                var values = line.TrimEnd(';').Split(';');
                var objects = new List<object>();
                foreach (var value in values)
                {
                    objects.Add(GoogleUtils.ParseObject(value));
                }
                items.Add(objects.ToArray());
            }
            return items;
        }

        private static IDictionary<object, object[]> GetDictionaryValues(string fileName)
        {
            var items = new Dictionary<object, object[]>();
            var lines = File.ReadAllLines(fileName, Encoding.GetEncoding(1251));
            foreach (var line in lines)
            {
                var values = line.TrimEnd(';').Split(';');
                var objects = new List<object>();
                object key = "";
                for (int i = 0; i < values.Length; i++)
                {
                    var value = values[i];
                    if (i == 0)
                        key = GoogleUtils.ParseObject(value);
                    objects.Add(GoogleUtils.ParseObject(value));
                }
                items.Add(key, objects.ToArray());
            }
            return items;
        }

        private static void Init()
        {
            // Поддержка кодировки 1251
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var culture = new CultureInfo("ru-RU")
            {
                DateTimeFormat =
                {
                    TimeSeparator = ":",
                    DateSeparator = "/"
                },
                NumberFormat =
                {
                    NumberDecimalSeparator = ".",
                    NumberGroupSeparator = " ",
                    PercentDecimalSeparator = ".",
                    CurrencyDecimalSeparator = "."
                }
            };
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
        }

        private static string ParseArg(string[] args, string prefix)
        {
            return args.FirstOrDefault(x => x.StartsWith(prefix, StringComparison.InvariantCultureIgnoreCase))?.Replace(prefix, "")?.Trim();
        }
    }
}