using GoogleSheetsHelper;
using System.Globalization;
using System.Text;

namespace ExportToGoogleConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {   // run.exe /file:C:\MyApp\Temp\Positions.csv /table:1_T-ENPjpjhfEiVmoSNpYZDrhfEpiwCFNeuPyNNwL2uI /sheet:TestWrite /row:0 /column:0 /key
            Init();
            var fileClientJson = "google_client.json";
            var file = ParseArg(args, "/file:");
            var table = ParseArg(args, "/table:");
            var sheet = ParseArg(args, "/sheet:");
            var row = ParseArg(args, "/row:");
            var column = ParseArg(args, "/column:");
            var byKey = args.Any(x => x.StartsWith("/key"));

            int.TryParse(row, out var rowStart);
            int.TryParse(column, out var columnStart);

            if (!File.Exists(fileClientJson))
                throw new FileNotFoundException(fileClientJson);
            
            using var client = new GoogleSheetsClient(fileClientJson, table);

            if (byKey)
            {
                var items = GetItems2(file);
                GoogleUtils.WriteByKey(client, sheet, items, rowStart, columnStart).Wait();
            }
            else
            {
                var items = GetItems(file);
                GoogleUtils.Write(client, sheet, items, rowStart, columnStart).Wait();
            }
        }

        private static IList<object[]> GetItems(string fileName)
        {
            var items = new List<object[]>();
            var lines = File.ReadAllLines(fileName, Encoding.GetEncoding(1251));
            foreach (var line in lines)
            {
                var values = line.Split(';');
                var objects = new List<object>();
                foreach (var value in values)
                {
                    if (double.TryParse(value.Replace(",", "."), out var d))
                        objects.Add(d);
                    else if (bool.TryParse(value, out var b))
                        objects.Add(b);
                    else if (DateTime.TryParse(value, out var dt))
                        objects.Add(dt);
                    else
                        objects.Add(value);
                }
                items.Add(objects.ToArray());
            }
            return items;
        }

        private static IDictionary<string, object[]> GetItems2(string fileName)
        {
            var items = new Dictionary<string, object[]>();
            var lines = File.ReadAllLines(fileName, Encoding.GetEncoding(1251));
            foreach (var line in lines)
            {
                var values = line.Split(';');
                var objects = new List<object>();
                var key = "";
                for (int i = 0; i < values.Length; i++)
                {
                    var value = values[i];
                    if (i == 0)
                        key = value;

                    if (double.TryParse(value.Replace(",", "."), out var d))
                        objects.Add(d);
                    else if (bool.TryParse(value, out var b))
                        objects.Add(b);
                    else if (DateTime.TryParse(value, out var dt))
                        objects.Add(dt);
                    else
                        objects.Add(value);
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
            return args.FirstOrDefault(x => x.StartsWith(prefix))?.Replace(prefix, "")?.Trim();
        }
    }
}