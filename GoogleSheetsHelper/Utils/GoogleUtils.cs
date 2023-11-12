namespace GoogleSheetsHelper
{
    public static partial class GoogleUtils
    {
        public static object ParseObject(string value)
        {
            if (double.TryParse(value.Replace(",", "."), out var d))
                return d;
            else if (bool.TryParse(value, out var b))
                return b;
            else if (DateTime.TryParse(value, out var dt))
                return dt;
            else
                return value;
        }

        public static DateTime? ParseDateTime(object obj)
        {
            try
            {
                if (obj == null)
                    return DateTime.MinValue;

                if (obj is double d)
                    return DateTime.FromOADate(d);

                if (obj is long l)
                    return DateTime.FromOADate(l);

                if (obj is int i)
                    return DateTime.FromOADate(i);

                if (obj is float f)
                    return DateTime.FromOADate(f);

                if (double.TryParse(obj.ToString(), out var dt))
                    return DateTime.FromOADate(dt);
            }
            catch
            {
            }
            return null;
        }

        public static bool EqualObjects(object obj1, object obj2)
        {
            if (obj1 == null || obj2 == null)
                return false;

            if (obj1 is DateTime || obj2 is DateTime)
            {
                var dt1 = obj1 as DateTime? ?? ParseDateTime(obj1);
                var dt2 = obj2 as DateTime? ?? ParseDateTime(obj2);
                return dt1 == dt2;
            }

            return obj1.ToString() == obj2.ToString();
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            var columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static string GetRangeName(string sheetName, int startRow, int startColumn, int countRows, int countColumns)
        {
            var r1 = startRow + 1;
            var r2 = startRow + countRows;
            var c1 = GetExcelColumnName(startColumn + 1);
            var c2 = GetExcelColumnName(startColumn + countColumns);
            var res = $"{sheetName}!{c1}{r1}:{c2}{r2}";
            return res;
        }
    }
}
