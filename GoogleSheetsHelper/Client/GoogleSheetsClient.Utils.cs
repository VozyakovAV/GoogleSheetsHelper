namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public static bool TryParseDateTime(object obj, DateTime result)
        {
            try
            {
                if (obj == null)
                {
                    result = default;
                    return false;
                }

                if (obj is double d)
                {
                    result = DateTime.FromOADate(d);
                    return true;
                }

                if (obj is long l)
                {
                    result = DateTime.FromOADate(l);
                    return true;
                }

                if (obj is int i)
                {
                    result = DateTime.FromOADate(i);
                    return true;
                }

                if (obj is float f)
                {
                    result = DateTime.FromOADate(f);
                    return true;
                }

                if (double.TryParse(obj.ToString(), out var dt))
                {
                    result = DateTime.FromOADate(dt);
                    return true;
                }
            }
            catch
            {
            }
            return false;
        }

        public static bool TryParseDouble(object s, out double result)
        {
            try
            {
                result = Convert.ToDouble(s);
                return true;
            }
            catch
            {
                result = default;
                return false;
            }
        }
    }
}
