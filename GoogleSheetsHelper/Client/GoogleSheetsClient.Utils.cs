//https://www.twilio.com/blog/2017/03/google-spreadsheets-and-net-core.html?utm_source=youtube&utm_medium=video&utm_campaign=google-sheets-dotnet
using System;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
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
    }
}
