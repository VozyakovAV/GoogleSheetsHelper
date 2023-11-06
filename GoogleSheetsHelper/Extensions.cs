namespace GoogleSheetsHelper
{
    public static class Extensions
    {
        public static GoogleColor ToGoogleColor(this DrawingColor color)
        {
            return new GoogleColor
            {
                Alpha = color.A / 255f,
                Blue = color.B / 255f,
                Green = color.G / 255f,
                Red = color.R / 255f
            };
        }

        public static bool EqualsIgnoreCase(this string value, string value2)
        {
            return string.Equals(value, value2, StringComparison.OrdinalIgnoreCase);
        }
    }
}
