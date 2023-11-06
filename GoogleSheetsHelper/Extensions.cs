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
    }
}
