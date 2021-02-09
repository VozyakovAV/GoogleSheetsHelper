using System;
using System.Drawing;

namespace AndreyPro.GoogleSheetsHelper
{
    public class GoogleSheetCell
    {
        public string StringValue { get; set; }
        public double? NumberValue { get; set; }
        public bool? BoolValue { get; set; }
        public DateTime? DateTimeValue { get; set; }
        public string NumberPattern { get; set; }
        public bool? Bold { get; set; }
        public Color? BackgroundColor { get; set; }
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        public GoogleSheetCell()
        {
        }

        public GoogleSheetCell(string value)
        {
            StringValue = value;
        }

        public GoogleSheetCell(double? value)
        {
            NumberValue = value;
        }

        public GoogleSheetCell(bool? value)
        {
            BoolValue = value;
        }

        public GoogleSheetCell(DateTime? value)
        {
            DateTimeValue = value;
        }
    }

    public enum HorizontalAlignment
    {
        Left,
        Right,
        Center
    }
}
