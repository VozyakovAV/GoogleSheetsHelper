using System;
using System.Drawing;

namespace GoogleSheetsHelper
{
    public class GoogleSheetCell
    {
        /// <summary>Значение ячейки: строка</summary>
        public string StringValue { get; set; }

        /// <summary>Значение ячейки: число</summary>
        public double? NumberValue { get; set; }

        /// <summary>Значение ячейки: логическое значение</summary>
        public bool? BoolValue { get; set; }

        /// <summary>Значение ячейки: дата и время</summary>
        public DateTime? DateTimeValue { get; set; }

        /// <summary>Формат числового значения</summary>
        public string NumberFormat { get; set; }

        /// <summary>Формат даты и времени</summary>
        public string DateTimeFormat { get; set; }

        /// <summary>Стиль ячейки: жирный</summary>
        public bool? Bold { get; set; }

        /// <summary>Стиль ячейки: цвет</summary>
        public Color? BackgroundColor { get; set; }

        /// <summary>Стиль ячейки: выравнивание по горизонтали</summary>
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

        public static GoogleSheetCell Create(object value)
        {
            if (value == null) return null;

            switch (value)
            {
                case string t: 
                    return new GoogleSheetCell(t);
                case byte t:
                    return new GoogleSheetCell(t);
                case short t:
                    return new GoogleSheetCell(t);
                case int t:
                    return new GoogleSheetCell(t);
                case long t:
                    return new GoogleSheetCell(t);
                case double t:
                    return new GoogleSheetCell(t);
                case decimal t:
                    return new GoogleSheetCell((double)t);
                case bool t:
                    return new GoogleSheetCell(t);
                case DateTime t:
                    return new GoogleSheetCell(t);
            }
            return new GoogleSheetCell(value.ToString());
        }
    }

    public enum HorizontalAlignment
    {
        Left,
        Right,
        Center
    }
}
