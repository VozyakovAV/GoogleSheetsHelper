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
        public DrawingColor? BackgroundColor { get; set; }

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
            return value switch
            {
                string t => new GoogleSheetCell(t),
                byte t => new GoogleSheetCell(t),
                short t => new GoogleSheetCell(t),
                int t => new GoogleSheetCell(t),
                long t => new GoogleSheetCell(t),
                double t => new GoogleSheetCell(t),
                decimal t => new GoogleSheetCell((double)t),
                bool t => new GoogleSheetCell(t),
                DateTime t => new GoogleSheetCell(t),
                _ => new GoogleSheetCell(value.ToString()),
            };
        }
    }

    public enum HorizontalAlignment
    {
        Left,
        Right,
        Center
    }
}
