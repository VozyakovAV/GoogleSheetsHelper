using System;
using System.Drawing;

namespace AndreyPro.GoogleSheetsHelper
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
        public string NumberPattern { get; set; }

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
    }

    public enum HorizontalAlignment
    {
        Left,
        Right,
        Center
    }
}
