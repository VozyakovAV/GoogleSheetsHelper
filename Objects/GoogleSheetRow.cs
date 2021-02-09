using System.Collections.Generic;

namespace AndreyPro.GoogleSheetsHelper
{
    public class GoogleSheetRow
    {
        public IList<GoogleSheetCell> Cells { get; set; }

        public GoogleSheetRow()
        {
            this.Cells = new List<GoogleSheetCell>();
        }

        public GoogleSheetRow(IList<GoogleSheetCell> cells)
        {
            this.Cells = cells;
        }

        public GoogleSheetRow(params GoogleSheetCell[] cells)
        {
            this.Cells = cells;
        }
    }
}
