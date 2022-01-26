﻿using System.Collections.Generic;

namespace GoogleSheetsHelper
{
    public class GoogleSheetUpdateRequest
    {
        public string SheetName { get; set; }
        public int ColumnStart { get; set; }
        public int RowStart { get; set; }
        public IList<GoogleSheetRow> Rows { get; set; }

        public GoogleSheetUpdateRequest(string sheetName)
        {
            this.SheetName = sheetName;
            Rows = new List<GoogleSheetRow>();
        }
    }
}
