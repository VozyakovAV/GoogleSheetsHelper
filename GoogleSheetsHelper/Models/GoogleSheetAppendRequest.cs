﻿namespace GoogleSheetsHelper
{
    public class GoogleSheetAppendRequest
    {
        public string SheetName { get; set; }
        public IList<GoogleSheetRow> Rows { get; set; }

        public GoogleSheetAppendRequest(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<GoogleSheetRow>();
        }
    }
}
