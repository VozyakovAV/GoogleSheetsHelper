namespace GoogleSheetsHelper
{
    public class GoogleSheetAppendRequest
    {
        public string Title { get; set; }
        public IList<GoogleSheetRow> Rows { get; set; }

        public GoogleSheetAppendRequest(string title)
        {
            Title = title;
            Rows = new List<GoogleSheetRow>();
        }
    }
}
