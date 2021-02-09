using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;

namespace AndreyPro.GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public void Clear(string range)
        {
            var requestBody = new ClearValuesRequest();
            var deleteRequest = _service.Value.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var result = deleteRequest.Execute();
        }

        public IList<IList<object>> Get(string range)
        {
            var request = _service.Value.Spreadsheets.Values.Get(SpreadsheetId, range);
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            request.DateTimeRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
            var response = request.Execute();
            var result = response.Values;
            return result;
        }
    }
}
