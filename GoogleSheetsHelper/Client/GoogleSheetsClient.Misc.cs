using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsHelper
{
    public partial class GoogleSheetsClient
    {
        public async Task Clear(string range, CancellationToken ct = default)
        {
            var requestBody = new ClearValuesRequest();
            var deleteRequest = _service.Value.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var result = await deleteRequest.ExecuteAsync(ct).ConfigureAwait(false);
        }

        public async Task<IList<IList<object>>> Get(string range, CancellationToken ct = default)
        {
            var request = _service.Value.Spreadsheets.Values.Get(SpreadsheetId, range);
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            request.DateTimeRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
            var response = await request.ExecuteAsync(ct).ConfigureAwait(false);
            return response.Values;
        }

        public async Task<IList<IList<object>>> GetOrAddSheet(string range, CancellationToken ct = default)
        {
            try
            {
                return await Get(range, ct);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Unable to parse range")) // не найдена таблица
                {
                    await AddSheet(range, ct: ct);
                    return await Get(range, ct);
                }
                else
                {
                    throw ex;
                }
            }
        }
    }
}
