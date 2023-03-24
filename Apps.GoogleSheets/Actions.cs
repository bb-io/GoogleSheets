using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.IO;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Models.Responses;

namespace Apps.GoogleSheets
{
    [ActionList]
    public class Actions
    {
        [Action("Get cell value", Description = "Get cell value")]
        public GetCellValueResponse GetCellValue(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] GetCellValueRequest input)
        {
            var cellAdress = $"{input.Column}{input.RowId}";
            var cellValue = GetSheetValues(authenticationCredentialsProvider.Value, input.SpreadSheetId, input.SheetName, cellAdress, cellAdress)[0][0].ToString();
            return new GetCellValueResponse()
            {
                CellValue = cellValue
            };
        }

        private IList<IList<object>> GetSheetValues(string serviceAccountConfString, string sheetId, string sheetName,
            string cellA, string cellB)
        {
            var range = $"{sheetName}!{cellA}:{cellB}";
            var request = GetGoogleSheetsClient(serviceAccountConfString).Spreadsheets.Values.Get(sheetId, range);
            var response = request.Execute();
            var values = response.Values;
            return values;
        }

        private SheetsService GetGoogleSheetsClient(string serviceAccountConfString)
        {
            string[] scopes = { SheetsService.Scope.Spreadsheets };
            ServiceAccountCredential? credential = GoogleCredential.FromJson(serviceAccountConfString)
                                                  .CreateScoped(scopes)
                                                  .UnderlyingCredential as ServiceAccountCredential;
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Blackbird"
            });
            return service;
        }   
    }
}
