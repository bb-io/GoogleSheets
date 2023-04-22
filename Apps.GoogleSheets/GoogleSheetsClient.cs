using Blackbird.Applications.Sdk.Common.Authentication;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets
{
    public class GoogleSheetsClient : SheetsService
    {
        private static Initializer GetInitializer(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
        {
            var serviceAccountConfString = authenticationCredentialsProviders.First(p => p.KeyName == "serviceAccountConfString").Value;
            string[] scopes = { SheetsService.Scope.Spreadsheets };
            ServiceAccountCredential? credential = GoogleCredential.FromJson(serviceAccountConfString)
                                                  .CreateScoped(scopes)
                                                  .UnderlyingCredential as ServiceAccountCredential;
            return new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Blackbird"
            };
        }

        public GoogleSheetsClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) : base(GetInitializer(authenticationCredentialsProviders)) { }

        public IList<IList<object>> GetSheetValues(string sheetId, string sheetName, string cellA, string cellB)
        {
            var range = $"{sheetName}!{cellA}:{cellB}";
            var request = this.Spreadsheets.Values.Get(sheetId, range);
            var response = request.Execute();
            var values = response.Values;
            return values;
        }
    }
}
