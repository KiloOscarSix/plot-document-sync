using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace plot_document_sync
{
    internal static class Program
    {
        private static readonly string[] Scopes =
        {
            DocsService.Scope.DocumentsReadonly,
            SheetsService.Scope.Spreadsheets
        };
        private const string ApplicationName = "Plot Document Sync";
        public static int SheetRows;

        private static UserCredential Login()
        {
            UserCredential credential;

            using (FileStream stream = new("credentials.json", FileMode.Open, FileAccess.Read))
            {
                const string credPath = "token.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                ).Result;
                
                Console.WriteLine($"Credential file saved to: {credPath}");
            }

            return credential;
        }

        private static void EventLoop(DocsService docsService, SheetsService sheetsService)
        {
            Console.WriteLine($"INFO: Begun event: {DateTime.Now}");
            
            // Get Document
            DocumentsResource.GetRequest docRequest = docsService.Documents.Get(HandleDocument.PlotDocumentId);
            Document document = docRequest.Execute();
            string[] documentContent = HandleDocument.GetDocumentContent(document).ToArray();
            
            // Get current spreadsheet data
            SpreadsheetsResource.ValuesResource spreadsheetResource = sheetsService.Spreadsheets.Values;
            ValueRange spreadsheet = spreadsheetResource.Get(HandleSpreadsheet.SpreadsheetId, HandleSpreadsheet.SheetRange).Execute();

            SheetRows = spreadsheet.Values.Count;
            
            spreadsheet = HandleSpreadsheet.EditSheetData(spreadsheet, documentContent);
            HandleSpreadsheet.WriteSheetData(spreadsheet, spreadsheetResource);
            
            Console.WriteLine($"INFO: End event: {DateTime.Now}");
        }
        
        private static void Main()
        {
            UserCredential credential = Login();
            
            // Create api services
            DocsService docsService = new(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
            
            SheetsService sheetsService = new(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            Timer timer = new ((_) =>
            {
                EventLoop(docsService, sheetsService);
            }, null, TimeSpan.Zero, TimeSpan.FromMinutes(5));

            Console.ReadLine();
        }
    }
}