using System;
using System.Windows.Forms;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.IO;

namespace googleExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // Nuget: google.Apis.Sheets.v4
        }

        private void Read_Click(object sender, EventArgs e)
        {
            //從客戶端密鑰檔案中載入憑證：
            GoogleCredential credential;
            using (var stream = new FileStream(@"./mimetic-plate-392915-66bde0131431.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
                string sheetName = "Sheet1";
                //建立SheetsService實例：
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = sheetName
                });

                //讀取指定的Google Sheets資料：
                // 指定要讀取的Google Sheets的ID和範圍（例如，Sheet1中的A1:C10儲存格範圍）
                // spreadsheet_id = 網址中，spreadsheets後面
                string spreadsheet_id = "1yejADfPXru_r9qJzGSQvrr7t9lrlouGdmGvk-vaYnYo";
                var spreadsheetId = spreadsheet_id;

                // 呼叫API進行資料讀取
                // 處理讀取的資料
                var lastColumn = "Z";
                var range = $"{sheetName}!A1:{lastColumn}";

                var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        foreach (var cell in row)
                        {
                            Console.WriteLine(cell);
                        }
                    }
                }
            }
        }
    }
}
