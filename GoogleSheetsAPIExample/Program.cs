using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        //static volatile bool exit = false;

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to:\r\n" + credPath + "\r\n");
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Prints the data from a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So/edit#gid=0

            // Define request parameters.
            String spreadsheetId = "1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So";

            // here is the actual data to be sent to sheet
            List<object> headerList = new List<object>() {
            "ID","DTStamp","DTShiftStart","ModelNbr","SerialNbr","PassFail","LineNbr","ShiftNbr","Computer","Word40","Word41","Word42"
            ,"Word43","Word44","Word45","Word46","Word47","Word48","Word49","Word50","Word51","Word52","Word53","Word54","Word55","Word56"
            ,"Word57","Word58","Word59","Word60","Word61","Word62","Word63","Word64","Word65","Word66","Word67","Word68","Word69","Word70"
            ,"Word71","Word72","Word73","Word74","Word75","Word76","Word77","Word78","Word79","Word80"};

            //var dataList = new List<object>();

            //Write some data
            String writeRange = "WriteData!A1:ZZ";
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "ROWS";

            Console.WriteLine("Clear the Sheet");

            //API method to clear the sheet
            ClearValuesRequest clearValuesRequest = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest cr = service.Spreadsheets.Values.Clear(clearValuesRequest, spreadsheetId, writeRange);
            ClearValuesResponse clearResponse = cr.Execute();

            Console.WriteLine("Delete all rows in Sheet");

            //API method to batch update
            DimensionRange dr = new DimensionRange
            {
                Dimension = "ROWS",
                StartIndex = 1,
                SheetId = 1809337217
            };

            DeleteDimensionRequest ddr = new DeleteDimensionRequest();
            ddr.Range = dr;

            Request r = new Request();
            r.DeleteDimension = ddr;

            List<Request> batchRequests = new List<Request>();// { "requests": [{ "deleteDimension": { "range": { "sheetId": 1809337217, "startIndex": 1}} }  ]};
            batchRequests.Add(r);

            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = batchRequests;

            SpreadsheetsResource.BatchUpdateRequest bRequest = service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
            BatchUpdateSpreadsheetResponse busr = bRequest.Execute();

            Console.WriteLine("Write the Headers to the Sheet");

            //API method to update the sheet
            valueRange.Values = new List<IList<object>> { headerList };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, writeRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result;
            result = update.Execute();

            SqlConnection sqlConnection = new SqlConnection("Data Source=tul-mssql;Initial Catalog=Division;User ID=tqisadmin;Password=admin2k");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
  
            cmd.CommandText = "SELECT TOP 1000 [ID],[DTStamp],[DTShiftStart],[ModelNbr],[SerialNbr],[PassFail],[LineNbr],[ShiftNbr],[Computer],[Word40],[Word41],[Word42]" +
            ",[Word43],[Word44],[Word45],[Word46],[Word47],[Word48],[Word49],[Word50],[Word51],[Word52],[Word53],[Word54],[Word55],[Word56]" +
            ",[Word57],[Word58],[Word59],[Word60],[Word61],[Word62],[Word63],[Word64],[Word65],[Word66],[Word67],[Word68],[Word69],[Word70]" +
            ",[Word71],[Word72],[Word73],[Word74],[Word75],[Word76],[Word77],[Word78],[Word79],[Word80] " +
            "FROM[Division].[dbo].[asyTestRecords] where LineNbr = 2 and computer = 'LN' and dtstamp > '2/1/2018 5:00' order by dtstamp desc";

            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection;

            Console.WriteLine("Open the SQL connection");
            sqlConnection.Open();
            cmd.CommandTimeout = 60;
            Console.WriteLine("Please wait while reading data from SQL");
            reader = cmd.ExecuteReader();

            // Data is accessible through the DataReader object here.
            ValueRange valueDataRange = new ValueRange();
            valueDataRange.MajorDimension = "ROWS";

            var dataList = new List<object>();
            valueDataRange.Values = new List<IList<object>> { dataList };

            //API to append data to sheet
            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueDataRange, spreadsheetId, writeRange);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            if (reader.HasRows)
            {
                //Console.WriteLine("{0}",reader.FieldCount);
                Object[] colValues = new Object[reader.FieldCount];

                int throttleCount = 0;
                int cnt = 0;
                while (reader.Read())
                {
                    //This logic is flawed. If we get hit by the quota then the data row gets lost the next time this runs.
                    dataList.Clear();
                    for (int i = 0; i < reader.GetValues(colValues); i++)
                    {
                        dataList.Add(colValues[i]);
                    }

                    try
                    {
                        //This is the GOOGLE query Throttle they only allow 500 writes per 100 sec
                        System.Threading.Thread.Sleep(20);
                        AppendValuesResponse appendValueResponse = appendRequest.Execute();
                        Console.WriteLine("Writing to Sheet: row{0}", cnt);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine("Whoa buddy slowdown {0}",throttleCount);
                        System.Threading.Thread.Sleep(3000);
                        throttleCount++;
                    }
                    cnt++;
                }
            }
            else
            {
                Console.WriteLine("No rows found.");
            }
            
            // sit here and wait for a while
            Console.WriteLine("Done waiting to close reader and SQL");
            System.Threading.Thread.Sleep(3000);

            reader.Close();
            sqlConnection.Close();
        }
    }
}
