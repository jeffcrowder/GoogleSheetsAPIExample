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
        static volatile bool exit = false;

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
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

            // Define request parameters.
            String spreadsheetId = "1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So";
            String readRange = "ReadData!A1:F";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, readRange);

            // Prints the data from a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So/edit#gid=0
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A through E, which correspond to indices 0 through 5.
                    //Need to seperate the headers from the data
                    Console.WriteLine("{0},         {1},    {2},    {3},    {4},    {5}", row[0], row[1], row[2], row[3], row[4], row[5]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            //Write some data
            String writeRange = "WriteData!A1:C3";
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "ROWS";//"ROWS";//"COLUMNS"

           // here is the actual data to be sent to sheet
            var headerList = new List<object>() { "CAL", "FPY", "AY" };

            var dataList = new List<object>() { 1, 2, 3 };

            valueRange.Values = new List<IList<object>> { headerList, dataList };
            
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, writeRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result;
            //result = update.Execute();


            //start the readKey in a new thread 
            Task.Factory.StartNew(() =>
            {
                while (Console.ReadKey(true).Key != ConsoleKey.Enter) ;
                exit = true;
            });

            Console.Write("Press <Enter> to exit...\r\n");
            int x = Console.CursorLeft;
            int y = Console.CursorTop;
            Random r = new Random();

            //Loop and update the data until enter key is hit to exit
            do
            {
                //clear the data
                dataList.RemoveRange(0, 3);

                //refresh the data
                dataList.Add(r.Next(80,100));
                dataList.Add(r.Next(75,99));
                dataList.Add(r.Next(90,99));

                //update data via the API
                result = update.Execute();

                Console.Write("Updated with: " + dataList.ElementAt(0) + " : " + dataList.ElementAt(1) + " : " + dataList.ElementAt(2));

                // Restore previous position
                Console.SetCursorPosition(x, y);

                //wait
                System.Threading.Thread.Sleep(3000);
            } while (!exit);




            SqlConnection sqlConnection = new SqlConnection("Data Source=tul-mssql;Initial Catalog=Division;User ID=tqisadmin;Password=admin2k");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
            //cmd.CommandText = "select top 5 [linenbr] from [Division].[dbo].[asyTestRecords]";
                       cmd.CommandText = "Select LineNbr,passfail,cast(count(*) * 100.0 / sum(count(*))over(partition BY linenbr) as decimal (10, 2)) pctTotals " +
            "FROM[Division].[dbo].[asyTestRecords]" +
            "WHERE Computer = 'LN'and shiftNbr = 1 and datepart(DAYOFYEAR, [DTShiftStart]) = DATEPART(DAyofyear, current_timestamp)" +
            "group BY linenbr,shiftnbr,passfail order bY LineNbr, passfail desc";

            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection;

            sqlConnection.Open();

            reader = cmd.ExecuteReader();

            // Data is accessible through the DataReader object here.
            if (reader.HasRows)
            {
                Console.Write("\r\n");
                int cnt = 0;
                while (reader.Read())
                {
                    Console.WriteLine("\t{0}\t{1}\t{2}", reader.GetSqlInt16(0).ToSqlString(), reader.GetSqlString(1), reader.GetSqlDecimal(2).ToSqlString());
                    //reader.GetString(1));
                    cnt++;
                    Console.Write(cnt);
                }
                System.Threading.Thread.Sleep(3000);
            }
            else
            {
                Console.WriteLine("No rows found.");
            }
            reader.Close();

            sqlConnection.Close();
        }
    }
}
