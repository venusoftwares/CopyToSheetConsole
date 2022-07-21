using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyToSheet
{
    // the program starts here
    internal class Program
    { 
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = ConfigurationManager.AppSettings["ApplicationName"].ToString();
        static readonly string SourceSpreadsheetId = ConfigurationManager.AppSettings["SSpreadsheetId"].ToString();
        static readonly string DestinationSpreadsheetId = ConfigurationManager.AppSettings["DSpreadsheetId"].ToString();
        static readonly string sheet = ConfigurationManager.AppSettings["SourceSheet"].ToString();
        static readonly string sheet2 = ConfigurationManager.AppSettings["DestinationSheet"].ToString();
        static readonly string credentialJsonName = ConfigurationManager.AppSettings["CredentialJsonName"].ToString();
        static readonly string ignoreColumn = ConfigurationManager.AppSettings["IgnoreColumns"].ToString(); 
        static SheetsService service;

        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(credentialJsonName, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            string filterDateString = "";
            string filterCopyColumn = "y";
         

            // the date input is entered here manually
            Console.WriteLine("Enter the date from when you want to copy the records (date format: yyyy/mm/dd. eg: 2022/12/30):");
            filterDateString = Console.ReadLine();
            Console.WriteLine("------------------");

            //// the yes or no input is entered here manually
            //Console.WriteLine("Filter y or n");
            //filterCopyColumn = Console.ReadLine();
            //Console.WriteLine("------------------");

           


            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ReadAndUpdateEntries(filterDateString, filterCopyColumn);

            Console.WriteLine("Do you want to continue the process? (Y/N)");
            string re = Console.ReadLine();
            if (re == "y")
            {
                Program.Main(new string[] { });
            }

        }

        // to read and update sheet data
        static void ReadAndUpdateEntries(string filterDateString, string filterCopyColumn)
        { 
            var range = $"{sheet}!A:Z";
            var range2 = $"{sheet2}!A:Z";
            string moveYesOrNo = "";
            var valueRange2 = new ValueRange(); 
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SourceSpreadsheetId, range); 
            var response = request.Execute();


            IList<IList<object>> values = response.Values;
            List<IList<object>> recordHeader = new List<IList<object>>();
            List<IList<object>> recordRows = new List<IList<object>>();
            List<IList<object>> filterRecords = new List<IList<object>>();
            List<IList<object>> filterCopyRecords = new List<IList<object>>();


            if (values != null && values.Count > 0)
            {
                
                List<int> ignoreColumns = new List<int>();
                List<FilterModel> filterColumns = new List<FilterModel>();
                int k = 0;
                foreach (var row in values)
                {
                    if(k==0)
                    {
                        for (int i = 0; i < row.Count; i++)
                        {
                            if (ignoreColumn.Contains(row[i].ToString()))
                            {
                                ignoreColumns.Add(i);  
                            }
                           
                        }
                    }

                    List<object> list = new List<object>();
                    for (int i = 0; i < row.Count; i++)
                    {
                        if (!ignoreColumns.Any(x=>x == i))
                        {       
                             
                            list.Add(row[i]);
                        }
                    }
                    if(k==0)
                    {
                        if (list.Count > 0)
                        {
                            for(int j = 0; j < list.Count; j++)
                            {
                                if (!string.IsNullOrEmpty(filterDateString) && list[j].ToString().ToLower() == "date")
                                {
                                    filterColumns.Add(new FilterModel { index = j, indexValue = list[j].ToString() });
                                }
                                if (list[j].ToString().ToLower() == "copy")
                                {
                                    //if (filterCopyColumn.ToLower() == "y" || filterCopyColumn.ToLower() == "n")
                                    //{
                                        filterColumns.Add(new FilterModel { index = j, indexValue = list[j].ToString() });
                                    //}
                                }
                            }
                            recordHeader.Add(list);
                        }
                    }
                    else
                    {
                        if (list.Count > 0)
                        {
                            recordRows.Add(list);
                        } 
                    } 
                    k++;
                }
                if (filterColumns.Any(x => x.indexValue == "Date"))
                {
                    var itemDefault = filterColumns.Where(x => x.indexValue == "Date").Select(x => x.index).FirstOrDefault();
                    foreach (var item in recordRows)
                    {
                        DateTime inputDate = Convert.ToDateTime(filterDateString.Replace('/','-'));
                        string[] strings = item[itemDefault].ToString().Split('-');
                        DateTime dateTime = new DateTime(Convert.ToInt32(strings[2]), Convert.ToInt32(strings[1]), Convert.ToInt32(strings[0]));
                        if (strings.Length > 0)
                        {
                            if (inputDate <= dateTime)
                            {
                                filterRecords.Add(item);
                            }
                        }
                    } 
                }
                else
                {
                    filterRecords = recordRows;
                }

                if (filterColumns.Any(x => x.indexValue == "Copy"))
                {
                    var itemDefault = filterColumns.Where(x => x.indexValue == "Copy").Select(x => x.index).FirstOrDefault();
                    foreach (var item in filterRecords)
                    { 
                        if (item[itemDefault].ToString().ToLower() == filterCopyColumn.ToLower())
                        {
                            if(item[itemDefault].ToString().ToLower() == "y")
                            {
                                filterCopyRecords.Add(item);
                            } 
                        }                        
                    }
                    filterRecords = filterCopyRecords;
                }

                recordHeader.AddRange(filterRecords);

                Console.WriteLine(JsonConvert.SerializeObject(recordHeader));
                // the copyToSheet decision input is entered here manually
                Console.WriteLine($"{filterRecords.Count()} records will be moved. Are you sure to continue? (Y/N)");
                moveYesOrNo = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(moveYesOrNo))
                {
                    if (moveYesOrNo.ToLower() == "y" || moveYesOrNo.ToLower() == "yes")
                    {
                        DeleteEntry();
                        //CreateEntry();

                        valueRange2.Values = recordHeader;
                        var appendRequest = service.Spreadsheets.Values.Append(valueRange2, DestinationSpreadsheetId, range2);
                        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                        var appendReponse = appendRequest.Execute();
                    }
                }

               
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        // an alternate reference to update data 
        static void UpdateEntry()
        {
            var range = $"{sheet2}!A:Z";
            var valueRange = new ValueRange();

            IList<IList<Object>> list = new List<IList<Object>>() { };
            valueRange.Values = list;

            //var oblist = new List<object>() { };
            //valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, DestinationSpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        // to delete the old data in the sheet
        static void DeleteEntry()
        {
            var range = $"{sheet2}!A:Z";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, DestinationSpreadsheetId, range);
            var deleteReponse = deleteRequest.Execute();
        }

        // model
        public class FilterModel
        {
            public int index { get; set; }
            public string indexValue { get; set; }
        }
    }
}
