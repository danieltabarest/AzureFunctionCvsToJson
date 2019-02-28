using System.IO;
using System.Linq;
using CsvHelper;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;

using System.Net;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.WebJobs;
using System.Linq;


namespace CSVConverter
{
    public static class JSONConverter
    {
       
       [FunctionName("JSONConverter")]
       public static void Run([BlobTrigger("to-convert/{name}")]
            Stream myBlob, string name,
            TraceWriter log)
        {
            log.Info($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");

            //Only convert CSV files
            if (name.Contains(".xlsx") || name.Contains(".csv"))
            {
                //var json = ConvertCsvToJson(myBlob);
                ConvertExcelToJson(myBlob,log);
                //log.Info(json);
            }
            else
            {
                log.Info("Not a CSV");
            }           
        }

        public static string ConvertCsvToJson(Stream blob)
        {
            var sReader = new StreamReader(blob);
            var csv = new CsvReader(sReader);

            csv.Read();
            csv.ReadHeader();

            var csvRecords = csv.GetRecords<object>().ToList();
           
            return JsonConvert.SerializeObject(csvRecords);
        }

        public static string ConvertExcelToJson(Stream blob, TraceWriter log)
        { 
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(blob, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet sheet = worksheetPart.Worksheet;

                var cells = sheet.Descendants<Cell>();
                var rows = sheet.Descendants<Row>();

                log.Info(string.Format("Row count = {0}", rows.LongCount()));
                log.Info(string.Format("Cell count = {0}", cells.LongCount()));

                // One way: go through each cell in the sheet
                foreach (Cell cell in cells)
                {
                    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                    {
                        int ssid = int.Parse(cell.CellValue.Text);
                        string str = sst.ChildElements[ssid].InnerText;
                        log.Info(string.Format("Shared string {0}: {1}", ssid, str));
                    }
                    else if (cell.CellValue != null)
                    {
                        log.Info(string.Format("Cell contents: {0}", cell.CellValue.Text));
                    }
                }

                return null;
            }

        }
    }
}