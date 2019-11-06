using AdapterUtils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataAdapter
{
    public class DataFetcher
    {
        public static void FetchAndFlushData(AdapterParams prms)
        {
            // get data from excel 
            // https://stackoverflow.com/questions/37495707/how-to-find-datetime-values-in-an-epplus-excelworksheet?noredirect=1&lq=1
            // get measurement id from command line arguments. measurement id is formatted as excelFilename|selectedDataCol|selectedTimeCol
            string measId = prms.MeasId;
            // get the start and end times
            DateTime startTime = prms.FromTime;
            DateTime endTime = prms.ToTime;
            // get measurement data sources
            string[] measSegments = measId.Split('|');
            if (measSegments.Length != 4)
            {
                ConsoleUtils.FlushChunks("");
                return;
            }
            string measFilename = measSegments[0];
            string measSheetname = measSegments[1];
            string measDataCol = measSegments[2];
            string measTimeCol = measSegments[3];
            // check if file exists
            if (!File.Exists(measFilename))
            {
                ConsoleUtils.FlushChunks("");
                return;
            }
            // get excel column names
            List<string> colNames = new List<string>();
            FileInfo excelFileInfo = new FileInfo(measFilename);
            ExcelPackage package = new ExcelPackage(excelFileInfo);
            ExcelWorksheet sheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == measSheetname);
            if (sheet == null)
            {
                ConsoleUtils.FlushChunks("");
                return;
            }
            foreach (var rowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
            {
                colNames.Add(rowCell.Text);
            }
            // check if desired columns exist
            int measDataColInd = colNames.FindIndex(x => x == measDataCol) + 1;
            int measTimeColInd = colNames.FindIndex(x => x == measTimeCol) + 1;
            if (measDataColInd == 0 || measTimeColInd == 0)
            {
                ConsoleUtils.FlushChunks("");
                return;
            }
            // Create output data string
            string outStr = "";
            List<object> fetchResult = new List<object>();
            // https://stackoverflow.com/questions/37495707/how-to-find-datetime-values-in-an-epplus-excelworksheet?noredirect=1&lq=1
            for (int rowIter = 2; rowIter <= sheet.Dimension.End.Row; rowIter++)
            {
                var timeCell = sheet.Cells[rowIter, measTimeColInd];
                var valCell = sheet.Cells[rowIter, measDataColInd];
                // check if cell type is time
                if (timeCell.Value is DateTime valTime)
                {
                    // we need to convert data time to utc time before comparing
                    if ((valTime.ToUniversalTime() >= startTime) && (valTime.ToUniversalTime() <= endTime))
                    {
                        fetchResult.Add(TimeUtils.ToMillisSinceUnixEpoch(valTime));
                        // check if we have a numeric value in the data column
                        fetchResult.Add((valCell.Value is double) ? valCell.Value : null);
                    }
                }
            }
            outStr = String.Join(",", fetchResult);
            // send the output data to console
            ConsoleUtils.FlushChunks(outStr);
        }
    }
}