using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Test2
{
    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var resourceName = "Test2.Ribbon1.xml";
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void OnCompareButtonClick(IRibbonControl control)
        {
            string filePath1 = @"C:\Users\<your_username>\Documents\UserStories1.xlsx";
            string filePath2 = @"C:\Users\<your_username>\Documents\UserStories2.xlsx";

            var app = Globals.ThisWorkbook.Application;
            Excel.Workbook wb1 = app.Workbooks.Open(filePath1);
            Excel.Workbook wb2 = app.Workbooks.Open(filePath2);

            Excel.Worksheet sheet1 = wb1.Sheets[1];
            Excel.Worksheet sheet2 = wb2.Sheets[1];

            Excel.Range data1 = sheet1.UsedRange;
            Excel.Range data2 = sheet2.UsedRange;

            // Simple similarity logic: match identical descriptions
            var matches = from r1 in data1.Rows.Cast<Excel.Range>()
                          from r2 in data2.Rows.Cast<Excel.Range>()
                          let desc1 = r1.Cells[1, 2].Text.ToString()
                          let desc2 = r2.Cells[1, 2].Text.ToString()
                          where !string.IsNullOrWhiteSpace(desc1) && desc1 == desc2
                          select new { ID1 = r1.Cells[1, 1].Text, ID2 = r2.Cells[1, 1].Text };

            Excel.Worksheet resultSheet = Globals.ThisWorkbook.Sheets.Add();
            resultSheet.Name = "ComparisonResults";
            resultSheet.Cells[1, 1] = "ID in File 1";
            resultSheet.Cells[1, 2] = "ID in File 2";

            int row = 2;
            foreach (var match in matches)
            {
                resultSheet.Cells[row, 1] = match.ID1;
                resultSheet.Cells[row, 2] = match.ID2;
                row++;
            }

            wb1.Close(false);
            wb2.Close(false);
        }
    }
}

