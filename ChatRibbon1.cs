using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace Test2
{
    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var resourceName = "Test2.Ribbon1.xml";
            using (Stream stream = GetType().Assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI) => ribbon = ribbonUI;

        public void OnCompare(IRibbonControl control)
        {
            // Open file dialogs for both Excel files
            OpenFileDialog dialog1 = new OpenFileDialog();
            dialog1.Filter = "Excel Files|*.xlsx";
            dialog1.Title = "Select First Excel File";

            OpenFileDialog dialog2 = new OpenFileDialog();
            dialog2.Filter = "Excel Files|*.xlsx";
            dialog2.Title = "Select Second Excel File";

            if (dialog1.ShowDialog() != DialogResult.OK || dialog2.ShowDialog() != DialogResult.OK)
                return;

            var wb1 = Globals.ThisWorkbook.Application.Workbooks.Open(dialog1.FileName);
            var wb2 = Globals.ThisWorkbook.Application.Workbooks.Open(dialog2.FileName);

            var ws1 = wb1.Sheets[1] as Worksheet;
            var ws2 = wb2.Sheets[1] as Worksheet;

            var data1 = ReadData(ws1);
            var data2 = ReadData(ws2);

            var results = CompareStories(data1, data2, 0.7); // 70% similarity threshold

            var outputWS = Globals.ThisWorkbook.Sheets.Add() as Worksheet;
            outputWS.Name = "Comparison Results";
            outputWS.Cells[1, 1].Value = "ID 1";
            outputWS.Cells[1, 2].Value = "ID 2";
            outputWS.Cells[1, 3].Value = "Similarity";

            int row = 2;
            foreach (var res in results)
            {
                outputWS.Cells[row, 1].Value = res.Item1;
                outputWS.Cells[row, 2].Value = res.Item2;
                outputWS.Cells[row, 3].Value = res.Item3;
                row++;
            }

            wb1.Close(false);
            wb2.Close(false);
        }

        private Dictionary<string, string> ReadData(Worksheet ws)
        {
            Dictionary<string, string> stories = new Dictionary<string, string>();
            int row = 2;

            while (true)
            {
                var id = ws.Cells[row, 1].Value?.ToString();
                var desc = ws.Cells[row, 2].Value?.ToString();

                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(desc))
                    break;

                stories[id] = desc;
                row++;
            }

            return stories;
        }

        private List<Tuple<string, string, double>> CompareStories(Dictionary<string, string> a, Dictionary<string, string> b, double threshold)
        {
            var results = new List<Tuple<string, string, double>>();

            foreach (var pairA in a)
            {
                foreach (var pairB in b)
                {
                    double sim = CosineSimilarity(pairA.Value, pairB.Value);
                    if (sim >= threshold)
                        results.Add(Tuple.Create(pairA.Key, pairB.Key, sim));
                }
            }

            return results;
        }

        private double CosineSimilarity(string s1, string s2)
        {
            var vec1 = GetVector(s1);
            var vec2 = GetVector(s2);

            var allKeys = vec1.Keys.Union(vec2.Keys);
            double dot = 0, mag1 = 0, mag2 = 0;

            foreach (var key in allKeys)
            {
                double v1 = vec1.ContainsKey(key) ? vec1[key] : 0;
                double v2 = vec2.ContainsKey(key) ? vec2[key] : 0;

                dot += v1 * v2;
                mag1 += v1 * v1;
                mag2 += v2 * v2;
            }

            return (mag1 == 0 || mag2 == 0) ? 0 : dot / (Math.Sqrt(mag1) * Math.Sqrt(mag2));
        }

        private Dictionary<string, int> GetVector(string text)
        {
            var words = text.ToLower().Split(new[] { ' ', ',', '.', ';', ':', '-', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            var dict = new Dictionary<string, int>();

            foreach (var word in words)
            {
                if (!dict.ContainsKey(word)) dict[word] = 0;
                dict[word]++;
            }

            return dict;
        }
    }
}
