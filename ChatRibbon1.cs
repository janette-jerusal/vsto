using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Test2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        public void OnCompareBtnClick(Office.IRibbonControl control)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (dialog.ShowDialog() == DialogResult.OK && dialog.FileNames.Length == 2)
            {
                var file1 = dialog.FileNames[0];
                var file2 = dialog.FileNames[1];

                var stories1 = ReadStoriesFromExcel(file1);
                var stories2 = ReadStoriesFromExcel(file2);

                var results = CompareDescriptions(stories1, stories2);

                OutputResults(results);
            }
        }

        private Dictionary<string, string> ReadStoriesFromExcel(string path)
        {
            var stories = new Dictionary<string, string>();

            var app = new Excel.Application();
            var workbook = app.Workbooks.Open(path);
            var sheet = workbook.Sheets[1] as Excel.Worksheet;
            var range = sheet.UsedRange;

            for (int i = 2; i <= range.Rows.Count; i++)
            {
                string id = Convert.ToString((range.Cells[i, 1] as Excel.Range)?.Value2);
                string desc = Convert.ToString((range.Cells[i, 2] as Excel.Range)?.Value2);
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(desc))
                    stories[id] = desc;
            }

            workbook.Close(false);
            app.Quit();

            return stories;
        }

        private List<(string ID1, string ID2, double Score)> CompareDescriptions(Dictionary<string, string> stories1, Dictionary<string, string> stories2)
        {
            var allDocs = stories1.Values.Concat(stories2.Values).ToList();
            var tfidfVectors = allDocs.Select(doc => GetTfIdfVector(doc, allDocs)).ToList();

            int split = stories1.Count;
            var results = new List<(string, string, double)>();

            int i = 0;
            foreach (var kvp1 in stories1)
            {
                int j = split;
                foreach (var kvp2 in stories2)
                {
                    double sim = CosineSimilarity(tfidfVectors[i], tfidfVectors[j]);
                    results.Add((kvp1.Key, kvp2.Key, sim));
                    j++;
                }
                i++;
            }

            return results.OrderByDescending(r => r.Score).ToList();
        }

        private Dictionary<string, double> GetTfIdfVector(string doc, List<string> allDocs)
        {
            var words = Tokenize(doc);
            var tf = words.GroupBy(w => w).ToDictionary(g => g.Key, g => (double)g.Count() / words.Count);
            var idf = words.Distinct().ToDictionary(
                w => w,
                w => Math.Log((double)allDocs.Count / allDocs.Count(d => Tokenize(d).Contains(w)))
            );
            return tf.ToDictionary(kv => kv.Key, kv => kv.Value * idf[kv.Key]);
        }

        private List<string> Tokenize(string text)
        {
            return Regex.Split(text.ToLower(), @"\W+").Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        }

        private double CosineSimilarity(Dictionary<string, double> vec1, Dictionary<string, double> vec2)
        {
            var allKeys = vec1.Keys.Union(vec2.Keys);
            double dot = allKeys.Sum(k => vec1.GetValueOrDefault(k) * vec2.GetValueOrDefault(k));
            double mag1 = Math.Sqrt(vec1.Values.Sum(v => v * v));
            double mag2 = Math.Sqrt(vec2.Values.Sum(v => v * v));
            return mag1 > 0 && mag2 > 0 ? dot / (mag1 * mag2) : 0.0;
        }

        private void OutputResults(List<(string ID1, string ID2, double Score)> results)
        {
            Excel.Worksheet newSheet = Globals.ThisWorkbook.Sheets.Add();
            newSheet.Name = "Similarity Results";

            newSheet.Cells[1, 1].Value2 = "ID1";
            newSheet.Cells[1, 2].Value2 = "ID2";
            newSheet.Cells[1, 3].Value2 = "Similarity";

            int row = 2;
            foreach (var (id1, id2, score) in results)
            {
                newSheet.Cells[row, 1].Value2 = id1;
                newSheet.Cells[row, 2].Value2 = id2;
                newSheet.Cells[row, 3].Value2 = Math.Round(score, 4);
                row++;
            }
        }
    }
}
