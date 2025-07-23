using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace Test2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        private void btnCompare_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog ofd1 = new OpenFileDialog();
            ofd1.Title = "Select first Excel file";
            ofd1.Filter = "Excel Files|*.xls;*.xlsx";

            if (ofd1.ShowDialog() != DialogResult.OK) return;
            string file1 = ofd1.FileName;

            OpenFileDialog ofd2 = new OpenFileDialog();
            ofd2.Title = "Select second Excel file";
            ofd2.Filter = "Excel Files|*.xls;*.xlsx";

            if (ofd2.ShowDialog() != DialogResult.OK) return;
            string file2 = ofd2.FileName;

            var stories1 = ReadUserStoriesFromExcel(file1);
            var stories2 = ReadUserStoriesFromExcel(file2);

            var results = CompareUserStories(stories1, stories2);
            WriteResultsToActiveWorkbook(results);
        }

        private Dictionary<string, string> ReadUserStoriesFromExcel(string path)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(path);
            var ws = (Worksheet)wb.Sheets[1];

            var result = new Dictionary<string, string>();
            int row = 2;
            while (true)
            {
                var idCell = ws.Cells[row, 1]?.Value;
                var descCell = ws.Cells[row, 2]?.Value;
                if (idCell == null || descCell == null)
                    break;

                string id = idCell.ToString();
                string description = descCell.ToString();
                result[id] = description;
                row++;
            }

            wb.Close(false);
            app.Quit();
            return result;
        }

        private List<(string ID1, string ID2, double Similarity)> CompareUserStories(Dictionary<string, string> set1, Dictionary<string, string> set2)
        {
            var results = new List<(string, string, double)>();

            foreach (var (id1, desc1) in set1)
            {
                foreach (var (id2, desc2) in set2)
                {
                    double score = CosineSimilarity(desc1, desc2);
                    results.Add((id1, id2, score));
                }
            }

            return results.OrderByDescending(r => r.Similarity).ToList();
        }

        private double CosineSimilarity(string text1, string text2)
        {
            var tf1 = GetTermFrequencies(text1);
            var tf2 = GetTermFrequencies(text2);
            var allTerms = tf1.Keys.Union(tf2.Keys).ToHashSet();

            double dot = allTerms.Sum(term => tf1.GetValueOrDefault(term) * tf2.GetValueOrDefault(term));
            double mag1 = Math.Sqrt(tf1.Values.Sum(v => v * v));
            double mag2 = Math.Sqrt(tf2.Values.Sum(v => v * v));

            return (mag1 == 0 || mag2 == 0) ? 0 : dot / (mag1 * mag2);
        }

        private Dictionary<string, double> GetTermFrequencies(string text)
        {
            var tf = new Dictionary<string, double>();
            string[] words = Regex.Split(text.ToLower(), @"\W+");

            foreach (var word in words.Where(w => !string.IsNullOrWhiteSpace(w)))
            {
                if (!tf.ContainsKey(word)) tf[word] = 0;
                tf[word]++;
            }

            return tf;
        }

        private void WriteResultsToActiveWorkbook(List<(string ID1, string ID2, double Similarity)> results)
        {
            var excel = Globals.ThisWorkbook.Application;
            Worksheet newSheet = excel.Worksheets.Add();
            newSheet.Name = "Similarity Results";

            newSheet.Cells[1, 1].Value = "ID 1";
            newSheet.Cells[1, 2].Value = "ID 2";
            newSheet.Cells[1, 3].Value = "Similarity";

            int row = 2;
            foreach (var (id1, id2, sim) in results)
            {
                newSheet.Cells[row, 1].Value = id1;
                newSheet.Cells[row, 2].Value = id2;
                newSheet.Cells[row, 3].Value = Math.Round(sim, 4);
                row++;
            }
        }
    }
}
