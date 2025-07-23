using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace Test2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        private void CompareButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Compare();
        }
    }
}
