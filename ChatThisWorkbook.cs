using System;
using Microsoft.Office.Tools.Excel;

namespace Test2
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, EventArgs e) { }

        private void ThisWorkbook_Shutdown(object sender, EventArgs e) { }

        public void Compare()
        {
            System.Windows.Forms.MessageBox.Show("Comparison logic triggered!");
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new EventHandler(ThisWorkbook_Shutdown);
        }
        #endregion
    }
}

