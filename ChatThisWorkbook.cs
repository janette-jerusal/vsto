using System;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test2
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            // No need to return or construct the Ribbon here
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Optional: Method to be called by Ribbon1.cs
        public void Compare()
        {
            System.Windows.Forms.MessageBox.Show("Comparison logic triggered.");
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion
    }
}
