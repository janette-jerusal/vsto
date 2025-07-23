using System;
using System.Windows.Forms;

namespace Test2
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, EventArgs e)
        {
            // You can put workbook startup logic here if needed.
        }

        private void ThisWorkbook_Shutdown(object sender, EventArgs e)
        {
            // Cleanup logic, if needed.
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

