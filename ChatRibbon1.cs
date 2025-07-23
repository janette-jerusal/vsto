using System;
using Office = Microsoft.Office.Core;

namespace Test2
{
    public partial class Ribbon1 : Office.IRibbonExtensibility
    {
        public Ribbon1() { }

        public string GetCustomUI(string ribbonID)
        {
            return Properties.Resources.Ribbon1;
        }

        public void OnComparePressed(Office.IRibbonControl control)
        {
            Globals.ThisWorkbook.Compare();
        }
    }
}

