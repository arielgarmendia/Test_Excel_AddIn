﻿using Excel = Microsoft.Office.Interop.Excel;

namespace Pricer.ExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += Application_WorkbookOpen;
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            Wb.Sheets.Add(After: Wb.Sheets[Wb.Sheets.Count]);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //this is a test
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
