using Microsoft.Office.Tools.Excel;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestExcelAddIn
{
    public partial class Ribbon1
    {
        private void Setup()
        {
            try
            {
                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[2]);
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

                Excel.Range firstRow = activeWorksheet.get_Range("B2");

                firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                Excel.Range newFirstRow = activeWorksheet.get_Range("B2");

                newFirstRow.Value2 = "Incremento dias";

                Excel.Range selection = activeWorksheet.get_Range("C2");

                if (selection != null)
                {
                    incremento = new Microsoft.Office.Tools.Excel.Controls.TextBox()
                    {
                        Name = "incremento",
                        Text = defaultInc.ToString()
                    };

                    incremento.TextChanged += Incremento_TextChanged;

                    worksheet.Controls.AddControl(incremento, selection, "incremento");
                }

                Excel.Range newSecondRow = activeWorksheet.get_Range("B3");

                newSecondRow.Value2 = "Campo A";

                selection = activeWorksheet.get_Range("C3");

                if (selection != null)
                {
                    campoA = new Microsoft.Office.Tools.Excel.Controls.DateTimePicker()
                    {
                        Name = "campoA",
                        Value = DateTime.Today
                    };

                    campoA.ValueChanged += CampoA_ValueChanged;

                    worksheet.Controls.AddControl(campoA, selection, "campoA");
                }

                Excel.Range newThirdRow = activeWorksheet.get_Range("B4");

                newThirdRow.Value2 = "Campo B";

                selection = activeWorksheet.get_Range("C4");

                if (selection != null)
                {
                    campoB = new Microsoft.Office.Tools.Excel.Controls.DateTimePicker()
                    {
                        Name = "campoB",
                        Value = DateTime.Today.AddDays(defaultInc)
                    };

                    worksheet.Controls.AddControl(campoB, selection, "campoB");
                }

                selection = activeWorksheet.get_Range("C6");

                if (selection != null)
                {
                    reset = new Microsoft.Office.Tools.Excel.Controls.Button()
                    {
                        Name = "reset",
                        Text = "Reset"
                    };

                    reset.Click += Reset_Click;

                    worksheet.Controls.AddControl(reset, selection, "reset");
                }

                selection = activeWorksheet.get_Range("D2");

                if (selection != null)
                {
                    incdays = new Microsoft.Office.Tools.Excel.Controls.Button()
                    {
                        Name = "incdays",
                        Text = "Inc. Días"
                    };

                    incdays.Click += Incdays_Click;
                    incButtonColor = incdays.BackColor;

                    worksheet.Controls.AddControl(incdays, selection, "incdays");
                }
            }
            catch
            {
            }
        }

        private void Reset()
        {
            try
            {
                campoA.Value = DateTime.Today;
                incremento.Text = defaultInc.ToString();
                incdays.BackColor = incButtonColor;
            }
            catch
            {
            }
        }
    }
}
