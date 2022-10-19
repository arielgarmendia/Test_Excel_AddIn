using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;

namespace Pricer.ExcelAddIn
{
    public partial class Ribbon
    {
        private const int defaultInc = 5;

        private Microsoft.Office.Tools.Excel.Controls.TextBox incremento;

        private Microsoft.Office.Tools.Excel.Controls.DateTimePicker campoA;
        private Microsoft.Office.Tools.Excel.Controls.DateTimePicker campoB;

        private Microsoft.Office.Tools.Excel.Controls.Button reset;
        private Microsoft.Office.Tools.Excel.Controls.Button incdays;

        private Color incButtonColor;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void RunButton_Click(object sender, RibbonControlEventArgs e)
        {
            Setup();
        }

        private void Incremento_TextChanged(object sender, EventArgs e)
        {
            sumDays();
        }

        private void CampoA_ValueChanged(object sender, EventArgs e)
        {
            sumDays();
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private void Incdays_Click(object sender, EventArgs e)
        {
            try
            {
                incremento.Text = (Convert.ToInt32(incremento.Text) + 1).ToString();

                if (Convert.ToInt32(incremento.Text) != defaultInc)
                    incdays.BackColor = Color.Red;
                else
                    incdays.BackColor = incButtonColor;
            }
            catch
            {
                incremento.Text = defaultInc.ToString();
            }
        }

        private void sumDays()
        {
            try
            {
                campoB.Value = campoA.Value.AddDays(Convert.ToInt32(incremento.Text));

                if (Convert.ToInt32(incremento.Text) != defaultInc)
                {
                    incremento.BackColor = Color.Yellow;
                    incdays.BackColor = Color.Red;
                }
                else
                {
                    incremento.BackColor = Color.White;
                    incdays.BackColor = incButtonColor;
                }
            }
            catch
            {
                incremento.Text = defaultInc.ToString();
            }
        }
    }
}
