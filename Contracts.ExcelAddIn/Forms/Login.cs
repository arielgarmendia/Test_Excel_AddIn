using Contracts.Module.Auth;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pricer.ExcelAddIn.Forms
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void LoginButton_Click(object sender, EventArgs e)
        {
            try
            {
                var valid = ValidateInputs();

                if (valid)
                {
                    //{
                    //    var withBlock = ThisWorkbook.Sheets("LVB Proced. Generator");
                    //    withBlock.Range("Underlying_L1").Value = withBlock.Range("Underlying_L1").Value;
                    //}

                    var Galleta = Authentication.InternalLogIn(this.User.Text, this.Password.Text);

                    if (Galleta != "")
                    {
                        //ResetButtons("LVB Proced. Generator");
                        //ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogin.BackColor = VBA.RGB(0, 255, 0);
                        //Interaction.MsgBox("Login OK.", Constants.vbInformation, "OK");

                        var iRespuesta = MessageBox.Show("Underlyings and clients will then be loaded. Do you want to load them?", "LOAD STATIC DATA", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                        if (iRespuesta == DialogResult.Yes)
                        {
                            //ButtonLoadStaticData("LVB Proced. Generator");
                        }

                        this.DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        //ResetButtons("LVB Proced. Generator");
                        //ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0);

                        MessageBox.Show("Login KO.", "KO", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        this.DialogResult = DialogResult.Retry;
                    }
                }
                else
                    this.DialogResult = DialogResult.Retry;
            }
            catch
            {
                MessageBox.Show("Login KO.", "KO", MessageBoxButtons.OK, MessageBoxIcon.Error);

                this.DialogResult = DialogResult.Retry;
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private bool ValidateInputs()
        {
            if (User.Text != "" && Password.Text != "")
                return true;
            else
            {
                if (User.Text == "")
                    UserErrorProvider.SetError(User, "You must provide a User name.");
                else
                    UserErrorProvider.SetError(User, "");

                if (Password.Text == "")
                    PasswordErrorProvider.SetError(Password, "You must provide a Password.");
                else
                    PasswordErrorProvider.SetError(Password, "");

                return false;
            }
        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = this.DialogResult == DialogResult.Retry;
        }
    }
}
