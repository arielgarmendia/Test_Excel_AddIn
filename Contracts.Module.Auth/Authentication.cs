using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contracts.Module.Auth
{
    public class Authentication
    {
        public static string InternalLogIn(string user, string password)
        {
            string InternalLogInRet = default;

            InternalLogInRet = InternalLogInBis("");

            return InternalLogInRet;
        }

        private static string InternalLogInBis(string action)
        {
            string InternalLogInBisRet = default;
            var ObjetoHTTP = default(object);
            string Peticion;
            string Respuesta;
            string pricerEnvironment = "";

            //pricerEnvironment = ThisWorkbook.Sheets("LVB Proced. Generator").Range("pricerEnvironment").Value;

            switch (true)
            {
                case object _ when pricerEnvironment == "PREproduction":
                    {
                        if (string.IsNullOrEmpty(action))
                        {
                            InternalLogInBisRet = "https://au-wbamdesktop.es.igrupobbva/";
                        }
                        else
                        {
                            InternalLogInBisRet = "https://ei-wbamdesktop.es.igrupobbva/";
                        }

                        break;
                    }
                case object _ when pricerEnvironment == "PROduction":
                    {
                        InternalLogInBisRet = "https://cibdesktop.es.igrupobbva/";
                        break;
                    }
                case object _ when pricerEnvironment == "WindowsLocal":
                    {
                        InternalLogInBisRet = "https://au-wbamdesktop.es.igrupobbva/";
                        break;
                    }
                case object _ when pricerEnvironment.Contains("CPrice"):
                    {
                        if (string.IsNullOrEmpty(action))
                        {
                            InternalLogInBisRet = "https://au-wbamdesktop.es.igrupobbva/";
                        }
                        else
                        {
                            InternalLogInBisRet = "https://ei-wbamdesktop.es.igrupobbva/";
                        }

                        break;
                    }

                default:
                    {
                        //Unload(LogIn);
                        //Galleta = "";

                        //ResetButtons("LVB Proced. Generator");
                        //ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0); // 
                        //ThisWorkbook.Worksheets("LVB Proced. Generator").Range("pricerEnvironment").Activate();

                        //Interaction.MsgBox("ERROR: Pricer server not configured property!!", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR");

                        return InternalLogInBisRet;
                    }
            };

            //ObjetoHTTP.Open("GET", InternalLogInBisRet + "pkmslogout", (object)false);
            //ObjetoHTTP.SetRequestHeader("Content-Type", "text/html");

            //ObjetoHTTP.Send();

            //if (Information.Err().Number == -2147024891 | Information.Err().Number == -2147012744)
            //{
            //    ObjetoHTTP.Open("GET", InternalLogInBisRet, (object)false);
            //    ObjetoHTTP.SetRequestHeader("Content-Type", "text/html");
            //    ObjetoHTTP.Send();
            //};

            //#error Cannot convert OnErrorGoToStatementSyntax - see comment for details
            ///* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 4331


            //Input:
            //            On Error GoTo 0

            // */
            //;
            //Peticion = "username=" + sUser + "&password=" + sPassword + "&login-form-type=pwd";
            //ObjetoHTTP.Open("POST", InternalLogInBisRet + "pkmslogin.form", (object)false);
            //ObjetoHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded");
            //ObjetoHTTP.Send(Peticion);

            //if (Information.Err().Number == -2147024891)
            //{
            //    ObjetoHTTP.Open("POST", InternalLogInBisRet, (object)false);
            //    ObjetoHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded");
            //    ObjetoHTTP.Send(Peticion);
            //}

            //Respuesta = Conversions.ToString(ObjetoHTTP.responseText);

            //if (Strings.InStr(1, Respuesta, "page: login_success.html") > 0)
            //{
            //    Galleta = ObjetoHTTP.getResponseHeader("Set-Cookie");
            //}

            return InternalLogInBisRet;
        }
    }
}
