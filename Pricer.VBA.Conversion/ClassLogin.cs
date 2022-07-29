public class ClassLogin
{
    public bool bLogin;
    public string sUser;
    public string sPassword;

    private void OK_Click()
    {
        object ObjetoHTTP;
        string Peticion;
        string Respuesta;
        int iRespuesta;
        ;

        // para el error #value!
        {
            var withBlock = ThisWorkbook.Sheets("LVB Proced. Generator");
            withBlock.Range("Underlying_L1").Value = withBlock.Range("Underlying_L1").Value;
        }

        bLogin = false;
        sUser = UserName.Value;
        sPassword = Password.Value;

        Unload(LogIn);

        var Galleta = InternalLogIn();

        if (Galleta != "")
        {
            ResetButtons("LVB Proced. Generator");
            ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogin.BackColor = VBA.RGB(0, 255, 0);
            Interaction.MsgBox("Login OK.", Constants.vbInformation, "OK");

            iRespuesta = (int)Interaction.MsgBox("Underlyings and clients will then be loaded. Do you want to load them?", (MsgBoxStyle)((int)Constants.vbInformation + (int)Constants.vbYesNo), "LOAD STATIC DATA");
            if (iRespuesta == (int)Constants.vbYes)
            {
                ButtonLoadStaticData("LVB Proced. Generator");
            }
        }
        else
        {
            ResetButtons("LVB Proced. Generator");
            ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0);
            Interaction.MsgBox("Login KO.", Constants.vbCritical, "KO");
        }
    }

    public string InternalLogIn()
    {
        string InternalLogInRet = default;
        InternalLogInRet = InternalLogInBis("");
        return InternalLogInRet;
    }

    public string InternalLogInBis(string action)
    {
        string InternalLogInBisRet = default;
        var ObjetoHTTP = default(object);
        string Peticion;
        string Respuesta;
        string pricerEnvironment;

        pricerEnvironment = ThisWorkbook.Sheets("LVB Proced. Generator").Range("pricerEnvironment").Value;

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
            case object _ when Strings.InStr(pricerEnvironment, "CPrice") == 1:
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
                    Unload(LogIn);
                    Galleta = "";
                    ResetButtons("LVB Proced. Generator");
                    ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0); // 
                    ThisWorkbook.Worksheets("LVB Proced. Generator").Range("pricerEnvironment").Activate();
                    Interaction.MsgBox("ERROR: Pricer server not configured property!!", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR");
                    return InternalLogInBisRet;
                }
        };

        ObjetoHTTP.Open("GET", InternalLogInBisRet + "pkmslogout", (object)false);
        ObjetoHTTP.SetRequestHeader("Content-Type", "text/html");

        ObjetoHTTP.Send();

        if (Information.Err().Number == -2147024891 | Information.Err().Number == -2147012744)
        {
            ObjetoHTTP.Open("GET", InternalLogInBisRet, (object)false);
            ObjetoHTTP.SetRequestHeader("Content-Type", "text/html");
            ObjetoHTTP.Send();
        };

        #error Cannot convert OnErrorGoToStatementSyntax - see comment for details
        /* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 4331


        Input:
                    On Error GoTo 0

         */
        ;
        Peticion = "username=" + sUser + "&password=" + sPassword + "&login-form-type=pwd";
        ObjetoHTTP.Open("POST", InternalLogInBisRet + "pkmslogin.form", (object)false);
        ObjetoHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        ObjetoHTTP.Send(Peticion);

        if (Information.Err().Number == -2147024891)
        {
            ObjetoHTTP.Open("POST", InternalLogInBisRet, (object)false);
            ObjetoHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded");
            ObjetoHTTP.Send(Peticion);
        }

        Respuesta = Conversions.ToString(ObjetoHTTP.responseText);

        if (Strings.InStr(1, Respuesta, "page: login_success.html") > 0)
        {
            Galleta = ObjetoHTTP.getResponseHeader("Set-Cookie");
        }

        return InternalLogInBisRet;
    }
}
