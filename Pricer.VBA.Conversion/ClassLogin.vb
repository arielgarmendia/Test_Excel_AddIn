Public Class ClassLogin
    'Public bLogin As Boolean
    'Public sUser As String
    'Public sPassword As String
    'Private Sub OK_Click()
    '    Dim ObjetoHTTP As Object
    '    Dim Peticion As String
    '    Dim Respuesta As String
    '    Dim iRespuesta As Integer

    '    On Error Resume Next

    '    'para el error #value!
    '    With ThisWorkbook.Sheets("LVB Proced. Generator")
    '        .Range("Underlying_L1").Value = .Range("Underlying_L1").Value
    '    End With

    '    bLogin = False
    '    sUser = UserName.Value
    '    sPassword = Password.Value

    '    Unload LogIn

    'Call InternalLogIn()

    '    If Galleta <> "" Then
    '        Call ResetButtons("LVB Proced. Generator")
    '        ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogin.BackColor = VBA.RGB(0, 255, 0)
    '        MsgBox "Login OK.", vbInformation, "OK"

    '    iRespuesta = MsgBox("Underlyings and clients will then be loaded. Do you want to load them?", vbInformation + vbYesNo, "LOAD STATIC DATA")
    '        If iRespuesta = vbYes Then
    '            Call ButtonLoadStaticData("LVB Proced. Generator")
    '        End If
    '    Else
    '        Call ResetButtons("LVB Proced. Generator")
    '        ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0)
    '        MsgBox "Login KO.", vbCritical, "KO"
    'End If
    'End Sub

    'Function InternalLogIn() As String
    '    InternalLogIn = InternalLogInBis("")
    'End Function

    'Function InternalLogInBis(action As String) As String
    '    Dim ObjetoHTTP As Object
    '    Dim Peticion As String
    '    Dim Respuesta As String
    '    Dim pricerEnvironment As String

    '    pricerEnvironment = ThisWorkbook.Sheets("LVB Proced. Generator").Range("pricerEnvironment").Value

    '    Select Case True
    '        Case (pricerEnvironment = "PREproduction")
    '            If (action = "") Then
    '                InternalLogInBis = "https://au-wbamdesktop.es.igrupobbva/"
    '            Else
    '                InternalLogInBis = "https://ei-wbamdesktop.es.igrupobbva/"
    '            End If

    '        Case (pricerEnvironment = "PROduction")
    '            InternalLogInBis = "https://cibdesktop.es.igrupobbva/"
    '        Case (pricerEnvironment = "WindowsLocal")
    '            InternalLogInBis = "https://au-wbamdesktop.es.igrupobbva/"
    '        Case (InStr(pricerEnvironment, "CPrice") = 1)
    '            If (action = "") Then
    '                InternalLogInBis = "https://au-wbamdesktop.es.igrupobbva/"
    '            Else
    '                InternalLogInBis = "https://ei-wbamdesktop.es.igrupobbva/"
    '            End If

    '        Case Else
    '            Unload LogIn
    '        Galleta = ""
    '            Call ResetButtons("LVB Proced. Generator")
    '            ThisWorkbook.Sheets("LVB Proced. Generator").ButtonLogout.BackColor = VBA.RGB(255, 0, 0) '
    '            ThisWorkbook.Worksheets("LVB Proced. Generator").Range("pricerEnvironment").Activate
    '            MsgBox "ERROR: Pricer server not configured property!!", vbCritical + vbOKOnly, "ERROR"
    '        Exit Function
    '    End Select

    'Set ObjetoHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'With ObjetoHTTP
    '        .Open "GET", InternalLogInBis & "pkmslogout", False
    '    .SetRequestHeader "Content-Type", "text/html"
    '    On Error Resume Next
    '        .Send
    '        If Err.Number = -2147024891 Or Err.Number = -2147012744 Then
    '            .Open "GET", InternalLogInBis, False
    '        .SetRequestHeader "Content-Type", "text/html"
    '        .Send
    '        End If
    '        On Error GoTo 0
    '    End With

    'Set ObjetoHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    'With ObjetoHTTP
    '        Peticion = "username=" & sUser & "&password=" & sPassword & "&login-form-type=pwd"
    '        .Open "POST", InternalLogInBis & "pkmslogin.form", False
    '    .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    '    .Send Peticion
    '    If Err.Number = -2147024891 Then
    '            .Open "POST", InternalLogInBis, False
    '        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    '        .Send Peticion
    '    End If
    '        Respuesta = .responseText

    '        If InStr(1, Respuesta, "page: login_success.html") > 0 Then
    '            Galleta = .getResponseHeader("Set-Cookie")
    '        End If
    '    End With
    'End Function
End Class
