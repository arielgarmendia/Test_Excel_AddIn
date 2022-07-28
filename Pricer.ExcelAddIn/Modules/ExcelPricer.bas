Attribute VB_Name = "ExcelPricer"
Option Base 1
Option Explicit
Option Compare Text

Dim bLoadDatesBulk As Boolean
Public bGuardar As Boolean

Public IsArrow As Boolean

'Errores
Public Const STRING_ERROR = "ERROR"

Public Galleta As String

'colors
Public Const lBlue As Long = 16247773
Public Const lRed As Long = 255
Public Const lOrange As Long = 49407
Public Const lLightGrayAPB As Long = 15921906 'autocall product block
Public Const lDarkGrayAPB As Long = 10921638
Public Const lLightGrayDGB As Long = 15592941 'date generator block
Public Const lDarkGrayDGB As Long = 13224393

'LVB Proced. Generator
Public Const sColEquityInitialFixingDates As String = "V"
Public Const sColEquityFixingDates As String = "W"
Public Const sColEquityPaymentDates As String = "X"
Public Const sColEarlyRedemption As String = "Y"
Public Const sColEarlyRedemptionTrigger As String = "Z"
Public Const sColACCoupon As String = "AA"
Public Const sColNonCallCoupon As String = "AB"
Public Const sColSwapFixingDates As String = "AE"
Public Const sColSwapStartDates As String = "AF"
Public Const sColSwapEndDates As String = "AG"
Public Const sColSwapPaymentDates As String = "AH"
Public Const sColSwapSpread As String = "AI"
Public Const sColBarrierObservationDates As String = "AK"
Public Const sColValores1 As String = "O"
Public Const sColValores2 As String = "S"
'Aux
Public Const sColMonthsEMTNRAR As String = "AA" 'typhoon
Public Const sColEMTN As String = "AB" 'typhoon
Public Const sColRAR As String = "AC" 'typhoon
Public Const sColMonths As String = "AD" '
Public Const sColSumMonths As String = "AE"
Public Const sColEMTN2 As String = "AF"


Public bMapea As Boolean
Public bReset As Boolean
Public bLoadFMM As Boolean
Public bInsert As Boolean 'para insertar o copiar columna bulk
Public bCopy As Boolean 'copy to bulk
Public bClone As Boolean 'clone bulk
Public bConcatenate As Boolean

Public sB_BarrierType_V As String
Public sB_Direction_V As String
Public sB_DeltaCap_V As Double
Public sB_DeltaFloor_V As Double
Public sB_DeltaLiquidityAlpha_V As Double
Public sB_MaxDelta_V As String
Public sB_MaxDeltaValue_V As Double
Public sAc_BarrierType_V As String
Public sAc_BarrierLevel_V As String
Public sAc_BarrierShiftType_V As String


Sub Auto_Open()
    Application.Calculation = xlCalculationAutomatic
End Sub

Function PeticionHTTP(ObjetoHTTP As Object, Peticion As String) As String
    If Galleta = "" Then Call LoginError
    
    With ObjetoHTTP
        .SetRequestHeader "Cookie", Galleta
    
        On Error Resume Next
        .Send Peticion
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical + vbOKOnly, "ERROR"
            'End
            Exit Function
        End If
        On Error GoTo 0
    
        PeticionHTTP = .responseText
    End With
End Function

Sub ButtonLoadDates(sHoja As String, iCol As Integer, sProduct As String)
    Dim NumFila As Long
    Dim Instruments As Variant
    Dim sCal As String
    Dim sCalendars As String
    Dim mDatesCal As Variant
    Dim mDatesCurrencies As Variant
    Dim mSetCalendar As Variant
    Dim sCurrency As String
    Dim mData1 As Variant
    Dim mData2 As Variant
    Dim mData As Variant
    Dim UltFila As Integer
    Dim mFixingDates As Variant
    Dim mEquityLeg As Variant
    Dim mEquitySwapLeg As Variant
    Dim mGenerateAutocallable As Variant
    Dim iDiffMonths As Integer
    Dim iMinDiff As Integer
    Dim iMonth As Integer
    Dim sPeriod As String
    Dim i As Integer
    Dim Rango As Range
    Dim mColBarrierObservationDates As Variant
    
    If Galleta = "" Then Call LoginError
    
    Call ProtectSheet(False, sHoja)
    
    On Error GoTo ERRORES
    
    If VBA.Left(sHoja, 4) <> "Bulk" Then
        With ThisWorkbook.Sheets(sHoja)
            .Range("Ac_Underlying_V").Offset(0, iCol).Formula = "=UPPER(ConcatenatePricer(Underlyings))"
            .Range("GI_InstrumentId_V").Offset(0, iCol).Value = "=Ac_Underlying_V"
            .Range("DatesAcCouponsBlock").Value = ""
            .Range("DatesSwapSpreadBlock").Value = ""
        End With
    End If
    
    Call ProtectSheet(False, sHoja)
    
    Call FormatDatesCells(sHoja, iCol, sProduct)
    
    Application.EnableEvents = False
    
    With ThisWorkbook.Sheets(sHoja)
        Instruments = Split(.Range("GI_InstrumentId_V").Offset(0, iCol).Value, ";")

        For NumFila = 0 To UBound(Instruments)
            sCal = getCalendar(Trim(Instruments(NumFila)))
            If VBA.InStr(1, sCalendars, sCal, vbTextCompare) = 0 Then sCalendars = sCalendars & "+" & sCal
        Next NumFila
        sCalendars = VBA.Mid(sCalendars, 2)
        .Range("EL_FixingCalendar_V").Offset(0, iCol).Formula = sCalendars
        .Range("BOD_FixingCalendar_V").Offset(0, iCol).Formula = sCalendars

        sCurrency = getCurrencies(Worksheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "PaymentCalendar")
        .Range("EL_PaymentCalendar_V").Offset(0, iCol).Value = sCurrency
        .Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value = sCurrency
        .Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value = sCurrency
        
        For NumFila = 0 To 0
            mDatesCal = getHolidaysArray(sCal)
            mDatesCurrencies = getHolidaysArray(sCurrency)
        Next NumFila
        
        With ThisWorkbook.Sheets("Aux")
            .Range("AG2:AG" & (UBound(mDatesCal) + 1)).Value = mDatesCal
            Set Rango = .Range("AG2:AG" & (UBound(mDatesCal) + 1))
            mSetCalendar = Application.Run("QBS.DateGen.SetCalendar", sCalendars, Rango)
        
            .Range("AI2:AI" & (UBound(mDatesCurrencies) + 1)).Value = mDatesCurrencies
            Set Rango = .Range("AI2:AI" & (UBound(mDatesCurrencies) + 1))
            mSetCalendar = Application.Run("QBS.DateGen.SetCalendar", sCurrency, Rango)
        End With
        
        If VBA.Left(sHoja, 4) <> "Bulk" Then
            With .Range("O_ValueDate_V").Offset(0, iCol)
                .Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,""TARGET"")"
                .Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,EL_PaymentCalendar_V)"
            End With
            With .Range("ESL_SwapEndDate_V").Offset(0, iCol)
                .Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,""TARGET"")"
                .Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,EL_PaymentCalendar_V)"
            End With
        Else 'bulk
            .Range("O_ValueDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" & .Range("Ac_ExpiryDate_V").Offset(0, iCol).Address & "," & .Range("Ac_PaymentShifter_V").Offset(0, iCol).Address & ",""TARGET"")"
            .Range("O_ValueDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" & .Range("Ac_ExpiryDate_V").Offset(0, iCol).Address & "," & .Range("Ac_PaymentShifter_V").Offset(0, iCol).Address & "," & .Range("EL_PaymentCalendar_V").Offset(0, iCol).Address & ")"
            
            .Range("ESL_SwapEndDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" & .Range("Ac_ExpiryDate_V").Offset(0, iCol).Address & "," & .Range("Ac_PaymentShifter_V").Offset(0, iCol).Address & ",""TARGET"")"
            .Range("ESL_SwapEndDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" & .Range("Ac_ExpiryDate_V").Offset(0, iCol).Address & "," & .Range("Ac_PaymentShifter_V").Offset(0, iCol).Address & "," & .Range("EL_PaymentCalendar_V").Offset(0, iCol).Address & ")"
        End If
        
        mFixingDates = Application.Run("QBS.DateGen.Param.FixingDates", VBA.CLng(.Range("AS_StartDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("AS_EndDate_V").Offset(0, iCol).Value), .Range("AS_Frequency_V").Offset(0, iCol).Value)
        mEquityLeg = Application.Run("QBS.DateGen.Param.EquityLeg", mFixingDates, VBA.CLng(.Range("D_InitialPaymentDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("D_FinalAlignmentDate_V").Offset(0, iCol).Value), .Range("EL_Frequency_V").Offset(0, iCol).Value, .Range("EL_PaymentLag_V").Offset(0, iCol).Value, .Range("EL_FixingCalendar_V").Offset(0, iCol).Value, .Range("EL_PaymentCalendar_V").Offset(0, iCol).Value, .Range("EL_Alignment_V").Offset(0, iCol).Value, .Range("EL_BrokenPeriod_V").Offset(0, iCol).Value, .Range("EL_FixingAdjustment_V").Offset(0, iCol).Value, .Range("EL_PaymentAdjustment_V").Offset(0, iCol).Value, .Range("EL_StickToMothEnd_V").Offset(0, iCol).Value, .Range("EL_AdjustInputDates_V").Offset(0, iCol).Value)
        mEquitySwapLeg = Application.Run("QBS.DateGen.Param.EquitySwapLeg", .Range("ESL_SwapFrequency_V").Offset(0, iCol).Value, .Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value, .Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value, "2B", .Range("ESL_SwapAlignment_V").Offset(0, iCol).Value, .Range("ESL_SwapBrokenPeriod_V").Offset(0, iCol).Value, .Range("ESL_SwapPaymentAdjustment_V").Offset(0, iCol).Value, VBA.CLng(.Range("ESL_SwapStartDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("ESL_SwapEndDate_V").Offset(0, iCol).Value), .Range("ESL_SwapStickToMothEnd_V").Offset(0, iCol).Value, IIf(VBA.IsEmpty(.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value) = True, False, .Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value))
        mGenerateAutocallable = Application.Run("QBS.DateGen.GenerateAutocallable", mEquityLeg, mEquitySwapLeg, .Range("A_EarlyRedemptionFreq_V").Offset(0, iCol).Value, .Range("A_FirstEarlyRedemptionPer_V").Offset(0, iCol).Value, .Range("A_EarlyRedemptionAlignment_V").Offset(0, iCol).Value, .Range("A_AllowEarlyRedemptionMat_V").Offset(0, iCol).Value, IIf(VBA.IsEmpty(.Range("A_Dates_V").Offset(0, iCol).Value) = True, False, .Range("A_Dates_V").Offset(0, iCol).Value), True, True)
        
        On Error Resume Next
        If VBA.InStr(1, mGenerateAutocallable, "Error:", vbTextCompare) > 0 And VBA.Left(sHoja, 4) = "Bulk" Then
            If Err.Number = 0 Then
                .Range("Error_V").Offset(0, iCol).Value = mGenerateAutocallable
                Exit Sub
            End If
        ElseIf VBA.InStr(1, mGenerateAutocallable, "Error:", vbTextCompare) > 0 And VBA.Left(sHoja, 4) <> "Bulk" Then
            If Err.Number = 0 Then
                MsgBox mGenerateAutocallable, vbCritical + vbOKOnly, "ERROR"
                Exit Sub
            End If
        End If
        On Error GoTo ERRORES
        
        ReDim mData1(1 To 10000, 1 To 3)
        ReDim mData2(1 To 10000, 1 To 3)
        For NumFila = 1 To UBound(mGenerateAutocallable)
            mData1(NumFila, 1) = .Range("ER_InitialTriggerRate_V").Offset(0, iCol).Value
            'AC Coupon (%)
            Select Case .Range("Ac_ACCoupon_V").Offset(0, iCol).Value
                Case "Flat" 'valor fijo
                    mData1(NumFila, 2) = .Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value
                Case "Coupon Step" 'se va incrementando valor
                    If NumFila = 2 Then
                        mData1(NumFila, 2) = .Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value
                    Else
                        mData1(NumFila, 2) = .Range(sColACCoupon & NumFila) + .Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value
                    End If
            End Select
            mData1(NumFila, 3) = .Range("Ac_NonCallCoupon_V").Offset(0, iCol).Value
            
            'swap spread
            If NumFila < UBound(mGenerateAutocallable) Then
                If mGenerateAutocallable((NumFila + 1), 7) <> "" Then
                    iDiffMonths = VBA.DateDiff("m", VBA.CDate(mGenerateAutocallable((NumFila + 1), 7)), VBA.CDate(mGenerateAutocallable((NumFila + 1), 8)))
                    If NumFila = 1 Then iMinDiff = iDiffMonths
                    Select Case iDiffMonths
                        Case Is < 3
                            mData2(NumFila, 1) = 1
                        Case Is < 6
                            mData2(NumFila, 1) = 3
                        Case Else
                            mData2(NumFila, 1) = 6
                    End Select
                    If NumFila = 1 Then
                        mData2(NumFila, 2) = mData2(NumFila, 1)
                    ElseIf NumFila > 1 Then
                        mData2(NumFila, 2) = mData2((NumFila - 1), 2) + mData2(NumFila, 1)
                    End If
                    If mData2(NumFila, 2) < 12 Then
                        sPeriod = mData2(NumFila, 2) & "m"
                    ElseIf mData2(NumFila, 2) = 12 Then 'igual al año
                        sPeriod = VBA.Int(mData2(NumFila, 2) / 12) & "y"
                    Else 'MÁS QUE UN AÑO
                        sPeriod = VBA.Replace(VBA.Int(mData2(NumFila, 2) / 12) & "y" & IIf(VBA.Int(mData2(NumFila, 2) Mod 12) = 0, "", VBA.Int(mData2(NumFila, 2) Mod 12) & "m"), "y0m", "")
                    End If
                    ThisWorkbook.Sheets("Aux").Range(sColEMTN2 & (NumFila + 2)).Formula = "=IFERROR(VLOOKUP(" & """" & sPeriod & """" & ", " & sColMonthsEMTNRAR & "1:" & sColEMTN & "121, 2, FALSE), 0)"
                    If iDiffMonths < iMinDiff Then iMinDiff = iDiffMonths
                End If
            End If
        Next NumFila

        Select Case iMinDiff
            Case Is < 3
                iMonth = 1
            Case Is < 6
                iMonth = 3
            Case Else
                iMonth = 6
        End Select

        With ThisWorkbook.Sheets("Aux")
            .Range(sColMonthsEMTNRAR & "1:" & sColEMTN2 & "121").NumberFormat = "General"
            If ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value = "EUR" Then
                .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").FormulaArray = "=bbva_GetTyMatrix(""REFERENCE"",""IR_NOTE_SPREAD"",TODAY()," & VBA.Chr(34) & VBA.Trim(VBA.Replace(ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "EUR", "") & " MTN " & iMonth) & "m Callable Spread"")"
            Else 'cualquier otra divisa
                .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").FormulaArray = "=bbva_GetTyMatrix(""REFERENCE"",""IR_NOTE_SPREAD"",TODAY()," & VBA.Chr(34) & ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value & " MTN 1m Callable Spread"")"
            End If
            If .Range(sColMonthsEMTNRAR & "1").Value = "You are not logged" Then
                MsgBox "You are not logged in Typhoon Add-Ins", vbCritical + vbOKOnly, "ERROR LOGIN"
                Exit Sub
            ElseIf VBA.InStr(1, .Range(sColMonthsEMTNRAR & "1").Value, "ERROR", vbTextCompare) > 0 Then
                MsgBox .Range(sColMonthsEMTNRAR & "1").Value, vbCritical + vbOKOnly, "ERROR"
                Exit Sub
            End If
            .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").Copy
            .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").PasteSpecial xlPasteValues
            Application.CutCopyMode = False

            UltFila = .Range(sColEMTN2 & .Rows.Count).End(xlUp).row
            mData = .Range(sColEMTN2 & "3:" & sColEMTN2 & UltFila).Value
            If UltFila >= 3 Then
                For i = 1 To (UltFila - 2)
                    mData(i, 1) = mData(i, 1) / 10000
                Next i
            End If
        End With
        '****
        
        For NumFila = 1 To .Range("Ac_NonCancelPeriods_V").Offset(0, iCol).Value
            mData1(NumFila, 1) = 99.99
            mData1(NumFila, 3) = 0
        Next NumFila
        
        If VBA.Left(sHoja, 4) <> "Bulk" Then
            Application.Calculation = xlCalculationManual
            
            On Error Resume Next
            For NumFila = 2 To UBound(mGenerateAutocallable) ' - 1
                If mGenerateAutocallable(NumFila, 1) <> "" Then _
                    .Range(sColEquityInitialFixingDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 1))
                .Range(sColEquityFixingDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable((NumFila + 1), 2))
                .Range(sColEquityPaymentDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable((NumFila + 1), 3))
                .Range(sColEarlyRedemption & (NumFila + 1)).Value = mGenerateAutocallable((NumFila + 1), 5)
                .Range(sColSwapFixingDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 6))
                .Range(sColSwapStartDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 7))
                .Range(sColSwapEndDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 8))
                .Range(sColSwapPaymentDates & (NumFila + 1)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 9))
            Next NumFila
            On Error GoTo ERRORES
            
            UltFila = .Range(sColEquityFixingDates & .Rows.Count).End(xlUp).row
            .Range(sColEarlyRedemptionTrigger & "3:" & sColNonCallCoupon & UltFila).Value = mData1
            
            UltFila = .Range(sColSwapStartDates & .Rows.Count).End(xlUp).row
            .Range(sColSwapSpread & "3:" & sColSwapSpread & UltFila).Value = mData
            
            UltFila = .Range(sColBarrierObservationDates & .Rows.Count).End(xlUp).row
            If UltFila >= 3 Then .Range(sColBarrierObservationDates & "3:" & sColBarrierObservationDates & UltFila).Formula = ""
            Select Case .Range("B_ObservationType_V").Offset(0, iCol).Value
                Case "AtExpiry"
                    .Range(sColBarrierObservationDates & "3").Formula = "=Ac_ExpiryDate_V"
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "CallSpread"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.015
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
                Case "Continuous"
                    .Range(sColBarrierObservationDates & "3").Formula = "=Ac_StrikeDate_V"
                    .Range(sColBarrierObservationDates & "4").Formula = "=Ac_ExpiryDate_V"
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
                Case "Daily"
                    '.Range(sColBarrierObservationDates & "3:" & sColBarrierObservationDates & "2000").FormulaArray = "=QBS.DateGen.GenerateSchedule(BOD_FirstObservationDate_V,BOD_LastObservationDate_V,""1B"",BOD_FixingCalendar_V,,,,,BOD_AdjustInputDates_V,TRUE)"
                    mColBarrierObservationDates = Application.Run("QBS.DateGen.GenerateSchedule", VBA.CLng(.Range("BOD_FirstObservationDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("BOD_LastObservationDate_V").Offset(0, iCol).Value), "1B", .Range("BOD_FixingCalendar_V").Offset(0, iCol).Value, , , , , .Range("BOD_AdjustInputDates_V").Offset(0, iCol).Value, True)
                    .Range(sColBarrierObservationDates & "3:" & sColBarrierObservationDates & (UBound(mColBarrierObservationDates) + 2)).Value = mColBarrierObservationDates
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
            End Select
            
            UltFila = .Range(sColBarrierObservationDates & .Rows.Count).End(xlUp).row
            If UltFila >= 3 Then .Range("B_ObservationDates_V").Formula = "=ConcatenatePricer($" & sColBarrierObservationDates & "$3:$" & sColBarrierObservationDates & "$" & UltFila & ")"
            'colors porque es automatico
            .Range("B_ObservationDates_V").Interior.Color = lDarkGrayAPB
            .Columns(sColBarrierObservationDates & ":" & sColBarrierObservationDates).Interior.Pattern = xlNone
            
            Application.Calculation = xlCalculationAutomatic
        Else 'bulk
            .Range("S_ObservationDates_V").Offset(0, iCol).Value = ""
            .Range("APB_Dates1").Offset(0, iCol).Value = ""
            .Range("APB_Dates2").Offset(0, iCol).Value = ""
            .Range("Error_V").Offset(0, iCol).Value = ""
            On Error Resume Next
            For NumFila = 1 To UBound(mGenerateAutocallable) ' - 2)
                If mGenerateAutocallable((NumFila + 1), 1) <> "" Then _
                    .Range("S_ObservationDates_V").Offset(0, iCol).Value = .Range("S_ObservationDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 1), 1)), "yyyy-mm-dd") & "; "
                .Range("ER_FixingDates_V").Offset(0, iCol).Value = .Range("ER_FixingDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 2), 2)), "yyyy-mm-dd") & "; "
                .Range("ER_SettlementDates_V").Offset(0, iCol).Value = .Range("ER_SettlementDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 2), 3)), "yyyy-mm-dd") & "; "

                If Err.Number = 9 Then Exit For
                If mGenerateAutocallable((NumFila + 2), 3) <> "" Then
                    .Range("ER_TriggerRates_V").Offset(0, iCol).Value = .Range("ER_TriggerRates_V").Offset(0, iCol).Value & mData1(NumFila, 1) & "; "
                    .Range("ER_TriggerPayments_V").Offset(0, iCol).Value = .Range("ER_TriggerPayments_V").Offset(0, iCol).Value & mData1(NumFila, 2) & "; "
                    .Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value = .Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value & mData1(NumFila, 3) & "; "
                End If
            Next NumFila
            Err.Clear
            For NumFila = 1 To UBound(mGenerateAutocallable) ' - 2)
                .Range("I_AccrualStartDates_V").Offset(0, iCol).Value = .Range("I_AccrualStartDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 1), 7)), "yyyy-mm-dd") & "; "
                .Range("I_AccrualEndDates_V").Offset(0, iCol).Value = .Range("I_AccrualEndDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 1), 8)), "yyyy-mm-dd") & "; "
                .Range("I_FixingDates_V").Offset(0, iCol).Value = .Range("I_FixingDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 1), 6)), "yyyy-mm-dd") & "; "
                .Range("I_SettlementDates_V").Offset(0, iCol).Value = .Range("I_SettlementDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mGenerateAutocallable((NumFila + 1), 9)), "yyyy-mm-dd") & "; "
                'mismo número de swap spread que fechas
                If Err.Number = 0 Then _
                    .Range("I_SpreadValues_V").Offset(0, iCol).Value = .Range("I_SpreadValues_V").Offset(0, iCol).Value & mData(NumFila, 1) & "; "
            Next NumFila
            On Error GoTo ERRORES
            With .Range("S_ObservationDates_V").Offset(0, iCol)
                .NumberFormat = "@"
                .Value = VBA.Left(.Value, VBA.Len(.Value) - 2)
            End With
            .Range("ER_FixingDates_V").Offset(0, iCol).Value = VBA.Left(.Range("ER_FixingDates_V").Offset(0, iCol).Value, VBA.Len(.Range("ER_FixingDates_V").Offset(0, iCol).Value) - 2)
            .Range("ER_SettlementDates_V").Offset(0, iCol).Value = VBA.Left(.Range("ER_SettlementDates_V").Offset(0, iCol).Value, VBA.Len(.Range("ER_SettlementDates_V").Offset(0, iCol).Value) - 2)
            .Range("ER_TriggerRates_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(.Range("ER_TriggerRates_V").Offset(0, iCol).Value, ",", "."), VBA.Len(.Range("ER_TriggerRates_V").Offset(0, iCol).Value) - 2)
            .Range("ER_TriggerPayments_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(.Range("ER_TriggerPayments_V").Offset(0, iCol).Value, ",", "."), VBA.Len(.Range("ER_TriggerPayments_V").Offset(0, iCol).Value) - 2)
            .Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value, ",", "."), VBA.Len(.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value) - 2)
            .Range("I_AccrualStartDates_V").Offset(0, iCol).Value = VBA.Left(.Range("I_AccrualStartDates_V").Offset(0, iCol).Value, VBA.Len(.Range("I_AccrualStartDates_V").Offset(0, iCol).Value) - 2)
            .Range("I_AccrualEndDates_V").Offset(0, iCol).Value = VBA.Left(.Range("I_AccrualEndDates_V").Offset(0, iCol).Value, VBA.Len(.Range("I_AccrualEndDates_V").Offset(0, iCol).Value) - 2)
            .Range("I_FixingDates_V").Offset(0, iCol).Value = VBA.Left(.Range("I_FixingDates_V").Offset(0, iCol).Value, VBA.Len(.Range("I_FixingDates_V").Offset(0, iCol).Value) - 2)
            .Range("I_SettlementDates_V").Offset(0, iCol).Value = VBA.Left(.Range("I_SettlementDates_V").Offset(0, iCol).Value, VBA.Len(.Range("I_SettlementDates_V").Offset(0, iCol).Value) - 2)
            .Range("I_SpreadValues_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(.Range("I_SpreadValues_V").Offset(0, iCol).Value, ",", "."), VBA.Len(.Range("I_SpreadValues_V").Offset(0, iCol).Value) - 2)
         
            .Range("B_ObservationDates_V").Offset(0, iCol).Value = ""
            .Range("B_ObservationDates_V").Offset(0, iCol).NumberFormat = "@"
            Select Case .Range("B_ObservationType_V").Offset(0, iCol).Value
                Case "AtExpiry"
                    mColBarrierObservationDates = VBA.Split(.Range("Ac_ExpiryDate_V").Offset(0, iCol).Value & ";", ";")
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "CallSpread"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.015
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
                    For NumFila = 0 To (UBound(mColBarrierObservationDates) - 1)
                        .Range("B_ObservationDates_V").Offset(0, iCol).Value = .Range("B_ObservationDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mColBarrierObservationDates(NumFila)), "yyyy-mm-dd") & "; "
                    Next NumFila
                
                Case "Continuous"
                    mColBarrierObservationDates = VBA.Split(.Range("Ac_StrikeDate_V").Offset(0, iCol).Value & ";" & .Range("Ac_ExpiryDate_V").Offset(0, iCol).Value & ";", ";")
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
                    For NumFila = 0 To (UBound(mColBarrierObservationDates) - 1)
                        .Range("B_ObservationDates_V").Offset(0, iCol).Value = .Range("B_ObservationDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mColBarrierObservationDates(NumFila)), "yyyy-mm-dd") & "; "
                    Next NumFila
                
                Case "Daily"
                    mColBarrierObservationDates = Application.Run("QBS.DateGen.GenerateSchedule", VBA.CLng(.Range("BOD_FirstObservationDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("BOD_LastObservationDate_V").Offset(0, iCol).Value), "1B", .Range("BOD_FixingCalendar_V").Offset(0, iCol).Value, , , , , .Range("BOD_AdjustInputDates_V").Offset(0, iCol).Value, True)
                    .Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift"
                    .Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01
                    .Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB
                    For NumFila = 1 To (UBound(mColBarrierObservationDates))
                        .Range("B_ObservationDates_V").Offset(0, iCol).Value = .Range("B_ObservationDates_V").Offset(0, iCol).Value & VBA.Format(VBA.CDate(mColBarrierObservationDates(NumFila, 1)), "yyyy-mm-dd") & "; "
                    Next NumFila
            End Select
            
            If VBA.Trim(.Range("B_ObservationDates_V").Offset(0, iCol).Value) <> "" Then .Range("B_ObservationDates_V").Offset(0, iCol).Value = VBA.Left(.Range("B_ObservationDates_V").Offset(0, iCol).Value, VBA.Len(.Range("B_ObservationDates_V").Offset(0, iCol).Value) - 2)
            
            'colors porque es automatico
            .Range("B_ObservationDates_V").Offset(0, iCol).Interior.Color = lDarkGrayAPB
        End If
    End With
    
    Application.EnableEvents = True
    Exit Sub
    
ERRORES:
    MsgBox Err.Description
    Err.Clear
    Call LoginError
End Sub

Sub ButtonGenerateFMM(sHoja As String, iCol As Integer, sProduct As String)
    Dim Mapa() As String
    Dim Peticion As String
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim SistemaArchivos As FileSystemObject
    Dim ArchivoTexto As TextStream
    Dim texto As String
    Dim sRuta As String
    Dim sFichero As String
    Const sHojaBulk As String = "Bulk mode"
    
    If Galleta = "" Then Call LoginError
    
    Mapa = Mapea(sHoja, iCol, sProduct)

    If Mapa(1, 1) <> STRING_ERROR Then
        Peticion = JsonConverter.ConvertToJson(ObjetoJSON(Mapa), 2)
        Set RespuestaJSON = Services.GenerateFMM(Peticion, Respuesta)
        
        If Respuesta = "" Then Exit Sub
        
        With ThisWorkbook.Sheets(sHoja)
            If VBA.Left(sHoja, 4) = "Bulk" Then .Range("Error_V").Offset(0, iCol).Value = ""
            sRuta = IIf(VBA.Left(sHoja, 4) <> "Bulk", ActiveWorkbook.Path, ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value)
            sFichero = IIf(VBA.Left(sHoja, 4) <> "Bulk", "fmm", ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value)
        
            If VBA.Trim(sRuta) = "" Then
                ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value = "Folder of the FMM is empty."
                Exit Sub
            ElseIf VBA.Trim(sFichero) = "" Then
                ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value = "Name of the FMM is empty."
                Exit Sub
            End If
        End With
            
        Set SistemaArchivos = New FileSystemObject
        On Error Resume Next
        Set ArchivoTexto = SistemaArchivos.CreateTextFile(sRuta & "/" & sFichero & ".txt", True)
        If Err.Number <> 0 Then
            If VBA.Left(sHoja, 4) = "Bulk" Then
                ThisWorkbook.Sheets(sHoja).Range("Error_V").Offset(0, iCol).Value = Err.Description
                Err.Clear
            End If
        Else
            With ArchivoTexto
                .Write RespuestaJSON("xmlFmm")
                .Close
            End With
            
            texto = DecodeBase64(RespuestaJSON("xmlFmm"))
            
            Set ArchivoTexto = SistemaArchivos.CreateTextFile(sRuta & "/" & sFichero & ".xml", True)
            With ArchivoTexto
                .Write texto
                .Close
            End With
            
            ThisWorkbook.Sheets(sHoja).Range("DealID").Offset(0, iCol).ClearContents
        End If
    End If
End Sub

Sub ButtonEditFMM()
    Shell "notepad " & ActiveWorkbook.Path & "/fmm.xml", vbNormalFocus
End Sub

Sub ButtonCalculatePrice(sHoja As String, iCol As Integer)
    Dim strFileExists As String
    Dim strFileName As String
    Dim Respuesta As String
    Dim pricerEnvironment As String
    Dim RespuestaJSON As Object
    Const sHojaBulk As String = "Bulk mode"
    
    If Galleta = "" Then Call LoginError
    
    With ThisWorkbook.Sheets(sHoja)
        Application.Calculation = xlCalculationManual
        .Range("DealID").Offset(0, iCol).Value = ""
    
        strFileName = IIf(VBA.Left(sHoja, 4) <> "Bulk", ActiveWorkbook.Path & "/fmm.xml", ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value & "\" & ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value & ".xml")
        strFileExists = Dir(strFileName)
        If strFileExists = "" Then
            Select Case VBA.Left(sHoja, 4)
                Case Is <> "Bulk"
                    MsgBox "Please click on generateFMM before calculating price"
                
                Case Else
                    .Range("DealID").Offset(0, iCol).Value = "Please click on generateFMM before calculating price"
            End Select
        Else
            Call ProtectSheet(False, sHoja)
            pricerEnvironment = ThisWorkbook.Sheets("LVB Proced. Generator").Range("pricerEnvironment").Value
    
            Set RespuestaJSON = Services.CalculatePrice(pricerEnvironment, strFileName, ThisWorkbook.Sheets("LVB Proced. Generator").Range("CalculationType").Value, .Range("DealID").Offset(0, iCol).Value & "", Respuesta)
            
            If Respuesta = "" Then Exit Sub
            
            Select Case VBA.Left(sHoja, 4)
                Case Is <> "Bulk"
                    Application.Calculation = xlCalculationManual
                    .Range("Result").Clear
                    If .Range("DealID").Value = "" Then
                        MsgBox "Deal ID " & RespuestaJSON("qtpdId") & " obtained.", vbInformation
                        .Range("DealID").Value = RespuestaJSON("qtpdId")
                    ElseIf .Range("DealID").Value = RespuestaJSON("qtpdId") Then
                        MsgBox "Deal ID " & RespuestaJSON("qtpdId") & " confirmed.", vbInformation
                    Else
                        MsgBox "An error has occurred: Deal ID " & RespuestaJSON("qtpdId") & " received.", vbExclamation
                    End If
                    
                Case Else
                    .Range("DealID").Offset(0, iCol).ClearContents
                    .Range("Result").Offset(0, iCol).ClearContents
                    If .Range("DealID").Offset(0, iCol).Value = "" Then
                        .Range("DealID").Offset(0, iCol).Value = RespuestaJSON("qtpdId")
                    ElseIf .Range("DealID").Offset(0, iCol).Value = RespuestaJSON("qtpdId") Then
                        .Range("DealID").Offset(0, iCol).Value = "Deal ID " & RespuestaJSON("qtpdId") & " confirmed."
'                    Else
                        .Range("DealID").Offset(0, iCol).Value = "An error has occurred: Deal ID " & RespuestaJSON("qtpdId") & " received."
                    End If
            End Select
        End If
        Application.Calculation = xlCalculationAutomatic
    End With
End Sub

Sub ButtonGenerateFMMAndCalculatePrice(sHoja As String, iCol As Integer, sProduct As String)
    Call ButtonGenerateFMM(sHoja, iCol, sProduct)
    Call ButtonCalculatePrice(sHoja, iCol)
End Sub

Sub ButtonGetResult(sHoja As String, iCol As Integer, sProduct As String)
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim NumGriega As Long
    Dim Destino As Range
    
    If Galleta = "" Then Call LoginError
    
    With ThisWorkbook.Sheets(sHoja)
        If IsEmpty(.Range("DealID").Offset(0, iCol).Value) Then
            Select Case VBA.Left(sHoja, 4)
                Case Is <> "Bulk"
                    MsgBox "Deal Id must be filled. Please send to calculate the deal before trying to get results"
                
                Case Else
                    .Range("DealID").Offset(0, iCol).Value = "Deal Id must be filled. Please send to calculate the deal before trying to get results"
            End Select
        Else
            Set RespuestaJSON = Services.GetResult(.Range("DealID").Offset(0, iCol).Value & "", Respuesta)
            
            If Respuesta = "" Then Exit Sub
            
            Call GetResults(sHoja, iCol, sProduct, Respuesta, RespuestaJSON)
        End If
    End With
End Sub

Sub ButtonRetrieveXMLs(sHoja As String)
    Dim ArchivoTexto As TextStream
    Dim SistemaArchivos As FileSystemObject
    Dim ObjetoHTTP As Object
    Dim URL As String
    Dim Peticion As String
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim objIEBrowser
    
    If Galleta = "" Then Call LoginError
    
    With ThisWorkbook.Sheets(sHoja)
        If IsEmpty(.Range("DealID").Value) Then
            MsgBox "Deal Id must be filled. Please send to calculate the deal before trying to retrieve xmls"
        Else
            Set RespuestaJSON = Services.RetrieveXMLs(.Range("DealID").Value & "", Respuesta)
            
            If Respuesta = "" Then Exit Sub
        
            If InStr(1, Respuesta, """error"":") > 0 Then
                MsgBox "Error retrieving XML.", vbExclamation
            ElseIf InStr(1, Respuesta, """content"":[]") > 0 Or InStr(1, Respuesta, "{}") > 0 Then
                MsgBox "Xml files not found.", vbExclamation
            Else
                Set SistemaArchivos = New FileSystemObject
                Set ArchivoTexto = SistemaArchivos.CreateTextFile(ActiveWorkbook.Path & "/fd" & .Range("DealID").Value & ".xml", True)
                With ArchivoTexto
                    .Write DecodeBase64(RespuestaJSON("content")(1)("content"))
                    .Close
                End With
                Set ArchivoTexto = SistemaArchivos.CreateTextFile(ActiveWorkbook.Path & "/fi" & .Range("DealID").Value & ".xml", True)
                With ArchivoTexto
                    .Write DecodeBase64(RespuestaJSON("content")(2)("content"))
                    .Close
                End With
                Set ArchivoTexto = SistemaArchivos.CreateTextFile(ActiveWorkbook.Path & "/fmm" & .Range("DealID").Value & ".xml", True)
                With ArchivoTexto
                    .Write DecodeBase64(RespuestaJSON("content")(3)("content"))
                    .Close
                End With
                
                Set objIEBrowser = CreateObject("InternetExplorer.Application")
                With objIEBrowser
                    .Visible = True
                    .Navigate2 ActiveWorkbook.Path & "/fd" & ThisWorkbook.Sheets(sHoja).Range("DealID").Value & ".xml"
                    .Navigate2 ActiveWorkbook.Path & "/fi" & ThisWorkbook.Sheets(sHoja).Range("DealID").Value & ".xml", 2048
                    .Navigate2 ActiveWorkbook.Path & "/fmm" & ThisWorkbook.Sheets(sHoja).Range("DealID").Value & ".xml", 2048
                    Do While .Busy
                    Loop
                End With
            End If
        End If
    End With
End Sub

Private Function ObjetoJSON(Mapa() As String) As Dictionary
    Dim NumFila As Long

    NumFila = 1
    Set ObjetoJSON = New Dictionary

    Do While Mapa(NumFila, 1) <> ""
        If Val(Mapa(NumFila, 3)) = 0 Then
            If Mapa(NumFila, 2) <> "" Then
                ObjetoJSON.Add Mapa(NumFila, 1), Mapa(NumFila, 2)
            Else
                ObjetoJSON.Add Mapa(NumFila, 1), MatrizJSON(Mapa, NumFila + 1)
            End If
        End If
        NumFila = NumFila + 1
    Loop
End Function

Private Function MatrizJSON(Mapa() As String, NumFilaInicial As Long) As Variant
    Dim Matriz() As Dictionary
    Dim ObjetoJSON As Dictionary
    Dim NumElementos As Long
    Dim NumFila As Long

    NumFila = NumFilaInicial
    NumElementos = 0

    Do While Val(Mapa(NumFila, 3)) >= Val(Mapa(NumFilaInicial, 3))
        If Val(Mapa(NumFila, 3)) = Val(Mapa(NumFilaInicial, 3)) Then
            If Mapa(NumFila, 1) = Mapa(NumFilaInicial, 1) Then
                If NumElementos >= 1 Then Set Matriz(NumElementos) = ObjetoJSON
                NumElementos = NumElementos + 1
                ReDim Preserve Matriz(1 To NumElementos)
                Set ObjetoJSON = New Dictionary
            End If
            If Mapa(NumFila, 2) <> "" Then
                ObjetoJSON.Add Mapa(NumFila, 1), Mapa(NumFila, 2)
            Else
                ObjetoJSON.Add Mapa(NumFila, 1), MatrizJSON(Mapa, NumFila + 1)
            End If
        End If
        NumFila = NumFila + 1
    Loop
    Set Matriz(NumElementos) = ObjetoJSON

    MatrizJSON = Matriz
End Function

Function getCurrencies(ByVal IdCurrency As String, Optional ByVal Key As String = "") As Variant
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim NumElemento As Long
    Dim NumColumna As Long
    Dim Origen As Range
    Dim Salida As Boolean
    
    If Galleta = "" Then Call LoginError
    
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getCurrencies", "currency", IdCurrency)
    
    If Respuesta = "" Then Exit Function
    
    If InStr(1, Respuesta, """error"":""""") > 0 Then
        getCurrencies = "Currency not found."
    ElseIf InStr(1, Respuesta, """object"":""currency""") Then
        If Key = "" Then
            ReDim Resultado(1 To 1, 1 To 4)
            Resultado(1, 1) = RespuestaJSON("PaymentCalendar")
            Resultado(1, 2) = RespuestaJSON("IBORIndex")
            Resultado(1, 3) = RespuestaJSON("SwapFixingCalendar")
            Resultado(1, 4) = RespuestaJSON("SwapPaymentCalendar")
            If Application.Caller.HasArray Then
                getCurrencies = Resultado
            Else
                Set Origen = Application.Caller
                NumElemento = 0
                NumColumna = Origen.Column
                Salida = False
                Do While (NumElemento <= 4) And Not Salida
                    If InStr(1, Origen.Offset(0, -NumElemento).Formula, "getCurrencies(") > 0 Then
                        NumElemento = NumElemento + 1
                    Else
                        Salida = True
                    End If
                    If NumColumna = 1 Then
                        Salida = True
                    Else
                        NumColumna = NumColumna - 1
                    End If
               Loop
                If Salida Then
                    getCurrencies = Resultado(1, NumElemento)
                Else
                    getCurrencies = CVErr(xlErrNA)
                End If
            End If
        Else
            getCurrencies = RespuestaJSON(Key)
        End If
    Else
        getCurrencies = "Error retrievig currency."
    End If
End Function

Function getUnderlyings(ByVal IdTicker As String, Optional ByVal Key As String = "") As Variant
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim NumElemento As Long
    Dim NumColumna As Long
    Dim Origen As Range
    Dim Salida As Boolean
    
    If Galleta = "" Then Call LoginError
    
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getUnderlyings", "ticker", IdTicker)
    
    If Respuesta = "" Then Exit Function
    
    If InStr(1, Respuesta, """object"":""underlying""") > 0 Then
        If Key = "" Then
            ReDim Resultado(1 To 1, 1 To 4)
            Resultado(1, 1) = RespuestaJSON("murexCode")
            Resultado(1, 2) = RespuestaJSON("calendar")
            Resultado(1, 3) = RespuestaJSON("currency")
            Resultado(1, 4) = RespuestaJSON("validCalendar")
            If Application.Caller.HasArray Then
                getUnderlyings = Resultado
            Else
                Set Origen = Application.Caller
                NumElemento = 0
                NumColumna = Origen.Column
                Salida = False
                Do While (NumElemento <= 4) And Not Salida
                    If InStr(1, Origen.Offset(0, -NumElemento).Formula, "getUnderlyings(") > 0 Then
                        NumElemento = NumElemento + 1
                    Else
                        Salida = True
                    End If
                    If NumColumna = 1 Then
                        Salida = True
                    Else
                        NumColumna = NumColumna - 1
                    End If
               Loop
                If Salida Then
                    getUnderlyings = Resultado(1, NumElemento)
                Else
                    getUnderlyings = CVErr(xlErrNA)
                End If
            End If
        Else
            getUnderlyings = RespuestaJSON(Key)
        End If
    Else
        getUnderlyings = "Error retrievig underlying."
    End If
End Function

Function getMurexCode(ByVal IdTicker As String, sHoja As String, iCol As Integer) As String
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    
    If Galleta = "" Then Call LoginError
    
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getUnderlyings", "ticker", IdTicker)
    
    If Respuesta = "" Then Exit Function
    
    If InStr(1, Respuesta, """object"":""underlying""") > 0 Then
        getMurexCode = RespuestaJSON("murexCode")
    Else
        getMurexCode = "Error retrievig underlying."
        If VBA.InStr(1, sHoja, "Bulk", vbTextCompare) > 0 Then
            ThisWorkbook.Sheets(sHoja).Range("GI_InstrumentId_V").Offset(0, iCol).Value = getMurexCode
        End If
    End If
End Function

Function getCalendar(ByVal IdTicker As String) As String
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    
    If Galleta = "" Then Call LoginError
    
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getUnderlyings", "ticker", IdTicker)
    
    If Respuesta = "" Then Exit Function
    
    If InStr(1, Respuesta, """object"":""underlying""") > 0 Then
        getCalendar = RespuestaJSON("calendar")
    Else
        getCalendar = "Error retrievig underlying."
        MsgBox "Error retrievig underlying.", vbCritical + vbOKOnly, "ERROR UNDERLYING"
    End If
End Function

Function getHolidaysArray(ByVal IdCalendar As String) As Variant
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim Resultado As Variant
    Dim NumElemento As Long
    
    If Galleta = "" Then Call LoginError
    
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getHolidays", "calendar", IdCalendar)
    
    If Respuesta = "" Then Exit Function
    
    If InStr(1, Respuesta, """object"":""calendar""") > 0 Then
        ReDim Resultado(1 To 100, 1)
        For NumElemento = 1 To 100
            Resultado(NumElemento, 1) = ""
            On Error Resume Next
            Resultado(NumElemento, 1) = FechaHolidays(RespuestaJSON("dates")(NumElemento))
            On Error GoTo 0
        Next NumElemento
        getHolidaysArray = Resultado
    Else
        getHolidaysArray = "Error retrievig underlying."
    End If
End Function

Function FechaHolidays(ByVal FechaTexto As String) As Date
    Dim Dia As Long
    Dim Mes As Long
    Dim Año As Long
    
    Dia = Mid(FechaTexto, 1, 2)
    Mes = Mid(FechaTexto, 4, 2)
    Año = Mid(FechaTexto, 7, 4)
    
    FechaHolidays = DateSerial(Año, Mes, Dia)
End Function

Function EncodeBase64(ByVal text$)
    Dim B
    With CreateObject("ADODB.Stream")
        .Open: .Type = 2: .Charset = "utf-8"
        .WriteText text: .Position = 0: .Type = 1: B = .Read
        With CreateObject("Microsoft.XMLDOM").createElement("b64")
            .DataType = "bin.base64": .nodeTypedValue = B
            EncodeBase64 = Replace(Mid(.text, 5), vbLf, "")
        End With
        .Close
    End With
End Function

Function DecodeBase64(ByVal b64$)
    Dim B
    With CreateObject("Microsoft.XMLDOM").createElement("b64")
        .DataType = "bin.base64": .text = b64
        B = .nodeTypedValue
        With CreateObject("ADODB.Stream")
            .Open: .Type = 1: .Write B: .Position = 0: .Type = 2: .Charset = "utf-8"
            DecodeBase64 = .ReadText
            .Close
        End With
    End With
End Function

Function ConcatenatePricer(ParamArray Rangos()) As String
    Dim Celda As Range
    Dim NumRango As Long
    
    If bConcatenate = True Then Exit Function
    
    If UBound(Rangos) > 1 Then
        ConcatenatePricer = "Function only admits 1 or 2 ranges."
        Exit Function
    End If

    If Rangos(0).Columns.Count > 1 Then
        ConcatenatePricer = "Ranges must be column type."
        Exit Function
    End If
    
    If UBound(Rangos) = 1 Then
        If Rangos(1).Columns.Count > 1 Then
            ConcatenatePricer = "Ranges must be column type."
            Exit Function
        End If
    End If
    
    If UBound(Rangos) = 1 Then
        If Rangos(0).row <> Rangos(1).row Then
            ConcatenatePricer = "Ranges must be start at same row."
            Exit Function
        End If
    End If
        
    If UBound(Rangos) = 1 Then
        If Rangos(0).Rows.Count <> Rangos(1).Rows.Count Then
            ConcatenatePricer = "Ranges must be the same size."
            Exit Function
        End If
    End If
        
    ConcatenatePricer = ""
    If UBound(Rangos) = 0 Then
        For Each Celda In Rangos(0)
            If VBA.IsError(Celda) = True Then Exit For
            If Celda <> "" Then
                If IsNumeric(Celda.Value) Then
                    ConcatenatePricer = ConcatenatePricer & Replace(Celda.Value & "", ",", ".") & "; "
                ElseIf IsDate(Celda.Value) Then
                    ConcatenatePricer = ConcatenatePricer & Format(Celda.Value, "yyyy-mm-dd") & "; "
                Else
                    ConcatenatePricer = ConcatenatePricer & Celda.Value & "; "
                End If
            End If
        Next Celda
    Else
        For Each Celda In Rangos(0)
            If Celda <> "" And Rangos(1).Cells(1.1).Offset(Celda.row - Rangos(0).row, 0) Then
                If IsNumeric(Celda.Value) Then
                    ConcatenatePricer = ConcatenatePricer & Replace(Celda.Value & "", ",", ".") & "; "
                ElseIf IsDate(Celda.Value) Then
                    ConcatenatePricer = ConcatenatePricer & Format(Celda.Value, "yyyy-mm-dd") & "; "
                Else
                    ConcatenatePricer = ConcatenatePricer & Celda.Value & "; "
                End If
            End If
        Next Celda
    End If

    If VBA.Trim(ConcatenatePricer) <> "" Then ConcatenatePricer = Left(ConcatenatePricer, Len(ConcatenatePricer) - 2)
    
End Function

Sub EscribeMatriz(ByRef Matriz() As String, ByVal NumFila As Long, ByVal Titulo As String, ByVal valor As Variant, Optional ByVal Nivel As String = "0")
    Matriz(NumFila, 1) = Titulo
    If IsNumeric(valor) Then
        Matriz(NumFila, 2) = Replace(valor & "", ",", ".")
    ElseIf IsDate(valor) Then
        Matriz(NumFila, 2) = Format(valor, "yyyy-mm-dd")
    Else
        Matriz(NumFila, 2) = valor
    End If
    Matriz(NumFila, 3) = Nivel
End Sub

Function Mapea(sHoja As String, iCol As Integer, sProduct As String)
    Select Case sProduct
        Case "Autocall"
            Mapea = MapeaAutocall(sHoja, iCol)
            
    End Select
End Function

Sub ResetFormulas(sHoja As String, iCol As Integer, sProduct As String)
    Select Case sProduct
        Case "Autocall"
            If VBA.Left(sHoja, 4) <> "Bulk" Then
                Call ResetFormulasAutocall(sHoja, iCol)
            Else
                bInsert = True
                Call DefaultDataAutocallBulk(sHoja, iCol)
                Call ResetFormulasAutocallBulk(sHoja, iCol)
                bInsert = False
            End If
            
    End Select
End Sub

Sub GetResults(sHoja As String, iCol As Integer, sProduct As String, Respuesta As String, RespuestaJSON As Object)
    Select Case sProduct
        Case "Autocall"
            If VBA.Left(sHoja, 4) <> "Bulk" Then
                Call GetResultsAutocall(sHoja, iCol, Respuesta, RespuestaJSON)
            Else
                Call GetResultsAutocallBulk(sHoja, iCol, Respuesta, RespuestaJSON)
            End If
            
    End Select
End Sub

Sub LoadUnderlyings(sHoja As String)
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim NumElemento As Long
    Dim UltFila As Integer
        
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getUnderlyings", "ticker", "")
    
    If Respuesta = "" Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    
    Worksheets(sHoja).cmbUnderlyings.ListFillRange = ""
    
    With Worksheets("Underlyings")
        .Columns("A:A").Clear
        If InStr(1, Respuesta, """object"":""list""") > 0 Then
            For NumElemento = 1 To RespuestaJSON("content").Count
                .Cells(NumElemento + 1, 1) = RespuestaJSON("content")(NumElemento)
            Next NumElemento
            
            With .Sort
                .SortFields.Clear
                .SortFields.Add Cells(1, 1), xlSortOnValues, xlAscending, xlSortNormal
                .SetRange Worksheets("Underlyings").Columns(1)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        
        If ThisWorkbook.Sheets(sHoja).Range("pricerEnvironment").Value = "PREproduction" Then
            .Columns("B:B").Copy
            .Columns("A:A").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
        
        IsArrow = True
        UltFila = .Range("A" & .Rows.Count).End(xlUp).row
        Worksheets(sHoja).cmbUnderlyings.ListFillRange = "=Underlyings!$A$1:$A$" & UltFila
    End With
        
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub LoadClients(sHoja As String)
    Dim Respuesta As String
    Dim RespuestaJSON As Object
    Dim NumElemento As Long
    Dim UltFila As Integer
            
    Set RespuestaJSON = Services.OtherServices(Respuesta, "getClients", "", "")
    
    If Respuesta = "" Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    
    With ThisWorkbook.Worksheets("Clients")
        .Cells.Clear
        If InStr(1, Respuesta, """object"":""list""") > 0 Then
            For NumElemento = 1 To RespuestaJSON("content").Count
                .Cells(NumElemento + 1, 1) = RespuestaJSON("content")(NumElemento)
            Next NumElemento
            With .Sort
                .SortFields.Clear
                .SortFields.Add Cells(1, 1), xlSortOnValues, xlAscending, xlSortNormal
                .SetRange ThisWorkbook.Worksheets("Clients").Columns(1)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        
        UltFila = .Range("A" & .Rows.Count).End(xlUp).row
    End With
    
    With ThisWorkbook.Worksheets(sHoja)
        With .Range("GI_Client_V").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Clients!$A$1:$A$" & UltFila
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "ERROR"
            .InputMessage = ""
            .errorMessage = "Clients does not exist"
            .ShowInput = True
            .ShowError = True
        End With
    End With
    
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub ButtonLoadStaticData(sHoja As String)
    Call ProtectSheet(False, sHoja)
    If Galleta = "" Then Call LoginError
    
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
    Call LoadUnderlyings(sHoja)
    Application.Cursor = xlDefault
    If Err.Number <> 0 Then
        Application.Cursor = xlDefault
        MsgBox "Vpn not connected (Error loading underlyings)", vbCritical + vbOKOnly, "ERROR"
        Exit Sub
    Else
        Application.Cursor = xlWait
        Call LoadClients(sHoja)
        Application.Cursor = xlDefault
    End If
    Call ProtectSheet(True, sHoja)
End Sub

Sub ButtonLoadFMM(sHoja As String)
    Dim sFichero As String
    Dim DocumentoXML As DOMDocument
    Dim Nodo1 As IXMLDOMNode
    Dim Nodo2 As IXMLDOMNode
    Dim Nodo3 As IXMLDOMNode
    Dim Nodo4 As IXMLDOMNode
    Dim Nodo5 As IXMLDOMNode
    Dim Nodo6 As IXMLDOMNode
    Dim Nodo7 As IXMLDOMNode
    Dim Nodo8 As IXMLDOMNode
    Dim Nodo9 As IXMLDOMNode
    Dim Nodo10 As IXMLDOMNode
    Dim Nodo11 As IXMLDOMNode
    
    bLoadFMM = True

    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Open FFM file"
        .Filters.Add "Files xml (*.xml)", "*.xml", 1
        .FilterIndex = 1
        .AllowMultiSelect = False
        On Error GoTo ERRORES
        If .Show = -1 Then
            sFichero = .SelectedItems(1)
            DoEvents
                
            Set DocumentoXML = New DOMDocument
            
            DocumentoXML.Load (sFichero)
            
            With ThisWorkbook.Sheets(sHoja)
                Call ProtectSheet(False, sHoja)
                Application.Calculation = xlCalculationManual
                .Range("GI_InstrumentId_V").Value = ""
                Call ProtectSheet(False, sHoja)
                .Range("S_ObservationDates_V").Value = ""
                .Range("APB_Dates1").Value = ""
                'Range("B_ObservationDates_V").Value = ""'se hace en el propio nodo porque salta evento change de altexpiry
                'SE QUITA EL BLOQUE COUPON
                'Range("O41:O44").Value = ""
                .Range("APB_Dates2").Value = ""
                For Each Nodo1 In DocumentoXML.ChildNodes
                    For Each Nodo2 In Nodo1.ChildNodes
                        For Each Nodo3 In Nodo2.ChildNodes
                            If Nodo3.HasChildNodes Then
                                Select Case Nodo3.BaseName
                                    Case "sentBy"
                                        .Range("GI_Client_V").Value = Nodo3.nodeTypedValue
                                        
                                    Case "tradeHeader"
                                        For Each Nodo4 In Nodo3.ChildNodes
                                            If Nodo4.HasChildNodes Then
                                                Select Case Nodo4.BaseName
                                                    Case "tradeDate"
                                                        .Range("GI_TradeDate_V").Value = Nodo4.nodeTypedValue
                                                End Select
                                            End If
                                        Next Nodo4
                                        
                                    Case "exoticEquityOption"
                                        For Each Nodo4 In Nodo3.ChildNodes
                                            If Nodo4.HasChildNodes Then
                                                Select Case Nodo4.BaseName
                                                    Case "productType"
                                                        .Range("GI_ProductType_V").Value = Nodo4.nodeTypedValue
                                                    Case "mode"
                                                        .Range("GI_Mode_V").Value = Nodo4.nodeTypedValue
                                                    Case "effectiveDate"
                                                        .Range("GI_EffectiveDate_V").Value = Nodo4.nodeTypedValue
                                                    Case "expiryDate"
                                                        .Range("GI_ExpiryDate_V").Value = Nodo4.nodeTypedValue
                                                    Case "valueDate"
                                                        .Range("GI_ValueDate_V").Value = Nodo4.nodeTypedValue
                                                    Case "notional"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "currency"
                                                                        .Range("GI_Currency_V").Value = Nodo5.nodeTypedValue
                                                                    Case "amount"
                                                                        .Range("GI_Amount_V").Value = Nodo5.nodeTypedValue
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    Case "underlyer"
                                                        For Each Nodo5 In Nodo4.ChildNodes 'basket
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "basket"
                                                                        For Each Nodo6 In Nodo5.ChildNodes 'basket
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "basketConstituent"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes 'basketConstituent
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "equity"
                                                                                                        For Each Nodo8 In Nodo7.ChildNodes 'equity
                                                                                                            If Nodo8.HasChildNodes Then
                                                                                                                Select Case Nodo8.BaseName
                                                                                                                    Case "instrumentId"
                                                                                                                        If VBA.Trim(.Range("GI_InstrumentId_V").Value) <> "" Then
                                                                                                                            .Range("GI_InstrumentId_V").Value = Range("GI_InstrumentId_V").Value & "; " & Nodo8.nodeTypedValue
                                                                                                                        Else
                                                                                                                            .Range("GI_InstrumentId_V").Value = Nodo8.nodeTypedValue
                                                                                                                        End If
                                                                                                                End Select
                                                                                                            End If
                                                                                                        Next Nodo8
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                    Case "basketId"
                                                                                        .Range("GI_BasketId_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "basketType"
                                                                                        .Range("GI_BasketType_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    Case "option"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "optionType"
                                                                        .Range("O_OptionType_V").Value = Nodo5.nodeTypedValue
                                                                    Case "effectiveDate"
                                                                        .Range("O_EffectiveDate_V").Value = Nodo5.nodeTypedValue
                                                                    Case "expiryDate"
                                                                        .Range("O_ExpiryDate_V").Value = Nodo5.nodeTypedValue
                                                                    Case "valueDate"
                                                                        .Range("O_ValueDate_V").Value = Nodo5.nodeTypedValue
                                                                    Case "priceDefinition"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "type"
                                                                                        .Range("O_PriceDefinitionType_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "observationDates"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "date"
                                                                                                        .Range("O_ObservationDates_V").Value = Nodo7.nodeTypedValue
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "strike"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "strikePrice"
                                                                                        .Range("O_StrikePrice_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "payoff"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "type"
                                                                                        .Range("O_PayOffType_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "optionFactor"
                                                                                        .Range("O_OptionFactor_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "leverageFactor"
                                                                                        .Range("O_LeverageFactor_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "floor"
                                                                                        .Range("O_Floor_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "leveraged"
                                                                        .Range("O_Leveraged_V").NumberFormat = "@"
                                                                        .Range("O_Leveraged_V").Value = Nodo5.nodeTypedValue
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    Case "barrier"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "barrierType"
                                                                        .Range("B_BarrierType_V").Value = Nodo5.nodeTypedValue
                                                                    Case "direction"
                                                                        .Range("B_Direction_V").Value = Nodo5.nodeTypedValue
                                                                    Case "observationType"
                                                                        .Range("B_ObservationType_V").Value = Nodo5.nodeTypedValue
                                                                    Case "observationDates"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "date"
                                                                                        .Range("B_ObservationDates_V").Value = ""
                                                                                        If VBA.IsError(.Range("B_ObservationDates_V").Value) = True Then .Range("B_ObservationDates_V").Value = ""
                                                                                        If VBA.Trim(.Range("B_ObservationDates_V").Value) <> "" Then
                                                                                            .Range("B_ObservationDates_V").Value = .Range("B_ObservationDates_V").Value & "; " & Nodo6.nodeTypedValue
                                                                                        Else
                                                                                            .Range("B_ObservationDates_V").Value = Nodo6.nodeTypedValue
                                                                                        End If
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "triggerRate"
                                                                        .Range("B_TriggerRate_V").Value = Nodo5.nodeTypedValue
                                                                    Case "costOfHedge"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "type"
                                                                                        .Range("B_CostOfHedgeType_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "delta"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "type"
                                                                                                        .Range("B_DeltaType_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "value"
                                                                                                        .Range("B_DeltaValue_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "cap"
                                                                                                        .Range("B_DeltaCap_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "floor"
                                                                                                        .Range("B_DeltaFloor_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "liquidityAlpha"
                                                                                                        .Range("B_DeltaLiquidityAlpha_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "maxDelta"
                                                                                                        .Range("B_MaxDelta_V").NumberFormat = "@"
                                                                                                        .Range("B_MaxDelta_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "maxDeltaValue"
                                                                                                        .Range("B_MaxDeltaValue_V").Value = Nodo7.nodeTypedValue
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    Case "strikeDefinition"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "type"
                                                                        .Range("S_StrikeDefinitionType_V").Value = Nodo5.nodeTypedValue
                                                                    Case "schedule"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "startDate"
                                                                                        .Range("S_StartDate_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "endDate"
                                                                                        .Range("S_EndDate_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "observationDates"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "date"
                                                                                        .Range("S_ObservationDates_V").NumberFormat = "@"
                                                                                        If VBA.Trim(.Range("S_ObservationDates_V").Value) <> "" Then
                                                                                            .Range("S_ObservationDates_V").Value = .Range("S_ObservationDates_V").Value & "; " & Nodo6.nodeTypedValue
                                                                                        Else
                                                                                            .Range("S_ObservationDates_V").Value = Nodo6.nodeTypedValue
                                                                                        End If
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                End Select
                                                           End If
                                                        Next Nodo5
                                                    Case "earlyRedemption"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "earlyRedemptionParameters"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "frequency"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "periodMultiplier"
                                                                                                        .Range("ER_PeriodMultiplier_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "period"
                                                                                                        .Range("ER_Frequency_V").Value = Nodo7.nodeTypedValue
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                    Case "initialTriggerRate"
                                                                                        .Range("ER_InitialTriggerRate_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "initialTriggerPayment"
                                                                                        .Range("ER_InitialTriggerPayment_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "triggerStepPayment"
                                                                                        .Range("ER_TriggerStepPayment_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "initialNoTriggerPayment"
                                                                                        .Range("ER_InitialNoTriggerPayment_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "nonCancelablePeriods"
                                                                                        .Range("ER_NonCancelablePeriods_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "earlyRedemptionPeriodSchedule"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "earlyRedemptionPeriod"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "fixingDate"
                                                                                                        .Range("ER_FixingDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("ER_FixingDates_V").Value) <> "" Then
                                                                                                            .Range("ER_FixingDates_V").Value = .Range("ER_FixingDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("ER_FixingDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "settlementDate"
                                                                                                        .Range("ER_SettlementDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("ER_SettlementDates_V").Value) <> "" Then
                                                                                                            .Range("ER_SettlementDates_V").Value = .Range("ER_SettlementDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("ER_SettlementDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "payoff"
                                                                                                        For Each Nodo8 In Nodo7.ChildNodes
                                                                                                            If Nodo8.HasChildNodes Then
                                                                                                                Select Case Nodo8.BaseName
                                                                                                                    Case "trigger"
                                                                                                                        For Each Nodo9 In Nodo8.ChildNodes
                                                                                                                            If Nodo9.HasChildNodes Then
                                                                                                                                Select Case Nodo9.BaseName
                                                                                                                                    Case "triggerRate"
                                                                                                                                        If VBA.Trim(.Range("ER_TriggerRates_V").Value) <> "" Then
                                                                                                                                            .Range("ER_TriggerRates_V").Value = .Range("ER_TriggerRates_V").Value & "; " & Nodo9.nodeTypedValue
                                                                                                                                        Else
                                                                                                                                            .Range("ER_TriggerRates_V").Value = Nodo9.nodeTypedValue
                                                                                                                                        End If
                                                                                                                                    Case "triggerPayment"
                                                                                                                                        If VBA.Trim(.Range("ER_TriggerPayments_V").Value) <> "" Then
                                                                                                                                            .Range("ER_TriggerPayments_V").Value = .Range("ER_TriggerPayments_V").Value & "; " & Nodo9.nodeTypedValue
                                                                                                                                        Else
                                                                                                                                            .Range("ER_TriggerPayments_V").Value = Nodo9.nodeTypedValue
                                                                                                                                        End If
                                                                                                                                    Case "noTriggerPayment"
                                                                                                                                        If VBA.Trim(.Range("ER_NoTriggerPayments_V").Value) <> "" Then
                                                                                                                                            .Range("ER_NoTriggerPayments_V").Value = .Range("ER_NoTriggerPayments_V").Value & "; " & Nodo9.nodeTypedValue
                                                                                                                                        Else
                                                                                                                                            .Range("ER_NoTriggerPayments_V").Value = Nodo9.nodeTypedValue
                                                                                                                                        End If
                                                                                                                                End Select
                                                                                                                            End If
                                                                                                                        Next Nodo9
                                                                                                                End Select
                                                                                                            End If
                                                                                                        Next Nodo8
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    'se quita el bloque coupon
    '                                                    Case "coupon"
    '                                                        For Each Nodo5 In Nodo4.ChildNodes
    '                                                            If Nodo5.HasChildNodes Then
    '                                                                Select Case Nodo5.BaseName
    '                                                                    Case "couponParameters"
    '                                                                        For Each Nodo6 In Nodo5.ChildNodes
    '                                                                            If Nodo6.HasChildNodes Then
    '                                                                                Select Case Nodo6.BaseName
    '                                                                                    Case "frequency"
    '                                                                                        For Each Nodo7 In Nodo6.ChildNodes
    '                                                                                            If Nodo7.HasChildNodes Then
    '                                                                                                Select Case Nodo7.BaseName
    '                                                                                                    Case "periodMultiplier"
    '                                                                                                        Range("ER_NoTriggerPayments_V").Value = Nodo7.nodeTypedValue
    '                                                                                                    Case "period"
    '                                                                                                        Range("O36").Value = Nodo7.nodeTypedValue
    '                                                                                                End Select
    '                                                                                            End If
    '                                                                                        Next Nodo7
    '                                                                                    Case "initialTriggerRate"
    '                                                                                        Range("O37").Value = Nodo6.nodeTypedValue
    '                                                                                    Case "initialTriggerPayment"
    '                                                                                        Range("O38").Value = Nodo6.nodeTypedValue
    '                                                                                    Case "triggerStepPayment"
    '                                                                                        Range("O39").Value = Nodo6.nodeTypedValue
    '                                                                                End Select
    '                                                                            End If
    '                                                                        Next Nodo6
    '                                                                    Case "couponPeriodSchedule"
    '                                                                        For Each Nodo6 In Nodo5.ChildNodes
    '                                                                            If Nodo6.HasChildNodes Then
    '                                                                                Select Case Nodo6.BaseName
    '                                                                                    Case "type"
    '                                                                                        Range("O40").Value = Nodo6.nodeTypedValue
    '                                                                                    Case "couponPeriod"
    '                                                                                        For Each Nodo7 In Nodo6.ChildNodes
    '                                                                                            If Nodo7.HasChildNodes Then
    '                                                                                                Select Case Nodo7.BaseName
    '                                                                                                    Case "fixingDate"
    '                                                                                                        Range("O41").NumberFormat = "@"
    '                                                                                                        If VBA.Trim(Range("O41").Value) <> "" Then
    '                                                                                                            Range("O41").Value = Range("O41").Value & "; " & Nodo7.nodeTypedValue ' VBA.Right(Nodo9.nodeTypedValue, 2) & "/" & VBA.Mid(Nodo9.nodeTypedValue, 6, 2) & "/" & VBA.Left(Nodo9.nodeTypedValue, 4)
    '                                                                                                        Else
    '                                                                                                            Range("O41").Value = Nodo7.nodeTypedValue ' VBA.Right(Nodo9.nodeTypedValue, 2) & "/" & VBA.Mid(Nodo9.nodeTypedValue, 6, 2) & "/" & VBA.Left(Nodo9.nodeTypedValue, 4)
    '                                                                                                        End If
    '                                                                                                    Case "settlementDate"
    '                                                                                                        Range("O42").NumberFormat = "@"
    '                                                                                                        If VBA.Trim(Range("O42").Value) <> "" Then
    '                                                                                                            Range("O42").Value = Range("O42").Value & "; " & Nodo7.nodeTypedValue ' VBA.Right(Nodo9.nodeTypedValue, 2) & "/" & VBA.Mid(Nodo9.nodeTypedValue, 6, 2) & "/" & VBA.Left(Nodo9.nodeTypedValue, 4)
    '                                                                                                        Else
    '                                                                                                            Range("O42").Value = Nodo7.nodeTypedValue ' VBA.Right(Nodo9.nodeTypedValue, 2) & "/" & VBA.Mid(Nodo9.nodeTypedValue, 6, 2) & "/" & VBA.Left(Nodo9.nodeTypedValue, 4)
    '                                                                                                        End If
    '                                                                                                    Case "payoff"
    '                                                                                                        For Each Nodo8 In Nodo7.ChildNodes
    '                                                                                                            If Nodo8.HasChildNodes Then
    '                                                                                                                Select Case Nodo8.BaseName
    '                                                                                                                    Case "trigger"
    '                                                                                                                        For Each Nodo9 In Nodo8.ChildNodes
    '                                                                                                                            If Nodo9.HasChildNodes Then
    '                                                                                                                                Select Case Nodo9.BaseName
    '                                                                                                                                    Case "triggerRate"
    '                                                                                                                                        If VBA.Trim(Range("O43").Value) <> "" Then
    '                                                                                                                                            Range("O43").Value = Range("O43").Value & "; " & Nodo9.nodeTypedValue
    '                                                                                                                                        Else
    '                                                                                                                                            Range("O43").Value = Nodo9.nodeTypedValue
    '                                                                                                                                        End If
    '                                                                                                                                    Case "triggerPayment"
    '                                                                                                                                        If VBA.Trim(Range("O44").Value) <> "" Then
    '                                                                                                                                            Range("O44").Value = Range("O44").Value & "; " & Nodo9.nodeTypedValue
    '                                                                                                                                        Else
    '                                                                                                                                            Range("O44").Value = Nodo9.nodeTypedValue
    '                                                                                                                                        End If
    '                                                                                                                                End Select
    '                                                                                                                            End If
    '                                                                                                                        Next Nodo9
    '                                                                                                                End Select
    '                                                                                                            End If
    '                                                                                                        Next Nodo8
    '                                                                                                End Select
    '                                                                                            End If
    '                                                                                        Next Nodo7
    '                                                                                End Select
    '                                                                            End If
    '                                                                        Next Nodo6
    '                                                                End Select
    '                                                            End If
    '                                                        Next Nodo5
                                                    Case "interestLeg"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "interestCalculation"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "floatingRateCalculation"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "floatingRateIndex"
                                                                                                        .Range("I_RateIndex_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "indexTenor"
                                                                                                        For Each Nodo8 In Nodo7.ChildNodes
                                                                                                            If Nodo8.HasChildNodes Then
                                                                                                                Select Case Nodo8.BaseName
                                                                                                                    Case "periodMultiplier"
                                                                                                                        .Range("I_PeriodMultiplier_V").Value = Nodo8.nodeTypedValue
                                                                                                                    Case "period"
                                                                                                                        .Range("I_Period_V").Value = Nodo8.nodeTypedValue
                                                                                                                End Select
                                                                                                            End If
                                                                                                        Next Nodo8
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                    Case "dayCountFraction"
                                                                                        .Range("I_DayCountFraction_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "accrued"
                                                                                        .Range("I_Accrued_V").NumberFormat = "@"
                                                                                        .Range("I_Accrued_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "frequency"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "periodMultiplier"
                                                                                                        .Range("I_PeriodMultiplier2_V").Value = Nodo7.nodeTypedValue
                                                                                                    Case "period"
                                                                                                        .Range("I_Period2_V").Value = Nodo7.nodeTypedValue
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                    Case "type"
                                                                                        .Range("I_Type_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "exchangeNotional"
                                                                                        .Range("I_ExchangeNotional_V").NumberFormat = "@"
                                                                                        .Range("I_ExchangeNotional_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "interestLegPeriodSchedule"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "interestLegPeriod"
                                                                                        For Each Nodo7 In Nodo6.ChildNodes
                                                                                            If Nodo7.HasChildNodes Then
                                                                                                Select Case Nodo7.BaseName
                                                                                                    Case "accrualStartDate"
                                                                                                        .Range("I_AccrualStartDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("I_AccrualStartDates_V").Value) <> "" Then
                                                                                                            .Range("I_AccrualStartDates_V").Value = .Range("I_AccrualStartDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("I_AccrualStartDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "accrualEndDate"
                                                                                                        .Range("I_AccrualEndDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("I_AccrualEndDates_V").Value) <> "" Then
                                                                                                            .Range("I_AccrualEndDates_V").Value = .Range("I_AccrualEndDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("I_AccrualEndDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "fixingDate"
                                                                                                        .Range("I_FixingDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("I_FixingDates_V").Value) <> "" Then
                                                                                                            .Range("I_FixingDates_V").Value = .Range("I_FixingDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("I_FixingDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "settlementDate"
                                                                                                        .Range("I_SettlementDates_V").NumberFormat = "@"
                                                                                                        If VBA.Trim(.Range("I_SettlementDates_V").Value) <> "" Then
                                                                                                            .Range("I_SettlementDates_V").Value = .Range("I_SettlementDates_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("I_SettlementDates_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                    Case "spreadValue"
                                                                                                        If VBA.Trim(.Range("I_SpreadValues_V").Value) <> "" Then
                                                                                                            .Range("I_SpreadValues_V").Value = .Range("I_SpreadValues_V").Value & "; " & Nodo7.nodeTypedValue
                                                                                                        Else
                                                                                                            .Range("I_SpreadValues_V").Value = Nodo7.nodeTypedValue
                                                                                                        End If
                                                                                                End Select
                                                                                            End If
                                                                                        Next Nodo7
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                    Case "costOfHedge"
                                                        For Each Nodo5 In Nodo4.ChildNodes
                                                            If Nodo5.HasChildNodes Then
                                                                Select Case Nodo5.BaseName
                                                                    Case "type"
                                                                        .Range("CH_Type_V").Value = Nodo5.nodeTypedValue
                                                                    Case "delta"
                                                                        For Each Nodo6 In Nodo5.ChildNodes
                                                                            If Nodo6.HasChildNodes Then
                                                                                Select Case Nodo6.BaseName
                                                                                    Case "type"
                                                                                        .Range("CH_DeltaType_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "floor"
                                                                                        .Range("CH_DeltaFloor_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "liquidityAlpha"
                                                                                        .Range("CH_DeltaLiquidityAlpha_V").Value = Nodo6.nodeTypedValue
                                                                                    Case "asianTailFactor"
                                                                                        .Range("CH_DeltaAsianTailFactor_V").Value = Nodo6.nodeTypedValue
                                                                                End Select
                                                                            End If
                                                                        Next Nodo6
                                                                    Case "jumps"
                                                                        .Range("CH_Jumps_V").Value = Nodo5.nodeTypedValue
                                                                End Select
                                                            End If
                                                        Next Nodo5
                                                End Select
                                            End If
                                        Next Nodo4
                                End Select
                            End If
                        Next Nodo3
                    Next Nodo2
                Next Nodo1
                Application.Calculation = xlCalculationAutomatic
                Call ProtectSheet(True, sHoja)
                
                bLoadFMM = False
                
                MsgBox "FMM file loaded correctly", vbInformation + vbOKOnly, "FMM FILE"
            End With
        End If
    End With
    
    Exit Sub
    
ERRORES:
    If Err.Number = 5 Then
        Exit Sub
    End If
End Sub

Sub ButtonResetFormulas(sHoja As String, iCol As Integer, sProduct As String)
    If Galleta = "" Then Call LoginError
    
    Application.ScreenUpdating = False
    
    bReset = True
    
    Call ResetFormulas(sHoja, iCol, sProduct)
    
    bReset = False
End Sub

Sub ResetButtons(sHoja As String)
    With ThisWorkbook.Sheets(sHoja)
        .ButtonLogin.BackColor = &H8000000F
        .ButtonLogout.BackColor = &H8000000F
        .ButtonLoadDates.BackColor = &H8000000F
        .ButtonGenerateFMM.BackColor = &H8000000F
        .ButtonCalculatePrice.BackColor = &H8000000F
        .ButtonGenFMMCalcPrice.BackColor = &H8000000F
        .ButtonGetResult.BackColor = &H8000000F
        .ButtonRetrieveXMLs.BackColor = &H8000000F
        .ButtonEditFMM.BackColor = &H8000000F
        .ButtonLoadStaticData.BackColor = &H8000000F
        .ButtonLoadFMM.BackColor = &H8000000F
        .ButtonResetFormulas.BackColor = &H8000000F
    End With
End Sub

Sub ProtectSheet(bProtect As Boolean, sHoja As String)
    On Error GoTo ERRORES
    Application.EnableCancelKey = xlErrorHandler
    Select Case bProtect
        Case True
            With ThisWorkbook
                .Sheets(sHoja).Protect "BBVA"
            End With
        Case False
            With ThisWorkbook
                .Sheets(sHoja).Unprotect "BBVA"
            End With
    End Select
    Exit Sub
ERRORES:
    If Err.Number = 18 Then Resume Next
End Sub

Sub FormatCells(sHoja As String, iCol As Integer, sProduct As String)
    Select Case sProduct
        Case "Autocall"
            Call FormatCellsAutocall(sHoja, iCol)
            
    End Select
End Sub

Sub FormatDatesCells(sHoja As String, iCol As Integer, sProduct As String)
    Select Case sProduct
        Case "Autocall"
            Call FormatDatesCellsAutocall(sHoja, iCol)
            
    End Select
End Sub

Sub CalculateSwapSpread(sHoja As String, iCol As Integer)
    Dim i As Integer
    Dim NumFila As Integer
    Dim mFixingDates As Variant
    Dim mEquityLeg As Variant
    Dim mEquitySwapLeg As Variant
    Dim mGenerateAutocallable As Variant
    Dim UltFila As Integer
    Dim iDiffMonths As Integer
    Dim iMinDiff As Integer
    Dim iMonth As Integer
    Dim sPeriod As String
    Dim mData As Variant

    With ThisWorkbook.Sheets("Aux")
        .Columns(sColMonths & ":" & sColEMTN2).Clear
    End With
    
    With ThisWorkbook.Sheets(sHoja)
        mFixingDates = Application.Run("QBS.DateGen.Param.FixingDates", VBA.CLng(.Range("AS_StartDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("AS_EndDate_V").Offset(0, iCol).Value), .Range("AS_Frequency_V").Offset(0, iCol).Value)
        mEquityLeg = Application.Run("QBS.DateGen.Param.EquityLeg", mFixingDates, VBA.CLng(.Range("D_InitialPaymentDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("D_FinalAlignmentDate_V").Offset(0, iCol).Value), .Range("EL_Frequency_V").Offset(0, iCol).Value, .Range("EL_PaymentLag_V").Offset(0, iCol).Value, .Range("EL_FixingCalendar_V").Offset(0, iCol).Value, .Range("EL_PaymentCalendar_V").Offset(0, iCol).Value, .Range("EL_Alignment_V").Offset(0, iCol).Value, .Range("EL_BrokenPeriod_V").Offset(0, iCol).Value, .Range("EL_FixingAdjustment_V").Offset(0, iCol).Value, .Range("EL_PaymentAdjustment_V").Offset(0, iCol).Value, .Range("EL_StickToMothEnd_V").Offset(0, iCol).Value, .Range("EL_AdjustInputDates_V").Offset(0, iCol).Value)
        mEquitySwapLeg = Application.Run("QBS.DateGen.Param.EquitySwapLeg", .Range("ESL_SwapFrequency_V").Offset(0, iCol).Value, .Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value, .Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value, "2B", .Range("ESL_SwapAlignment_V").Offset(0, iCol).Value, .Range("ESL_SwapBrokenPeriod_V").Offset(0, iCol).Value, .Range("ESL_SwapPaymentAdjustment_V").Offset(0, iCol).Value, VBA.CLng(.Range("ESL_SwapStartDate_V").Offset(0, iCol).Value), VBA.CLng(.Range("ESL_SwapEndDate_V").Offset(0, iCol).Value), .Range("ESL_SwapStickToMothEnd_V").Offset(0, iCol).Value, IIf(VBA.IsEmpty(.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value) = True, False, .Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value))
        mGenerateAutocallable = Application.Run("QBS.DateGen.GenerateAutocallable", mEquityLeg, mEquitySwapLeg, .Range("A_EarlyRedemptionFreq_V").Offset(0, iCol).Value, .Range("A_FirstEarlyRedemptionPer_V").Offset(0, iCol).Value, .Range("A_EarlyRedemptionAlignment_V").Offset(0, iCol).Value, .Range("A_AllowEarlyRedemptionMat_V").Offset(0, iCol).Value, IIf(VBA.IsEmpty(.Range("A_Dates_V").Offset(0, iCol).Value) = True, False, .Range("A_Dates_V").Offset(0, iCol).Value), True, True)
    
        ReDim mData2(1 To 100, 1 To 3)
        UltFila = .Range(sColSwapStartDates & .Rows.Count).End(xlUp).row
        i = 1
        If UltFila >= 3 Then
            For NumFila = 3 To UltFila
                iDiffMonths = VBA.DateDiff("m", .Range(sColSwapStartDates & NumFila).Value, .Range(sColSwapEndDates & NumFila).Value)
                If NumFila = 3 Then iMinDiff = iDiffMonths
                Select Case iDiffMonths
                    Case Is < 3
                        mData2(i, 1) = 1
                    Case Is < 6
                        mData2(i, 1) = 3
                    Case Else
                        mData2(i, 1) = 6
                End Select
                If i = 1 Then
                    mData2(i, 2) = mData2(i, 1)
                ElseIf i > 1 Then
                    mData2(i, 2) = mData2((i - 1), 2) + mData2(i, 1)
                End If
                If mData2(i, 2) < 12 Then
                    sPeriod = mData2(i, 2) & "m"
                ElseIf mData2(i, 2) = 12 Then 'igual al año
                    sPeriod = VBA.Int(mData2(i, 2) / 12) & "y"
                Else 'MÁS QUE UN AÑO
                    sPeriod = VBA.Replace(VBA.Int(mData2(i, 2) / 12) & "y" & IIf(VBA.Int(mData2(i, 2) Mod 12) = 0, "", VBA.Int(mData2(i, 2) Mod 12) & "m"), "y0m", "")
                End If
                ThisWorkbook.Sheets("Aux").Range(sColEMTN2 & NumFila).Formula = "=IFERROR(VLOOKUP(" & """" & sPeriod & """" & ", " & sColMonthsEMTNRAR & "1:" & sColEMTN & "121, 2, FALSE), 0)"
                If iDiffMonths < iMinDiff Then iMinDiff = iDiffMonths
                i = i + 1
            Next NumFila
        End If
        
        Select Case iMinDiff
            Case Is < 3
                iMonth = 1
            Case Is < 6
                iMonth = 3
            Case Else
                iMonth = 6
        End Select
        
        With ThisWorkbook.Sheets("Aux")
            .Range(sColMonthsEMTNRAR & "1:" & sColEMTN2 & "121").NumberFormat = "General"
            If ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value = "EUR" Then
                .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").FormulaArray = "=bbva_GetTyMatrix(""REFERENCE"",""IR_NOTE_SPREAD"",TODAY()," & VBA.Chr(34) & VBA.Trim(VBA.Replace(ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "EUR", "") & " MTN " & iMonth) & "m Callable Spread"")"
            Else 'cualquier otra divisa
                .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").FormulaArray = "=bbva_GetTyMatrix(""REFERENCE"",""IR_NOTE_SPREAD"",TODAY()," & VBA.Chr(34) & ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value & " MTN 1m Callable Spread"")"
            End If
            If .Range(sColMonthsEMTNRAR & "1").Value = "You are not logged" Then
                MsgBox "You are not logged in Typhoon Add-Ins", vbCritical + vbOKOnly, "ERROR LOGIN"
                Exit Sub
            ElseIf VBA.InStr(1, .Range(sColMonthsEMTNRAR & "1").Value, "ERROR", vbTextCompare) > 0 Then
                MsgBox .Range(sColMonthsEMTNRAR & "1").Value, vbCritical + vbOKOnly, "ERROR"
                Exit Sub
            End If
            .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").Copy
            .Range(sColMonthsEMTNRAR & "1:" & sColRAR & "121").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            
            UltFila = .Range(sColEMTN2 & .Rows.Count).End(xlUp).row
            mData = .Range(sColEMTN2 & "3:" & sColEMTN2 & UltFila).Value
            If UltFila >= 3 Then
                For i = 1 To (UltFila - 2)
                    mData(i, 1) = mData(i, 1) / 10000
                Next i
            End If
        End With
        If VBA.Left(sHoja, 4) <> "Bulk" Then _
            .Range(sColSwapSpread & "3:" & sColSwapSpread & UltFila).Value = mData
    End With
End Sub
