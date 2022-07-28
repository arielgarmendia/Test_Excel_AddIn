using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pricer.ExcelAddIn.Modules
{
    internal class ExcelPricer
    {
        class _failedMemberConversionMarker1
        {
        }
#error Cannot convert OptionStatementSyntax - see comment for details
        /* Cannot convert OptionStatementSyntax, CONVERSION ERROR: Conversion for OptionStatement not implemented, please report this issue in 'Option ' at character 25


        Input:
            Option OnBase 1

         */
        class _failedMemberConversionMarker2
        {
        }
#error Cannot convert OptionStatementSyntax - see comment for details
        /* Cannot convert OptionStatementSyntax, CONVERSION ERROR: Conversion for OptionStatement not implemented, please report this issue in 'Option Explicit On' at character 42


        Input:
        Option Explicit On

         */
        class _failedMemberConversionMarker3
        {
        }

        private bool bLoadDatesBulk;
        public bool bGuardar;

        public bool IsArrow;

        // Errores
        public const string STRING_ERROR = "ERROR";

        public string Galleta;

        // colors
        public const long lBlue = 16247773L;
        public const long lRed = 255L;
        public const long lOrange = 49407L;
        public const long lLightGrayAPB = 15921906L; // autocall product block
        public const long lDarkGrayAPB = 10921638L;
        public const long lLightGrayDGB = 15592941L; // date generator block
        public const long lDarkGrayDGB = 13224393L;

        // LVB Proced. Generator
        public const string sColEquityInitialFixingDates = "V";
        public const string sColEquityFixingDates = "W";
        public const string sColEquityPaymentDates = "X";
        public const string sColEarlyRedemption = "Y";
        public const string sColEarlyRedemptionTrigger = "Z";
        public const string sColACCoupon = "AA";
        public const string sColNonCallCoupon = "AB";
        public const string sColSwapFixingDates = "AE";
        public const string sColSwapStartDates = "AF";
        public const string sColSwapEndDates = "AG";
        public const string sColSwapPaymentDates = "AH";
        public const string sColSwapSpread = "AI";
        public const string sColBarrierObservationDates = "AK";
        public const string sColValores1 = "O";
        public const string sColValores2 = "S";
        // Aux
        public const string sColMonthsEMTNRAR = "AA"; // typhoon
        public const string sColEMTN = "AB"; // typhoon
        public const string sColRAR = "AC"; // typhoon
        public const string sColMonths = "AD"; // 
        public const string sColSumMonths = "AE";
        public const string sColEMTN2 = "AF";


        public bool bMapea;
        public bool bReset;
        public bool bLoadFMM;
        public bool bInsert; // para insertar o copiar columna bulk
        public bool bCopy; // copy to bulk
        public bool bClone; // clone bulk
        public bool bConcatenate;

        public string sB_BarrierType_V;
        public string sB_Direction_V;
        public double sB_DeltaCap_V;
        public double sB_DeltaFloor_V;
        public double sB_DeltaLiquidityAlpha_V;
        public string sB_MaxDelta_V;
        public double sB_MaxDeltaValue_V;
        public string sAc_BarrierType_V;
        public string sAc_BarrierLevel_V;
        public string sAc_BarrierShiftType_V;


        public void Auto_Open()
        {
            Application.Calculation = xlCalculationAutomatic;
        }

        public string PeticionHTTP(object ObjetoHTTP, string Peticion)
        {
            string PeticionHTTPRet = default;
            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            ObjetoHTTP.SetRequestHeader("Cookie", Galleta);
            ;
            ObjetoHTTP.Send(Peticion);
            if (Information.Err().Number != 0)
            {
                Interaction.MsgBox(Information.Err().Description, (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR");
                // End
                return PeticionHTTPRet;
            };

            PeticionHTTPRet = Conversions.ToString(ObjetoHTTP.responseText);
            return PeticionHTTPRet;
        }

        public void ButtonLoadDates(string sHoja, int iCol, string sProduct)
        {
            long NumFila;
            Variant Instruments;
            var sCal = default(string);
            var sCalendars = default(string);
            var mDatesCal = default(Variant);
            var mDatesCurrencies = default(Variant);
            Variant mSetCalendar;
            string sCurrency;
            Variant mData1;
            Variant mData2;
            Variant mData;
            int UltFila;
            Variant mFixingDates;
            Variant mEquityLeg;
            Variant mEquitySwapLeg;
            Variant mGenerateAutocallable;
            int iDiffMonths;
            var iMinDiff = default(int);
            int iMonth;
            string sPeriod;
            int i;
            var Rango = default(Range);
            Variant mColBarrierObservationDates;

            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            ProtectSheet(false, sHoja);
            ;

            if (VBA.Left(sHoja, 4) != "Bulk")
            {
                {
                    var withBlock = ThisWorkbook.Sheets(sHoja);
                    withBlock.Range("Ac_Underlying_V").Offset(0, iCol).Formula = "=UPPER(ConcatenatePricer(Underlyings))";
                    withBlock.Range("GI_InstrumentId_V").Offset(0, iCol).Value = "=Ac_Underlying_V";
                    withBlock.Range("DatesAcCouponsBlock").Value = "";
                    withBlock.Range("DatesSwapSpreadBlock").Value = "";
                }
            }

            ProtectSheet(false, sHoja);

            FormatDatesCells(sHoja, iCol, sProduct);

            Application.EnableEvents = false;

            {
                var withBlock1 = ThisWorkbook.Sheets(sHoja);
                Instruments = Strings.Split(withBlock1.Range("GI_InstrumentId_V").Offset(0, iCol).Value, ";");

                var loopTo = (long)Information.UBound(Instruments);
                for (NumFila = 0L; NumFila <= loopTo; NumFila++)
                {
                    sCal = getCalendar(Strings.Trim(Instruments[NumFila]));
                    if (VBA.InStr(1, sCalendars, sCal, Constants.vbTextCompare) == 0)
                        sCalendars = sCalendars + "+" + sCal;
                }
                sCalendars = VBA.Mid(sCalendars, 2);
                withBlock1.Range("EL_FixingCalendar_V").Offset(0, iCol).Formula = sCalendars;
                withBlock1.Range("BOD_FixingCalendar_V").Offset(0, iCol).Formula = sCalendars;

                sCurrency = getCurrencies(Worksheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "PaymentCalendar");
                withBlock1.Range("EL_PaymentCalendar_V").Offset(0, iCol).Value = sCurrency;
                withBlock1.Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value = sCurrency;
                withBlock1.Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value = sCurrency;

                for (NumFila = 0L; NumFila <= 0L; NumFila++)
                {
                    mDatesCal = getHolidaysArray(sCal);
                    mDatesCurrencies = getHolidaysArray(sCurrency);
                }

                {
                    var withBlock2 = ThisWorkbook.Sheets("Aux");
                    withBlock2.Range("AG2:AG" + (Information.UBound(mDatesCal) + 1)).Value = mDatesCal;
                    ;
                    mSetCalendar = Application.Run("QBS.DateGen.SetCalendar", sCalendars, Rango);

                    withBlock2.Range("AI2:AI" + (Information.UBound(mDatesCurrencies) + 1)).Value = mDatesCurrencies;
                    ;
                    mSetCalendar = Application.Run("QBS.DateGen.SetCalendar", sCurrency, Rango);
                }

                if (VBA.Left(sHoja, 4) != "Bulk")
                {
                    {
                        var withBlock3 = withBlock1.Range("O_ValueDate_V").Offset(0, iCol);
                        withBlock3.Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,\"TARGET\")";
                        withBlock3.Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,EL_PaymentCalendar_V)";
                    }
                    {
                        var withBlock4 = withBlock1.Range("ESL_SwapEndDate_V").Offset(0, iCol);
                        withBlock4.Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,\"TARGET\")";
                        withBlock4.Formula = "=QBS.DateGen.AddPeriod(Ac_ExpiryDate_V,Ac_PaymentShifter_V,EL_PaymentCalendar_V)";
                    }
                }
                else // bulk
                {
                    withBlock1.Range("O_ValueDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" + withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Address + "," + withBlock1.Range("Ac_PaymentShifter_V").Offset(0, iCol).Address + ",\"TARGET\")";
                    withBlock1.Range("O_ValueDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" + withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Address + "," + withBlock1.Range("Ac_PaymentShifter_V").Offset(0, iCol).Address + "," + withBlock1.Range("EL_PaymentCalendar_V").Offset(0, iCol).Address + ")";

                    withBlock1.Range("ESL_SwapEndDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" + withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Address + "," + withBlock1.Range("Ac_PaymentShifter_V").Offset(0, iCol).Address + ",\"TARGET\")";
                    withBlock1.Range("ESL_SwapEndDate_V").Offset(0, iCol).Formula = "=QBS.DateGen.AddPeriod(" + withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Address + "," + withBlock1.Range("Ac_PaymentShifter_V").Offset(0, iCol).Address + "," + withBlock1.Range("EL_PaymentCalendar_V").Offset(0, iCol).Address + ")";
                }

                mFixingDates = Application.Run("QBS.DateGen.Param.FixingDates", VBA.CLng(withBlock1.Range("AS_StartDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("AS_EndDate_V").Offset(0, iCol).Value), withBlock1.Range("AS_Frequency_V").Offset(0, iCol).Value);
                mEquityLeg = Application.Run("QBS.DateGen.Param.EquityLeg", mFixingDates, VBA.CLng(withBlock1.Range("D_InitialPaymentDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("D_FinalAlignmentDate_V").Offset(0, iCol).Value), withBlock1.Range("EL_Frequency_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentLag_V").Offset(0, iCol).Value, withBlock1.Range("EL_FixingCalendar_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentCalendar_V").Offset(0, iCol).Value, withBlock1.Range("EL_Alignment_V").Offset(0, iCol).Value, withBlock1.Range("EL_BrokenPeriod_V").Offset(0, iCol).Value, withBlock1.Range("EL_FixingAdjustment_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentAdjustment_V").Offset(0, iCol).Value, withBlock1.Range("EL_StickToMothEnd_V").Offset(0, iCol).Value, withBlock1.Range("EL_AdjustInputDates_V").Offset(0, iCol).Value);
                mEquitySwapLeg = Application.Run("QBS.DateGen.Param.EquitySwapLeg", withBlock1.Range("ESL_SwapFrequency_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value, "2B", withBlock1.Range("ESL_SwapAlignment_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapBrokenPeriod_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapPaymentAdjustment_V").Offset(0, iCol).Value, VBA.CLng(withBlock1.Range("ESL_SwapStartDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("ESL_SwapEndDate_V").Offset(0, iCol).Value), withBlock1.Range("ESL_SwapStickToMothEnd_V").Offset(0, iCol).Value, Interaction.IIf(VBA.IsEmpty(withBlock1.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value) == true, false, withBlock1.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value));
                mGenerateAutocallable = Application.Run("QBS.DateGen.GenerateAutocallable", mEquityLeg, mEquitySwapLeg, withBlock1.Range("A_EarlyRedemptionFreq_V").Offset(0, iCol).Value, withBlock1.Range("A_FirstEarlyRedemptionPer_V").Offset(0, iCol).Value, withBlock1.Range("A_EarlyRedemptionAlignment_V").Offset(0, iCol).Value, withBlock1.Range("A_AllowEarlyRedemptionMat_V").Offset(0, iCol).Value, Interaction.IIf(VBA.IsEmpty(withBlock1.Range("A_Dates_V").Offset(0, iCol).Value) == true, false, withBlock1.Range("A_Dates_V").Offset(0, iCol).Value), true, true);
                ;
                if (VBA.InStr(1, mGenerateAutocallable, "Error:", Constants.vbTextCompare) > 0 & VBA.Left(sHoja, 4) == "Bulk")
                {
                    if (Information.Err().Number == 0)
                    {
                        withBlock1.Range("Error_V").Offset(0, iCol).Value = mGenerateAutocallable;
                        return;
                    }
                }
                else if (VBA.InStr(1, mGenerateAutocallable, "Error:", Constants.vbTextCompare) > 0 & VBA.Left(sHoja, 4) != "Bulk")
                {
                    if (Information.Err().Number == 0)
                    {
                        Interaction.MsgBox(mGenerateAutocallable, (int)Constants.vbCritical + (int)Constants.vbOKOnly, "ERROR");
                        return;
                    }
                };
#error Cannot convert OnErrorGoToStatementSyntax - see comment for details
                /* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo ERRORES' at character 11566


                Input:
                            On Error GoTo ERRORES

                 */
                ;
#error Cannot convert ReDimStatementSyntax - see comment for details
                /* Cannot convert ReDimStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Symbols.Metadata.PE.PENamedTypeSymbolWithEmittedNamespaceName' to type 'Microsoft.CodeAnalysis.IArrayTypeSymbol'.
                   at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.CreateNewArrayAssignment(ExpressionSyntax vbArrayExpression, ExpressionSyntax csArrayExpression, List`1 convertedBounds)
                   at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<ConvertRedimClauseAsync>d__41.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<<VisitReDimStatement>b__40_0>d.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectAsync>d__3`2.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectManyAsync>d__0`2.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<VisitReDimStatement>d__40.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.CSharp.PerScopeStateVisitorDecorator.<AddLocalVariablesAsync>d__6.MoveNext()
                --- End of stack trace from previous location where exception was thrown ---
                   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.<DefaultVisitInnerAsync>d__3.MoveNext()

                Input:

                            ReDim mData1(1 To 10000, 1 To 3)

                 */
                ;
                var loopTo1 = (long)Information.UBound(mGenerateAutocallable);
                for (NumFila = 1L; NumFila <= loopTo1; NumFila++)
                {
                    mData1(NumFila, 1) = withBlock1.Range("ER_InitialTriggerRate_V").Offset(0, iCol).Value;
                    // AC Coupon (%)
                    switch (withBlock1.Range("Ac_ACCoupon_V").Offset(0, iCol).Value)
                    {
                        case "Flat": // valor fijo
                            {
                                mData1(NumFila, 2) = withBlock1.Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value;
                                break;
                            }
                        case "Coupon Step": // se va incrementando valor
                            {
                                if (NumFila == 2L)
                                {
                                    mData1(NumFila, 2) = withBlock1.Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value;
                                }
                                else
                                {
                                    mData1(NumFila, 2) = withBlock1.Range(sColACCoupon + NumFila) + withBlock1.Range("Ac_ACCouponPorc_V").Offset(0, iCol).Value;
                                }

                                break;
                            }
                    }
                    mData1(NumFila, 3) = withBlock1.Range("Ac_NonCallCoupon_V").Offset(0, iCol).Value;

                    // swap spread
                    if (NumFila < (long)Information.UBound(mGenerateAutocallable))
                    {
                        if (mGenerateAutocallable(NumFila + 1L, 7) != "")
                        {
                            iDiffMonths = VBA.DateDiff("m", VBA.CDate(mGenerateAutocallable(NumFila + 1L, 7)), VBA.CDate(mGenerateAutocallable(NumFila + 1L, 8)));
                            if (NumFila == 1L)
                                iMinDiff = iDiffMonths;
                            switch (iDiffMonths)
                            {
                                case var @case when @case < 3:
                                    {
                                        mData2(NumFila, 1) = 1;
                                        break;
                                    }
                                case var case1 when case1 < 6:
                                    {
                                        mData2(NumFila, 1) = 3;
                                        break;
                                    }

                                default:
                                    {
                                        mData2(NumFila, 1) = 6;
                                        break;
                                    }
                            }
                            if (NumFila == 1L)
                            {
                                mData2(NumFila, 2) = mData2(NumFila, 1);
                            }
                            else if (NumFila > 1L)
                            {
                                mData2(NumFila, 2) = mData2(NumFila - 1L, 2) + mData2(NumFila, 1);
                            }
                            if (mData2(NumFila, 2) < 12)
                            {
                                sPeriod = mData2(NumFila, 2) + "m";
                            }
                            else if (mData2(NumFila, 2) == 12) // igual al año
                            {
                                sPeriod = VBA.Int(mData2(NumFila, 2) / 12) + "y";
                            }
                            else // MÁS QUE UN AÑO
                            {
                                sPeriod = VBA.Replace(VBA.Int(mData2(NumFila, 2) / 12) + "y" + Interaction.IIf(VBA.Int(mData2(NumFila, 2) % 12) == 0, "", VBA.Int(mData2(NumFila, 2) % 12) + "m"), "y0m", "");
                            }
                            ThisWorkbook.Sheets("Aux").Range(sColEMTN2 + (NumFila + 2L)).Formula = "=IFERROR(VLOOKUP(" + "\"" + sPeriod + "\"" + ", " + sColMonthsEMTNRAR + "1:" + sColEMTN + "121, 2, FALSE), 0)";
                            if (iDiffMonths < iMinDiff)
                                iMinDiff = iDiffMonths;
                        }
                    }
                }

                switch (iMinDiff)
                {
                    case var case2 when case2 < 3:
                        {
                            iMonth = 1;
                            break;
                        }
                    case var case3 when case3 < 6:
                        {
                            iMonth = 3;
                            break;
                        }

                    default:
                        {
                            iMonth = 6;
                            break;
                        }
                }

                {
                    var withBlock5 = ThisWorkbook.Sheets("Aux");
                    withBlock5.Range(sColMonthsEMTNRAR + "1:" + sColEMTN2 + "121").NumberFormat = "General";
                    if (ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value == "EUR")
                    {
                        withBlock5.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").FormulaArray = "=bbva_GetTyMatrix(\"REFERENCE\",\"IR_NOTE_SPREAD\",TODAY()," + VBA.Chr(34) + VBA.Trim(VBA.Replace(ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "EUR", "") + " MTN " + iMonth) + "m Callable Spread\")";
                    }
                    else // cualquier otra divisa
                    {
                        withBlock5.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").FormulaArray = "=bbva_GetTyMatrix(\"REFERENCE\",\"IR_NOTE_SPREAD\",TODAY()," + VBA.Chr(34) + ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value + " MTN 1m Callable Spread\")";
                    }
                    if (withBlock5.Range(sColMonthsEMTNRAR + "1").Value == "You are not logged")
                    {
                        Interaction.MsgBox("You are not logged in Typhoon Add-Ins", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR LOGIN");
                        return;
                    }
                    else if (VBA.InStr(1, withBlock5.Range(sColMonthsEMTNRAR + "1").Value, "ERROR", Constants.vbTextCompare) > 0)
                    {
                        Interaction.MsgBox.Range(sColMonthsEMTNRAR + "1").Value(default, (int)Constants.vbCritical + (int)Constants.vbOKOnly, "ERROR");
                        return;
                    }
                    withBlock5.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").Copy();
                    withBlock5.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").PasteSpecial(xlPasteValues);
                    Application.CutCopyMode = false;

                    UltFila = withBlock5.Range(sColEMTN2 + withBlock5.Rows.Count).End(xlUp).row;
                    mData = withBlock5.Range(sColEMTN2 + "3:" + sColEMTN2 + UltFila).Value;
                    if (UltFila >= 3)
                    {
                        var loopTo2 = UltFila - 2;
                        for (i = 1; i <= loopTo2; i++)
                            mData(i, 1) = mData(i, 1) / 10000;
                    }
                }
                // ****

                var loopTo3 = withBlock1.Range("Ac_NonCancelPeriods_V").Offset(0, iCol).Value;
                for (NumFila = 1L; NumFila <= loopTo3; NumFila++)
                {
                    mData1(NumFila, 1) = 99.99d;
                    mData1(NumFila, 3) = 0;
                }

                if (VBA.Left(sHoja, 4) != "Bulk")
                {
                    Application.Calculation = xlCalculationManual;
                    ;
                    var loopTo4 = (long)Information.UBound(mGenerateAutocallable);
                    for (NumFila = 2L; NumFila <= loopTo4; NumFila++) // - 1
                    {
                        if (mGenerateAutocallable(NumFila, 1) != "")
                            withBlock1.Range(sColEquityInitialFixingDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 1));
                        withBlock1.Range(sColEquityFixingDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila + 1L, 2));
                        withBlock1.Range(sColEquityPaymentDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila + 1L, 3));
                        withBlock1.Range(sColEarlyRedemption + (NumFila + 1L)).Value = mGenerateAutocallable(NumFila + 1L, 5);
                        withBlock1.Range(sColSwapFixingDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 6));
                        withBlock1.Range(sColSwapStartDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 7));
                        withBlock1.Range(sColSwapEndDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 8));
                        withBlock1.Range(sColSwapPaymentDates + (NumFila + 1L)).Value = VBA.CDate(mGenerateAutocallable(NumFila, 9));
                    };

                    UltFila = withBlock1.Range(sColEquityFixingDates + withBlock1.Rows.Count).End(xlUp).row;
                    withBlock1.Range(sColEarlyRedemptionTrigger + "3:" + sColNonCallCoupon + UltFila).Value = mData1;

                    UltFila = withBlock1.Range(sColSwapStartDates + withBlock1.Rows.Count).End(xlUp).row;
                    withBlock1.Range(sColSwapSpread + "3:" + sColSwapSpread + UltFila).Value = mData;

                    UltFila = withBlock1.Range(sColBarrierObservationDates + withBlock1.Rows.Count).End(xlUp).row;
                    if (UltFila >= 3)
                        withBlock1.Range(sColBarrierObservationDates + "3:" + sColBarrierObservationDates + UltFila).Formula = "";
                    switch (withBlock1.Range("B_ObservationType_V").Offset(0, iCol).Value)
                    {
                        case "AtExpiry":
                            {
                                withBlock1.Range(sColBarrierObservationDates + "3").Formula = "=Ac_ExpiryDate_V";
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "CallSpread";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.015d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                break;
                            }
                        case "Continuous":
                            {
                                withBlock1.Range(sColBarrierObservationDates + "3").Formula = "=Ac_StrikeDate_V";
                                withBlock1.Range(sColBarrierObservationDates + "4").Formula = "=Ac_ExpiryDate_V";
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                break;
                            }
                        case "Daily":
                            {
                                // .Range(sColBarrierObservationDates & "3:" & sColBarrierObservationDates & "2000").FormulaArray = "=QBS.DateGen.GenerateSchedule(BOD_FirstObservationDate_V,BOD_LastObservationDate_V,""1B"",BOD_FixingCalendar_V,,,,,BOD_AdjustInputDates_V,TRUE)"
                                mColBarrierObservationDates = Application.Run("QBS.DateGen.GenerateSchedule", VBA.CLng(withBlock1.Range("BOD_FirstObservationDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("BOD_LastObservationDate_V").Offset(0, iCol).Value), "1B", withBlock1.Range("BOD_FixingCalendar_V").Offset(0, iCol).Value, default, default, default, default, withBlock1.Range("BOD_AdjustInputDates_V").Offset(0, iCol).Value, true);
                                withBlock1.Range(sColBarrierObservationDates + "3:" + sColBarrierObservationDates + (Information.UBound(mColBarrierObservationDates) + 2)).Value = mColBarrierObservationDates;
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                break;
                            }
                    }

                    UltFila = withBlock1.Range(sColBarrierObservationDates + withBlock1.Rows.Count).End(xlUp).row;
                    if (UltFila >= 3)
                        withBlock1.Range("B_ObservationDates_V").Formula = "=ConcatenatePricer($" + sColBarrierObservationDates + "$3:$" + sColBarrierObservationDates + "$" + UltFila + ")";
                    // colors porque es automatico
                    withBlock1.Range("B_ObservationDates_V").Interior.Color = lDarkGrayAPB;
                    withBlock1.Columns(sColBarrierObservationDates + ":" + sColBarrierObservationDates).Interior.Pattern = xlNone;

                    Application.Calculation = xlCalculationAutomatic;
                }
                else // bulk
                {
                    withBlock1.Range("S_ObservationDates_V").Offset(0, iCol).Value = "";
                    withBlock1.Range("APB_Dates1").Offset(0, iCol).Value = "";
                    withBlock1.Range("APB_Dates2").Offset(0, iCol).Value = "";
                    withBlock1.Range("Error_V").Offset(0, iCol).Value = "";
                    ;
                    var loopTo5 = (long)Information.UBound(mGenerateAutocallable);
                    for (NumFila = 1L; NumFila <= loopTo5; NumFila++) // - 2)
                    {
                        if (mGenerateAutocallable(NumFila + 1L, 1) != "")
                            withBlock1.Range("S_ObservationDates_V").Offset(0, iCol).Value = withBlock1.Range("S_ObservationDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 1L, 1)), "yyyy-mm-dd") + "; ";
                        withBlock1.Range("ER_FixingDates_V").Offset(0, iCol).Value = withBlock1.Range("ER_FixingDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 2L, 2)), "yyyy-mm-dd") + "; ";
                        withBlock1.Range("ER_SettlementDates_V").Offset(0, iCol).Value = withBlock1.Range("ER_SettlementDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 2L, 3)), "yyyy-mm-dd") + "; ";

                        if (Information.Err().Number == 9)
                            break;
                        if (mGenerateAutocallable(NumFila + 2L, 3) != "")
                        {
                            withBlock1.Range("ER_TriggerRates_V").Offset(0, iCol).Value = withBlock1.Range("ER_TriggerRates_V").Offset(0, iCol).Value + mData1(NumFila, 1) + "; ";
                            withBlock1.Range("ER_TriggerPayments_V").Offset(0, iCol).Value = withBlock1.Range("ER_TriggerPayments_V").Offset(0, iCol).Value + mData1(NumFila, 2) + "; ";
                            withBlock1.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value = withBlock1.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value + mData1(NumFila, 3) + "; ";
                        }
                    }
                    Information.Err().Clear();

                    var loopTo6 = (long)Information.UBound(mGenerateAutocallable);
                    for (NumFila = 1L; NumFila <= loopTo6; NumFila++) // - 2)
                    {
                        withBlock1.Range("I_AccrualStartDates_V").Offset(0, iCol).Value = withBlock1.Range("I_AccrualStartDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 1L, 7)), "yyyy-mm-dd") + "; ";
                        withBlock1.Range("I_AccrualEndDates_V").Offset(0, iCol).Value = withBlock1.Range("I_AccrualEndDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 1L, 8)), "yyyy-mm-dd") + "; ";
                        withBlock1.Range("I_FixingDates_V").Offset(0, iCol).Value = withBlock1.Range("I_FixingDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 1L, 6)), "yyyy-mm-dd") + "; ";
                        withBlock1.Range("I_SettlementDates_V").Offset(0, iCol).Value = withBlock1.Range("I_SettlementDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mGenerateAutocallable(NumFila + 1L, 9)), "yyyy-mm-dd") + "; ";
                        // mismo número de swap spread que fechas
                        if (Information.Err().Number == 0)
                            withBlock1.Range("I_SpreadValues_V").Offset(0, iCol).Value = withBlock1.Range("I_SpreadValues_V").Offset(0, iCol).Value + mData(NumFila, 1) + "; ";
                    };
                    {
                        var withBlock6 = withBlock1.Range("S_ObservationDates_V").Offset(0, iCol);
                        withBlock6.NumberFormat = "@";
                        withBlock6.Value = VBA.Left(withBlock6.Value, VBA.Len(withBlock6.Value) - 2);
                    }
                    withBlock1.Range("ER_FixingDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("ER_FixingDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("ER_FixingDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("ER_SettlementDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("ER_SettlementDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("ER_SettlementDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("ER_TriggerRates_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(withBlock1.Range("ER_TriggerRates_V").Offset(0, iCol).Value, ",", "."), VBA.Len(withBlock1.Range("ER_TriggerRates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("ER_TriggerPayments_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(withBlock1.Range("ER_TriggerPayments_V").Offset(0, iCol).Value, ",", "."), VBA.Len(withBlock1.Range("ER_TriggerPayments_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(withBlock1.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value, ",", "."), VBA.Len(withBlock1.Range("ER_NoTriggerPayments_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("I_AccrualStartDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("I_AccrualStartDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("I_AccrualStartDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("I_AccrualEndDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("I_AccrualEndDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("I_AccrualEndDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("I_FixingDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("I_FixingDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("I_FixingDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("I_SettlementDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("I_SettlementDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("I_SettlementDates_V").Offset(0, iCol).Value) - 2);
                    withBlock1.Range("I_SpreadValues_V").Offset(0, iCol).Value = VBA.Left(VBA.Replace(withBlock1.Range("I_SpreadValues_V").Offset(0, iCol).Value, ",", "."), VBA.Len(withBlock1.Range("I_SpreadValues_V").Offset(0, iCol).Value) - 2);

                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value = "";
                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).NumberFormat = "@";
                    switch (withBlock1.Range("B_ObservationType_V").Offset(0, iCol).Value)
                    {
                        case "AtExpiry":
                            {
                                mColBarrierObservationDates = VBA.Split(withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Value + ";", ";");
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "CallSpread";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.015d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                var loopTo7 = (long)(Information.UBound(mColBarrierObservationDates) - 1);
                                for (NumFila = 0L; NumFila <= loopTo7; NumFila++)
                                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value = withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mColBarrierObservationDates[NumFila]), "yyyy-mm-dd") + "; ";
                                break;
                            }

                        case "Continuous":
                            {
                                mColBarrierObservationDates = VBA.Split(withBlock1.Range("Ac_StrikeDate_V").Offset(0, iCol).Value + ";" + withBlock1.Range("Ac_ExpiryDate_V").Offset(0, iCol).Value + ";", ";");
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                var loopTo8 = (long)(Information.UBound(mColBarrierObservationDates) - 1);
                                for (NumFila = 0L; NumFila <= loopTo8; NumFila++)
                                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value = withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mColBarrierObservationDates[NumFila]), "yyyy-mm-dd") + "; ";
                                break;
                            }

                        case "Daily":
                            {
                                mColBarrierObservationDates = Application.Run("QBS.DateGen.GenerateSchedule", VBA.CLng(withBlock1.Range("BOD_FirstObservationDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("BOD_LastObservationDate_V").Offset(0, iCol).Value), "1B", withBlock1.Range("BOD_FixingCalendar_V").Offset(0, iCol).Value, default, default, default, default, withBlock1.Range("BOD_AdjustInputDates_V").Offset(0, iCol).Value, true);
                                withBlock1.Range("B_CostOfHedgeType_V").Offset(0, iCol).Value = "Shift";
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Value = -0.01d;
                                withBlock1.Range("B_DeltaValue_V").Offset(0, iCol).Interior.Color = lLightGrayAPB;
                                var loopTo9 = (long)Information.UBound(mColBarrierObservationDates);
                                for (NumFila = 1L; NumFila <= loopTo9; NumFila++)
                                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value = withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value + VBA.Format(VBA.CDate(mColBarrierObservationDates(NumFila, 1)), "yyyy-mm-dd") + "; ";
                                break;
                            }
                    }

                    if (VBA.Trim(withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value) != "")
                        withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value = VBA.Left(withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value, VBA.Len(withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Value) - 2);

                    // colors porque es automatico
                    withBlock1.Range("B_ObservationDates_V").Offset(0, iCol).Interior.Color = lDarkGrayAPB;
                }
            }

            Application.EnableEvents = true;
            return;

        ERRORES:
            ;

            Interaction.MsgBox(Information.Err().Description);
            Information.Err().Clear();
            LoginError();
        }

        public void ButtonGenerateFMM(string sHoja, int iCol, string sProduct)
        {
            string[] Mapa;
            string Peticion;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            FileSystemObject SistemaArchivos;
            var ArchivoTexto = default(TextStream);
            string texto;
            string sRuta;
            string sFichero;
            const string sHojaBulk = "Bulk mode";

            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            Mapa = (string[])Mapea(sHoja, iCol, sProduct);

            if ((Mapa[1, 1] ?? "") != STRING_ERROR)
            {
                Peticion = JsonConverter.ConvertToJson(ObjetoJSON(Mapa), 2);
                ;

                if (string.IsNullOrEmpty(Respuesta))
                    return;

                {
                    var withBlock = ThisWorkbook.Sheets(sHoja);
                    if (VBA.Left(sHoja, 4) == "Bulk")
                        withBlock.Range("Error_V").Offset(0, iCol).Value = "";
                    sRuta = Conversions.ToString(Interaction.IIf(VBA.Left(sHoja, 4) != "Bulk", ActiveWorkbook.Path, ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value));
                    sFichero = Conversions.ToString(Interaction.IIf(VBA.Left(sHoja, 4) != "Bulk", "fmm", ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value));

                    if (VBA.Trim(sRuta) == "")
                    {
                        ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value = "Folder of the FMM is empty.";
                        return;
                    }
                    else if (VBA.Trim(sFichero) == "")
                    {
                        ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value = "Name of the FMM is empty.";
                        return;
                    }
                };
#error Cannot convert EmptyStatementSyntax - see comment for details
                /* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 33773


                Input:

                        Set SistemaArchivos = New FileSystemObject

                 */
                ;
#error Cannot convert OnErrorResumeNextStatementSyntax - see comment for details
                /* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 33847


                Input:
                        On Error Resume Next

                 */
                ;
                if (Information.Err().Number != 0)
                {
                    if (VBA.Left(sHoja, 4) == "Bulk")
                    {
                        ThisWorkbook.Sheets(sHoja).Range("Error_V").Offset(0, iCol).Value = Information.Err().Description;
                        Information.Err().Clear();
                    }
                }
                else
                {
                    ArchivoTexto.Write(RespuestaJSON("xmlFmm"));
                    ArchivoTexto.Close();

                    texto = Conversions.ToString(DecodeBase64(Conversions.ToString(RespuestaJSON("xmlFmm"))));
                    ;
                    ArchivoTexto.Write(texto);
                    ArchivoTexto.Close();

                    ThisWorkbook.Sheets(sHoja).Range("DealID").Offset(0, iCol).ClearContents();
                }
            }
        }

        public void ButtonEditFMM()
        {
            Interaction.Shell("notepad " + ActiveWorkbook.Path + "/fmm.xml", Constants.vbNormalFocus);
        }

        public void ButtonCalculatePrice(string sHoja, int iCol)
        {
            string strFileExists;
            string strFileName;
            var Respuesta = default(string);
            string pricerEnvironment;
            var RespuestaJSON = default(object);
            const string sHojaBulk = "Bulk mode";

            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            {
                var withBlock = ThisWorkbook.Sheets(sHoja);
                Application.Calculation = xlCalculationManual;
                withBlock.Range("DealID").Offset(0, iCol).Value = "";

                strFileName = Conversions.ToString(Interaction.IIf(VBA.Left(sHoja, 4) != "Bulk", ActiveWorkbook.Path + "/fmm.xml", ThisWorkbook.Sheets(sHojaBulk).Range("FolderFMM_V").Offset(0, iCol).Value + @"\" + ThisWorkbook.Sheets(sHojaBulk).Range("NameFMM_V").Offset(0, iCol).Value + ".xml"));
                strFileExists = FileSystem.Dir(strFileName);
                if (string.IsNullOrEmpty(strFileExists))
                {
                    switch (VBA.Left(sHoja, 4))
                    {
                        case var @case when @case != "Bulk":
                            {
                                Interaction.MsgBox("Please click on generateFMM before calculating price");
                                break;
                            }

                        default:
                            {
                                withBlock.Range("DealID").Offset(0, iCol).Value = "Please click on generateFMM before calculating price";
                                break;
                            }
                    }
                }
                else
                {
                    ProtectSheet(false, sHoja);
                    pricerEnvironment = ThisWorkbook.Sheets("LVB Proced. Generator").Range("pricerEnvironment").Value;
                    ;

                    if (string.IsNullOrEmpty(Respuesta))
                        return;

                    switch (VBA.Left(sHoja, 4))
                    {
                        case var case1 when case1 != "Bulk":
                            {
                                Application.Calculation = xlCalculationManual;
                                withBlock.Range("Result").Clear();
                                if (withBlock.Range("DealID").Value == "")
                                {
                                    Interaction.MsgBox(Operators.ConcatenateObject(Operators.ConcatenateObject("Deal ID ", RespuestaJSON("qtpdId")), " obtained."), Constants.vbInformation);
                                    withBlock.Range("DealID").Value = RespuestaJSON("qtpdId");
                                }
                                else if (Operators.ConditionalCompareObjectEqual(withBlock.Range("DealID").Value, RespuestaJSON("qtpdId"), false))
                                {
                                    Interaction.MsgBox(Operators.ConcatenateObject(Operators.ConcatenateObject("Deal ID ", RespuestaJSON("qtpdId")), " confirmed."), Constants.vbInformation);
                                }
                                else
                                {
                                    Interaction.MsgBox(Operators.ConcatenateObject(Operators.ConcatenateObject("An error has occurred: Deal ID ", RespuestaJSON("qtpdId")), " received."), Constants.vbExclamation);
                                }

                                break;
                            }

                        default:
                            {
                                withBlock.Range("DealID").Offset(0, iCol).ClearContents();
                                withBlock.Range("Result").Offset(0, iCol).ClearContents();
                                if (withBlock.Range("DealID").Offset(0, iCol).Value == "")
                                {
                                    withBlock.Range("DealID").Offset(0, iCol).Value = RespuestaJSON("qtpdId");
                                }
                                else if (Operators.ConditionalCompareObjectEqual(withBlock.Range("DealID").Offset(0, iCol).Value, RespuestaJSON("qtpdId"), false))
                                {
                                    withBlock.Range("DealID").Offset(0, iCol).Value = Operators.ConcatenateObject(Operators.ConcatenateObject("Deal ID ", RespuestaJSON("qtpdId")), " confirmed.");
                                    // Else
                                    withBlock.Range("DealID").Offset(0, iCol).Value = Operators.ConcatenateObject(Operators.ConcatenateObject("An error has occurred: Deal ID ", RespuestaJSON("qtpdId")), " received.");
                                }

                                break;
                            }
                    }
                }
                Application.Calculation = xlCalculationAutomatic;
            }
        }

        public void ButtonGenerateFMMAndCalculatePrice(string sHoja, int iCol, string sProduct)
        {
            ButtonGenerateFMM(sHoja, iCol, sProduct);
            ButtonCalculatePrice(sHoja, iCol);
        }

        public void ButtonGetResult(string sHoja, int iCol, string sProduct)
        {
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            long NumGriega;
            Range Destino;

            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            {
                var withBlock = ThisWorkbook.Sheets(sHoja);
                if (IsEmpty(withBlock.Range("DealID").Offset(0, iCol).Value))
                {
                    switch (VBA.Left(sHoja, 4))
                    {
                        case var @case when @case != "Bulk":
                            {
                                Interaction.MsgBox("Deal Id must be filled. Please send to calculate the deal before trying to get results");
                                break;
                            }

                        default:
                            {
                                withBlock.Range("DealID").Offset(0, iCol).Value = "Deal Id must be filled. Please send to calculate the deal before trying to get results";
                                break;
                            }
                    }
                }
                else
                {
                    ;

                    if (string.IsNullOrEmpty(Respuesta))
                        return;

                    GetResults(sHoja, iCol, sProduct, Respuesta, RespuestaJSON);
                }
            }
        }

        public void ButtonRetrieveXMLs(string sHoja)
        {
            var ArchivoTexto = default(TextStream);
            FileSystemObject SistemaArchivos;
            object ObjetoHTTP;
            string URL;
            string Peticion;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            var objIEBrowser = default(object);

            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            {
                var withBlock = ThisWorkbook.Sheets(sHoja);
                if (IsEmpty(withBlock.Range("DealID").Value))
                {
                    Interaction.MsgBox("Deal Id must be filled. Please send to calculate the deal before trying to retrieve xmls");
                }
                else
                {
                    ;

                    if (string.IsNullOrEmpty(Respuesta))
                        return;

                    if (Strings.InStr(1, Respuesta, "\"error\":") > 0)
                    {
                        Interaction.MsgBox("Error retrieving XML.", Constants.vbExclamation);
                    }
                    else if (Strings.InStr(1, Respuesta, "\"content\":[]") > 0 | Strings.InStr(1, Respuesta, "{}") > 0)
                    {
                        Interaction.MsgBox("Xml files not found.", Constants.vbExclamation);
                    }
                    else
                    {
                        ;
#error Cannot convert EmptyStatementSyntax - see comment for details
                        /* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 41610


                        Input:
                                        Set SistemaArchivos = New FileSystemObject

                         */
                        ;
                        ArchivoTexto.Write(DecodeBase64(Conversions.ToString(RespuestaJSON("content")((object)1)("content"))));
                        ArchivoTexto.Close();
                        ;
                        ArchivoTexto.Write(DecodeBase64(Conversions.ToString(RespuestaJSON("content")((object)2)("content"))));
                        ArchivoTexto.Close();
                        ;
                        ArchivoTexto.Write(DecodeBase64(Conversions.ToString(RespuestaJSON("content")((object)3)("content"))));
                        ArchivoTexto.Close();
                        ;
                        objIEBrowser.Visible = (object)true;
                        objIEBrowser.Navigate2(ActiveWorkbook.Path + "/fd" + ThisWorkbook.Sheets(sHoja).Range("DealID").Value + ".xml");
                        objIEBrowser.Navigate2(ActiveWorkbook.Path + "/fi" + ThisWorkbook.Sheets(sHoja).Range("DealID").Value + ".xml", (object)2048);
                        objIEBrowser.Navigate2(ActiveWorkbook.Path + "/fmm" + ThisWorkbook.Sheets(sHoja).Range("DealID").Value + ".xml", (object)2048);
                        while (objIEBrowser.Busy)
                        {
                        }
                    }
                }
            }
        }

        private Dictionary ObjetoJSON(string[] Mapa)
        {
            Dictionary ObjetoJSONRet = default;
            long NumFila;

            NumFila = 1L;
            ;

            while (!string.IsNullOrEmpty(Mapa[(int)NumFila, 1]))
            {
                if (Val(Mapa[(int)NumFila, 3]) == 0)
                {
                    if (!string.IsNullOrEmpty(Mapa[(int)NumFila, 2]))
                    {
                        ObjetoJSONRet.Add(Mapa[(int)NumFila, 1], Mapa[(int)NumFila, 2]);
                    }
                    else
                    {
                        ObjetoJSONRet.Add(Mapa[(int)NumFila, 1], MatrizJSON(Mapa, NumFila + 1L));
                    }
                }
                NumFila = NumFila + 1L;
            }

            return ObjetoJSONRet;
        }

        private Variant MatrizJSON(string[] Mapa, long NumFilaInicial)
        {
            Variant MatrizJSONRet = default;
            var Matriz = default(Dictionary[]);
            var ObjetoJSON = default(Dictionary);
            long NumElementos;
            long NumFila;

            NumFila = NumFilaInicial;
            NumElementos = 0L;

            while (Val(Mapa[(int)NumFila, 3]) >= Val(Mapa[(int)NumFilaInicial, 3]))
            {
                if (Val(Mapa[(int)NumFila, 3]) == Val(Mapa[(int)NumFilaInicial, 3]))
                {
                    if ((Mapa[(int)NumFila, 1] ?? "") == (Mapa[(int)NumFilaInicial, 1] ?? ""))
                    {
                        if (NumElementos >= 1L)
                            ; NumElementos = NumElementos + 1L;
                        Array.Resize(ref Matriz, (int)(NumElementos + 1));
                        ;
                    }
                    if (!string.IsNullOrEmpty(Mapa[(int)NumFila, 2]))
                    {
                        ObjetoJSON.Add(Mapa[(int)NumFila, 1], Mapa[(int)NumFila, 2]);
                    }
                    else
                    {
                        ObjetoJSON.Add(Mapa[(int)NumFila, 1], MatrizJSON(Mapa, NumFila + 1L));
                    }
                }
                NumFila = NumFila + 1L;
            };

            MatrizJSONRet = Matriz;
            return MatrizJSONRet;
        }

        public Variant getCurrencies(string IdCurrency, string Key = "")
        {
            Variant getCurrenciesRet = default;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            long NumElemento;
            long NumColumna;
            var Origen = default(Range);
            bool Salida;

            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return getCurrenciesRet;

            if (Strings.InStr(1, Respuesta, "\"error\":\"\"") > 0)
            {
                getCurrenciesRet = "Currency not found.";
            }
            else if (Conversions.ToBoolean(Strings.InStr(1, Respuesta, "\"object\":\"currency\"")))
            {
                if (string.IsNullOrEmpty(Key))
                {
                    ;
                    Resultado(1, 1) = RespuestaJSON("PaymentCalendar");
                    Resultado(1, 2) = RespuestaJSON("IBORIndex");
                    Resultado(1, 3) = RespuestaJSON("SwapFixingCalendar");
                    Resultado(1, 4) = RespuestaJSON("SwapPaymentCalendar");
                    if (Application.Caller.HasArray)
                    {
                        getCurrenciesRet = Resultado;
                    }
                    else
                    {
                        ;
                        NumElemento = 0L;
                        NumColumna = Origen.Column;
                        Salida = false;
                        while (NumElemento <= 4L & !Salida)
                        {
                            if (InStr(1, Origen.Offset(0, -NumElemento).Formula, "getCurrencies(") > 0)
                            {
                                NumElemento = NumElemento + 1L;
                            }
                            else
                            {
                                Salida = true;
                            }
                            if (NumColumna == 1L)
                            {
                                Salida = true;
                            }
                            else
                            {
                                NumColumna = NumColumna - 1L;
                            }
                        }
                        if (Salida)
                        {
                            getCurrenciesRet = Resultado(1, NumElemento);
                        }
                        else
                        {
                            getCurrenciesRet = CVErr(xlErrNA);
                        }
                    }
                }
                else
                {
                    getCurrenciesRet = RespuestaJSON(Key);
                }
            }
            else
            {
                getCurrenciesRet = "Error retrievig currency.";
            }

            return getCurrenciesRet;
        }

        public Variant getUnderlyings(string IdTicker, string Key = "")
        {
            Variant getUnderlyingsRet = default;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            long NumElemento;
            long NumColumna;
            var Origen = default(Range);
            bool Salida;

            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return getUnderlyingsRet;

            if (Strings.InStr(1, Respuesta, "\"object\":\"underlying\"") > 0)
            {
                if (string.IsNullOrEmpty(Key))
                {
                    ;
                    Resultado(1, 1) = RespuestaJSON("murexCode");
                    Resultado(1, 2) = RespuestaJSON("calendar");
                    Resultado(1, 3) = RespuestaJSON("currency");
                    Resultado(1, 4) = RespuestaJSON("validCalendar");
                    if (Application.Caller.HasArray)
                    {
                        getUnderlyingsRet = Resultado;
                    }
                    else
                    {
                        ;
                        NumElemento = 0L;
                        NumColumna = Origen.Column;
                        Salida = false;
                        while (NumElemento <= 4L & !Salida)
                        {
                            if (InStr(1, Origen.Offset(0, -NumElemento).Formula, "getUnderlyings(") > 0)
                            {
                                NumElemento = NumElemento + 1L;
                            }
                            else
                            {
                                Salida = true;
                            }
                            if (NumColumna == 1L)
                            {
                                Salida = true;
                            }
                            else
                            {
                                NumColumna = NumColumna - 1L;
                            }
                        }
                        if (Salida)
                        {
                            getUnderlyingsRet = Resultado(1, NumElemento);
                        }
                        else
                        {
                            getUnderlyingsRet = CVErr(xlErrNA);
                        }
                    }
                }
                else
                {
                    getUnderlyingsRet = RespuestaJSON(Key);
                }
            }
            else
            {
                getUnderlyingsRet = "Error retrievig underlying.";
            }

            return getUnderlyingsRet;
        }

        public string getMurexCode(string IdTicker, string sHoja, int iCol)
        {
            string getMurexCodeRet = default;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);

            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return getMurexCodeRet;

            if (Strings.InStr(1, Respuesta, "\"object\":\"underlying\"") > 0)
            {
                getMurexCodeRet = Conversions.ToString(RespuestaJSON("murexCode"));
            }
            else
            {
                getMurexCodeRet = "Error retrievig underlying.";
                if (VBA.InStr(1, sHoja, "Bulk", Constants.vbTextCompare) > 0)
                {
                    ThisWorkbook.Sheets(sHoja).Range("GI_InstrumentId_V").Offset(0, iCol).Value = getMurexCodeRet;
                }
            }

            return getMurexCodeRet;
        }

        public string getCalendar(string IdTicker)
        {
            string getCalendarRet = default;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);

            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return getCalendarRet;

            if (Strings.InStr(1, Respuesta, "\"object\":\"underlying\"") > 0)
            {
                getCalendarRet = Conversions.ToString(RespuestaJSON("calendar"));
            }
            else
            {
                getCalendarRet = "Error retrievig underlying.";
                Interaction.MsgBox("Error retrievig underlying.", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR UNDERLYING");
            }

            return getCalendarRet;
        }

        public Variant getHolidaysArray(string IdCalendar)
        {
            Variant getHolidaysArrayRet = default;
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            Variant Resultado;
            long NumElemento;

            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return getHolidaysArrayRet;

            if (Strings.InStr(1, Respuesta, "\"object\":\"calendar\"") > 0)
            {
                ;
                for (NumElemento = 1L; NumElemento <= 100L; NumElemento++)
                {
                    Resultado(NumElemento, 1) = "";
                    ;
                    Resultado(NumElemento, 1) = FechaHolidays(Conversions.ToString(RespuestaJSON("dates")((object)NumElemento)));
                    ;
                }
                getHolidaysArrayRet = Resultado;
            }
            else
            {
                getHolidaysArrayRet = "Error retrievig underlying.";
            }

            return getHolidaysArrayRet;
        }

        public DateTime FechaHolidays(string FechaTexto)
        {
            DateTime FechaHolidaysRet = default;
            long Dia;
            long Mes;
            long Año;

            Dia = Conversions.ToLong(Strings.Mid(FechaTexto, 1, 2));
            Mes = Conversions.ToLong(Strings.Mid(FechaTexto, 4, 2));
            Año = Conversions.ToLong(Strings.Mid(FechaTexto, 7, 4));

            FechaHolidaysRet = DateAndTime.DateSerial((int)Año, (int)Mes, (int)Dia);
            return FechaHolidaysRet;
        }

        public object EncodeBase64(string text)
        {
            object EncodeBase64Ret = default;
            object B;
            {
                var withBlock = Interaction.CreateObject("ADODB.Stream");
                withBlock.Open();
                withBlock.Type = (object)2;
                withBlock.Charset = "utf-8";
                withBlock.WriteText(text);
                withBlock.Position = (object)0;
                withBlock.Type = (object)1;
                B = withBlock.Read;
                {
                    var withBlock1 = Interaction.CreateObject("Microsoft.XMLDOM").createElement("b64");
                    withBlock1.DataType = "bin.base64";
                    withBlock1.nodeTypedValue = B;
                    EncodeBase64Ret = Strings.Replace(Strings.Mid(Conversions.ToString(withBlock1.text), 5), Constants.vbLf, "");
                }
                withBlock.Close();
            }

            return EncodeBase64Ret;
        }

        public object DecodeBase64(string b64)
        {
            object DecodeBase64Ret = default;
            object B;
            {
                var withBlock = Interaction.CreateObject("Microsoft.XMLDOM").createElement("b64");
                withBlock.DataType = "bin.base64";
                withBlock.text = b64;
                B = withBlock.nodeTypedValue;
                {
                    var withBlock1 = Interaction.CreateObject("ADODB.Stream");
                    withBlock1.Open();
                    withBlock1.Type = (object)1;
                    withBlock1.Write(B);
                    withBlock1.Position = (object)0;
                    withBlock1.Type = (object)2;
                    withBlock1.Charset = "utf-8";
                    DecodeBase64Ret = withBlock1.ReadText;
                    withBlock1.Close();
                }
            }

            return DecodeBase64Ret;
        }

        public string ConcatenatePricer(params object[] Rangos)
        {
            string ConcatenatePricerRet = default;
            Range Celda;
            long NumRango;

            if (bConcatenate == true)
                return ConcatenatePricerRet;

            if (Information.UBound(Rangos) > 1)
            {
                ConcatenatePricerRet = "Function only admits 1 or 2 ranges.";
                return ConcatenatePricerRet;
            }

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Rangos[0].Columns.Count, 1, false)))
            {
                ConcatenatePricerRet = "Ranges must be column type.";
                return ConcatenatePricerRet;
            }

            if (Information.UBound(Rangos) == 1)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Rangos[1].Columns.Count, 1, false)))
                {
                    ConcatenatePricerRet = "Ranges must be column type.";
                    return ConcatenatePricerRet;
                }
            }

            if (Information.UBound(Rangos) == 1)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(Rangos[0].row, Rangos[1].row, false)))
                {
                    ConcatenatePricerRet = "Ranges must be start at same row.";
                    return ConcatenatePricerRet;
                }
            }

            if (Information.UBound(Rangos) == 1)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(Rangos[0].Rows.Count, Rangos[1].Rows.Count, false)))
                {
                    ConcatenatePricerRet = "Ranges must be the same size.";
                    return ConcatenatePricerRet;
                }
            }

            ConcatenatePricerRet = "";
            if (Information.UBound(Rangos) == 0)
            {
                foreach (Range currentCelda in (IEnumerable)Rangos[0])
                {
                    Celda = currentCelda;
                    if (VBA.IsError(Celda) == true)
                        break;
                    if (Celda != "")
                    {
                        if (Information.IsNumeric(Celda.Value))
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Strings.Replace(Celda.Value + "", ",", ".") + "; ";
                        }
                        else if (Information.IsDate(Celda.Value))
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Strings.Format(Celda.Value, "yyyy-mm-dd") + "; ";
                        }
                        else
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Celda.Value + "; ";
                        }
                    }
                }
            }
            else
            {
                foreach (Range currentCelda1 in (IEnumerable)Rangos[0])
                {
                    Celda = currentCelda1;
                    if (Celda != "" & Rangos[1].Cells((object)1.1d).Offset(Celda.row - Rangos[0].row, (object)0))
                    {
                        if (Information.IsNumeric(Celda.Value))
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Strings.Replace(Celda.Value + "", ",", ".") + "; ";
                        }
                        else if (Information.IsDate(Celda.Value))
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Strings.Format(Celda.Value, "yyyy-mm-dd") + "; ";
                        }
                        else
                        {
                            ConcatenatePricerRet = ConcatenatePricerRet + Celda.Value + "; ";
                        }
                    }
                }
            }

            if (VBA.Trim(ConcatenatePricerRet) != "")
                ConcatenatePricerRet = Strings.Left(ConcatenatePricerRet, Strings.Len(ConcatenatePricerRet) - 2);
            return ConcatenatePricerRet;

        }

        public void EscribeMatriz(ref string[] Matriz, long NumFila, string Titulo, Variant valor, string Nivel = "0")
        {
            Matriz[(int)NumFila, 1] = Titulo;
            if (Information.IsNumeric(valor))
            {
                Matriz[(int)NumFila, 2] = Strings.Replace(valor + "", ",", ".");
            }
            else if (Information.IsDate(valor))
            {
                Matriz[(int)NumFila, 2] = Strings.Format(valor, "yyyy-mm-dd");
            }
            else
            {
                Matriz[(int)NumFila, 2] = valor;
            }
            Matriz[(int)NumFila, 3] = Nivel;
        }

        public object Mapea(string sHoja, int iCol, string sProduct)
        {
            object MapeaRet = default;
            switch (sProduct ?? "")
            {
                case "Autocall":
                    {
                        MapeaRet = MapeaAutocall(sHoja, iCol);
                        break;
                    }

            }

            return MapeaRet;
        }

        public void ResetFormulas(string sHoja, int iCol, string sProduct)
        {
            switch (sProduct ?? "")
            {
                case "Autocall":
                    {
                        if (VBA.Left(sHoja, 4) != "Bulk")
                        {
                            ResetFormulasAutocall(sHoja, iCol);
                        }
                        else
                        {
                            bInsert = true;
                            DefaultDataAutocallBulk(sHoja, iCol);
                            ResetFormulasAutocallBulk(sHoja, iCol);
                            bInsert = false;
                        }

                        break;
                    }

            }
        }

        public void GetResults(string sHoja, int iCol, string sProduct, string Respuesta, object RespuestaJSON)
        {
            switch (sProduct ?? "")
            {
                case "Autocall":
                    {
                        if (VBA.Left(sHoja, 4) != "Bulk")
                        {
                            GetResultsAutocall(sHoja, iCol, Respuesta, RespuestaJSON);
                        }
                        else
                        {
                            GetResultsAutocallBulk(sHoja, iCol, Respuesta, RespuestaJSON);
                        }

                        break;
                    }

            }
        }

        public void LoadUnderlyings(string sHoja)
        {
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            long NumElemento;
            int UltFila;
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return;

            Application.Calculation = xlCalculationManual;

            Worksheets(sHoja).cmbUnderlyings.ListFillRange = "";

            {
                var withBlock = Worksheets("Underlyings");
                withBlock.Columns("A:A").Clear();
                if (Strings.InStr(1, Respuesta, "\"object\":\"list\"") > 0)
                {
                    var loopTo = Conversions.ToLong(RespuestaJSON("content").Count);
                    for (NumElemento = 1L; NumElemento <= loopTo; NumElemento++)
                        withBlock.Cells(NumElemento + 1L, 1) = RespuestaJSON("content")((object)NumElemento);

                    {
                        var withBlock1 = withBlock.Sort;
                        withBlock1.SortFields.Clear();
                        withBlock1.SortFields.Add(Cells(1, 1), xlSortOnValues, xlAscending, xlSortNormal);
                        withBlock1.SetRange(Worksheets("Underlyings").Columns(1));
                        withBlock1.Header = xlNo;
                        withBlock1.MatchCase = false;
                        withBlock1.Orientation = xlTopToBottom;
                        withBlock1.SortMethod = xlPinYin;
                        withBlock1.Apply();
                    }
                }

                if (ThisWorkbook.Sheets(sHoja).Range("pricerEnvironment").Value == "PREproduction")
                {
                    withBlock.Columns("B:B").Copy();
                    withBlock.Columns("A:A").PasteSpecial(xlPasteValues);
                    Application.CutCopyMode = false;
                }

                IsArrow = true;
                UltFila = withBlock.Range("A" + withBlock.Rows.Count).End(xlUp).row;
                Worksheets(sHoja).cmbUnderlyings.ListFillRange = "=Underlyings!$A$1:$A$" + UltFila;
            }

            Application.Calculation = xlCalculationAutomatic;
        }

        public void LoadClients(string sHoja)
        {
            var Respuesta = default(string);
            var RespuestaJSON = default(object);
            long NumElemento;
            int UltFila;
            ;

            if (string.IsNullOrEmpty(Respuesta))
                return;

            Application.Calculation = xlCalculationManual;

            {
                var withBlock = ThisWorkbook.Worksheets("Clients");
                withBlock.Cells.Clear();
                if (Strings.InStr(1, Respuesta, "\"object\":\"list\"") > 0)
                {
                    var loopTo = Conversions.ToLong(RespuestaJSON("content").Count);
                    for (NumElemento = 1L; NumElemento <= loopTo; NumElemento++)
                        withBlock.Cells(NumElemento + 1L, 1) = RespuestaJSON("content")((object)NumElemento);
                    {
                        var withBlock1 = withBlock.Sort;
                        withBlock1.SortFields.Clear();
                        withBlock1.SortFields.Add(Cells(1, 1), xlSortOnValues, xlAscending, xlSortNormal);
                        withBlock1.SetRange(ThisWorkbook.Worksheets("Clients").Columns(1));
                        withBlock1.Header = xlNo;
                        withBlock1.MatchCase = false;
                        withBlock1.Orientation = xlTopToBottom;
                        withBlock1.SortMethod = xlPinYin;
                        withBlock1.Apply();
                    }
                }

                UltFila = withBlock.Range("A" + withBlock.Rows.Count).End(xlUp).row;
            }

            {
                var withBlock2 = ThisWorkbook.Worksheets(sHoja);
                {
                    var withBlock3 = withBlock2.Range("GI_Client_V").Validation;
                    withBlock3.Delete();
                    withBlock3.Add(Type: xlValidateList, AlertStyle: xlValidAlertStop, Operator: xlBetween, Formula1: "=Clients!$A$1:$A$" + UltFila);
                    withBlock3.IgnoreBlank = true;
                    withBlock3.InCellDropdown = true;
                    withBlock3.InputTitle = "";
                    withBlock3.ErrorTitle = "ERROR";
                    withBlock3.InputMessage = "";
                    withBlock3.errorMessage = "Clients does not exist";
                    withBlock3.ShowInput = true;
                    withBlock3.ShowError = true;
                }
            }

            Application.Calculation = xlCalculationAutomatic;
        }

        public void ButtonLoadStaticData(string sHoja)
        {
            ProtectSheet(false, sHoja);
            if (string.IsNullOrEmpty(Galleta))
                LoginError();
            ;
            Application.Calculation = xlCalculationManual;
            Application.Cursor = xlWait;
            LoadUnderlyings(sHoja);
            Application.Cursor = xlDefault;
            if (Information.Err().Number != 0)
            {
                Application.Cursor = xlDefault;
                Interaction.MsgBox("Vpn not connected (Error loading underlyings)", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR");
                return;
            }
            else
            {
                Application.Cursor = xlWait;
                LoadClients(sHoja);
                Application.Cursor = xlDefault;
            }
            ProtectSheet(true, sHoja);
        }

        public void ButtonLoadFMM(string sHoja)
        {
            string sFichero;
            var DocumentoXML = default(DOMDocument);
            IXMLDOMNode Nodo4;
            IXMLDOMNode Nodo5;
            IXMLDOMNode Nodo6;
            IXMLDOMNode Nodo7;
            IXMLDOMNode Nodo8;
            IXMLDOMNode Nodo10;
            IXMLDOMNode Nodo11;

            bLoadFMM = true;

            {
                var withBlock = Application.FileDialog(msoFileDialogOpen);
                withBlock.Title = "Open FFM file";
                withBlock.Filters.Add("Files xml (*.xml)", "*.xml", 1);
                withBlock.FilterIndex = 1;
                withBlock.AllowMultiSelect = false;
                ;
                if (withBlock.Show == -1)
                {
                    sFichero = withBlock.SelectedItems(1);
                    DoEvents();
                    ;

                    DocumentoXML.Load(sFichero);

                    {
                        var withBlock1 = ThisWorkbook.Sheets(sHoja);
                        ProtectSheet(false, sHoja);
                        Application.Calculation = xlCalculationManual;
                        withBlock1.Range("GI_InstrumentId_V").Value = "";
                        ProtectSheet(false, sHoja);
                        withBlock1.Range("S_ObservationDates_V").Value = "";
                        withBlock1.Range("APB_Dates1").Value = "";
                        // Range("B_ObservationDates_V").Value = ""'se hace en el propio nodo porque salta evento change de altexpiry
                        // SE QUITA EL BLOQUE COUPON
                        // Range("O41:O44").Value = ""
                        withBlock1.Range("APB_Dates2").Value = "";
                        foreach (IXMLDOMNode Nodo1 in DocumentoXML.ChildNodes)
                        {
                            foreach (IXMLDOMNode Nodo2 in Nodo1.ChildNodes)
                            {
                                foreach (IXMLDOMNode Nodo3 in Nodo2.ChildNodes)
                                {
                                    if (Nodo3.HasChildNodes)
                                    {
                                        switch (Nodo3.BaseName)
                                        {
                                            case "sentBy":
                                                {
                                                    withBlock1.Range("GI_Client_V").Value = Nodo3.nodeTypedValue;
                                                    break;
                                                }

                                            case "tradeHeader":
                                                {
                                                    foreach (IXMLDOMNode currentNodo4 in Nodo3.ChildNodes)
                                                    {
                                                        Nodo4 = currentNodo4;
                                                        if (Nodo4.HasChildNodes)
                                                        {
                                                            switch (Nodo4.BaseName)
                                                            {
                                                                case "tradeDate":
                                                                    {
                                                                        withBlock1.Range("GI_TradeDate_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                            }
                                                        }
                                                    }

                                                    break;
                                                }

                                            case "exoticEquityOption":
                                                {
                                                    foreach (IXMLDOMNode currentNodo41 in Nodo3.ChildNodes)
                                                    {
                                                        Nodo4 = currentNodo41;
                                                        if (Nodo4.HasChildNodes)
                                                        {
                                                            switch (Nodo4.BaseName)
                                                            {
                                                                case "productType":
                                                                    {
                                                                        withBlock1.Range("GI_ProductType_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                                case "mode":
                                                                    {
                                                                        withBlock1.Range("GI_Mode_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                                case "effectiveDate":
                                                                    {
                                                                        withBlock1.Range("GI_EffectiveDate_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                                case "expiryDate":
                                                                    {
                                                                        withBlock1.Range("GI_ExpiryDate_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                                case "valueDate":
                                                                    {
                                                                        withBlock1.Range("GI_ValueDate_V").Value = Nodo4.nodeTypedValue;
                                                                        break;
                                                                    }
                                                                case "notional":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo5 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo5;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "currency":
                                                                                        {
                                                                                            withBlock1.Range("GI_Currency_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "amount":
                                                                                        {
                                                                                            withBlock1.Range("GI_Amount_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "underlyer":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo51 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo51; // basket
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "basket":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo6 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo6; // basket
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "basketConstituent":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo7 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo7; // basketConstituent
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "equity":
                                                                                                                                {
                                                                                                                                    foreach (IXMLDOMNode currentNodo8 in Nodo7.ChildNodes)
                                                                                                                                    {
                                                                                                                                        Nodo8 = currentNodo8; // equity
                                                                                                                                        if (Nodo8.HasChildNodes)
                                                                                                                                        {
                                                                                                                                            switch (Nodo8.BaseName)
                                                                                                                                            {
                                                                                                                                                case "instrumentId":
                                                                                                                                                    {
                                                                                                                                                        if (VBA.Trim(withBlock1.Range("GI_InstrumentId_V").Value) != "")
                                                                                                                                                        {
                                                                                                                                                            withBlock1.Range("GI_InstrumentId_V").Value = Range("GI_InstrumentId_V").Value + "; " + Nodo8.nodeTypedValue;
                                                                                                                                                        }
                                                                                                                                                        else
                                                                                                                                                        {
                                                                                                                                                            withBlock1.Range("GI_InstrumentId_V").Value = Nodo8.nodeTypedValue;
                                                                                                                                                        }

                                                                                                                                                        break;
                                                                                                                                                    }
                                                                                                                                            }
                                                                                                                                        }
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                        case "basketId":
                                                                                                            {
                                                                                                                withBlock1.Range("GI_BasketId_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "basketType":
                                                                                                            {
                                                                                                                withBlock1.Range("GI_BasketType_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "option":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo52 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo52;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "optionType":
                                                                                        {
                                                                                            withBlock1.Range("O_OptionType_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "effectiveDate":
                                                                                        {
                                                                                            withBlock1.Range("O_EffectiveDate_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "expiryDate":
                                                                                        {
                                                                                            withBlock1.Range("O_ExpiryDate_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "valueDate":
                                                                                        {
                                                                                            withBlock1.Range("O_ValueDate_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "priceDefinition":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo61 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo61;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "type":
                                                                                                            {
                                                                                                                withBlock1.Range("O_PriceDefinitionType_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "observationDates":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo71 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo71;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "date":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("O_ObservationDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "strike":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo62 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo62;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "strikePrice":
                                                                                                            {
                                                                                                                withBlock1.Range("O_StrikePrice_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "payoff":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo63 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo63;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "type":
                                                                                                            {
                                                                                                                withBlock1.Range("O_PayOffType_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "optionFactor":
                                                                                                            {
                                                                                                                withBlock1.Range("O_OptionFactor_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "leverageFactor":
                                                                                                            {
                                                                                                                withBlock1.Range("O_LeverageFactor_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "floor":
                                                                                                            {
                                                                                                                withBlock1.Range("O_Floor_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "leveraged":
                                                                                        {
                                                                                            withBlock1.Range("O_Leveraged_V").NumberFormat = "@";
                                                                                            withBlock1.Range("O_Leveraged_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "barrier":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo53 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo53;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "barrierType":
                                                                                        {
                                                                                            withBlock1.Range("B_BarrierType_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "direction":
                                                                                        {
                                                                                            withBlock1.Range("B_Direction_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "observationType":
                                                                                        {
                                                                                            withBlock1.Range("B_ObservationType_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "observationDates":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo64 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo64;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "date":
                                                                                                            {
                                                                                                                withBlock1.Range("B_ObservationDates_V").Value = "";
                                                                                                                if (VBA.IsError(withBlock1.Range("B_ObservationDates_V").Value) == true)
                                                                                                                    withBlock1.Range("B_ObservationDates_V").Value = "";
                                                                                                                if (VBA.Trim(withBlock1.Range("B_ObservationDates_V").Value) != "")
                                                                                                                {
                                                                                                                    withBlock1.Range("B_ObservationDates_V").Value = withBlock1.Range("B_ObservationDates_V").Value + "; " + Nodo6.nodeTypedValue;
                                                                                                                }
                                                                                                                else
                                                                                                                {
                                                                                                                    withBlock1.Range("B_ObservationDates_V").Value = Nodo6.nodeTypedValue;
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "triggerRate":
                                                                                        {
                                                                                            withBlock1.Range("B_TriggerRate_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "costOfHedge":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo65 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo65;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "type":
                                                                                                            {
                                                                                                                withBlock1.Range("B_CostOfHedgeType_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "delta":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo72 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo72;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "type":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_DeltaType_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "value":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_DeltaValue_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "cap":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_DeltaCap_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "floor":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_DeltaFloor_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "liquidityAlpha":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_DeltaLiquidityAlpha_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "maxDelta":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_MaxDelta_V").NumberFormat = "@";
                                                                                                                                    withBlock1.Range("B_MaxDelta_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "maxDeltaValue":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("B_MaxDeltaValue_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "strikeDefinition":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo54 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo54;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "type":
                                                                                        {
                                                                                            withBlock1.Range("S_StrikeDefinitionType_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "schedule":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo66 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo66;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "startDate":
                                                                                                            {
                                                                                                                withBlock1.Range("S_StartDate_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "endDate":
                                                                                                            {
                                                                                                                withBlock1.Range("S_EndDate_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "observationDates":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo67 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo67;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "date":
                                                                                                            {
                                                                                                                withBlock1.Range("S_ObservationDates_V").NumberFormat = "@";
                                                                                                                if (VBA.Trim(withBlock1.Range("S_ObservationDates_V").Value) != "")
                                                                                                                {
                                                                                                                    withBlock1.Range("S_ObservationDates_V").Value = withBlock1.Range("S_ObservationDates_V").Value + "; " + Nodo6.nodeTypedValue;
                                                                                                                }
                                                                                                                else
                                                                                                                {
                                                                                                                    withBlock1.Range("S_ObservationDates_V").Value = Nodo6.nodeTypedValue;
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "earlyRedemption":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo55 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo55;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "earlyRedemptionParameters":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo68 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo68;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "frequency":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo73 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo73;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "periodMultiplier":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("ER_PeriodMultiplier_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "period":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("ER_Frequency_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                        case "initialTriggerRate":
                                                                                                            {
                                                                                                                withBlock1.Range("ER_InitialTriggerRate_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "initialTriggerPayment":
                                                                                                            {
                                                                                                                withBlock1.Range("ER_InitialTriggerPayment_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "triggerStepPayment":
                                                                                                            {
                                                                                                                withBlock1.Range("ER_TriggerStepPayment_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "initialNoTriggerPayment":
                                                                                                            {
                                                                                                                withBlock1.Range("ER_InitialNoTriggerPayment_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "nonCancelablePeriods":
                                                                                                            {
                                                                                                                withBlock1.Range("ER_NonCancelablePeriods_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "earlyRedemptionPeriodSchedule":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo69 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo69;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "earlyRedemptionPeriod":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo74 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo74;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "fixingDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("ER_FixingDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("ER_FixingDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("ER_FixingDates_V").Value = withBlock1.Range("ER_FixingDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("ER_FixingDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "settlementDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("ER_SettlementDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("ER_SettlementDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("ER_SettlementDates_V").Value = withBlock1.Range("ER_SettlementDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("ER_SettlementDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "payoff":
                                                                                                                                {
                                                                                                                                    foreach (IXMLDOMNode currentNodo81 in Nodo7.ChildNodes)
                                                                                                                                    {
                                                                                                                                        Nodo8 = currentNodo81;
                                                                                                                                        if (Nodo8.HasChildNodes)
                                                                                                                                        {
                                                                                                                                            switch (Nodo8.BaseName)
                                                                                                                                            {
                                                                                                                                                case "trigger":
                                                                                                                                                    {
                                                                                                                                                        foreach (IXMLDOMNode Nodo9 in Nodo8.ChildNodes)
                                                                                                                                                        {
                                                                                                                                                            if (Nodo9.HasChildNodes)
                                                                                                                                                            {
                                                                                                                                                                switch (Nodo9.BaseName)
                                                                                                                                                                {
                                                                                                                                                                    case "triggerRate":
                                                                                                                                                                        {
                                                                                                                                                                            if (VBA.Trim(withBlock1.Range("ER_TriggerRates_V").Value) != "")
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_TriggerRates_V").Value = withBlock1.Range("ER_TriggerRates_V").Value + "; " + Nodo9.nodeTypedValue;
                                                                                                                                                                            }
                                                                                                                                                                            else
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_TriggerRates_V").Value = Nodo9.nodeTypedValue;
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                        }
                                                                                                                                                                    case "triggerPayment":
                                                                                                                                                                        {
                                                                                                                                                                            if (VBA.Trim(withBlock1.Range("ER_TriggerPayments_V").Value) != "")
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_TriggerPayments_V").Value = withBlock1.Range("ER_TriggerPayments_V").Value + "; " + Nodo9.nodeTypedValue;
                                                                                                                                                                            }
                                                                                                                                                                            else
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_TriggerPayments_V").Value = Nodo9.nodeTypedValue;
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                        }
                                                                                                                                                                    case "noTriggerPayment":
                                                                                                                                                                        {
                                                                                                                                                                            if (VBA.Trim(withBlock1.Range("ER_NoTriggerPayments_V").Value) != "")
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_NoTriggerPayments_V").Value = withBlock1.Range("ER_NoTriggerPayments_V").Value + "; " + Nodo9.nodeTypedValue;
                                                                                                                                                                            }
                                                                                                                                                                            else
                                                                                                                                                                            {
                                                                                                                                                                                withBlock1.Range("ER_NoTriggerPayments_V").Value = Nodo9.nodeTypedValue;
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                        }
                                                                                                                                                                }
                                                                                                                                                            }
                                                                                                                                                        }

                                                                                                                                                        break;
                                                                                                                                                    }
                                                                                                                                            }
                                                                                                                                        }
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "interestLeg":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo56 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo56;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "interestCalculation":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo610 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo610;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "floatingRateCalculation":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo75 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo75;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "floatingRateIndex":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_RateIndex_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "indexTenor":
                                                                                                                                {
                                                                                                                                    foreach (IXMLDOMNode currentNodo82 in Nodo7.ChildNodes)
                                                                                                                                    {
                                                                                                                                        Nodo8 = currentNodo82;
                                                                                                                                        if (Nodo8.HasChildNodes)
                                                                                                                                        {
                                                                                                                                            switch (Nodo8.BaseName)
                                                                                                                                            {
                                                                                                                                                case "periodMultiplier":
                                                                                                                                                    {
                                                                                                                                                        withBlock1.Range("I_PeriodMultiplier_V").Value = Nodo8.nodeTypedValue;
                                                                                                                                                        break;
                                                                                                                                                    }
                                                                                                                                                case "period":
                                                                                                                                                    {
                                                                                                                                                        withBlock1.Range("I_Period_V").Value = Nodo8.nodeTypedValue;
                                                                                                                                                        break;
                                                                                                                                                    }
                                                                                                                                            }
                                                                                                                                        }
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                        case "dayCountFraction":
                                                                                                            {
                                                                                                                withBlock1.Range("I_DayCountFraction_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "accrued":
                                                                                                            {
                                                                                                                withBlock1.Range("I_Accrued_V").NumberFormat = "@";
                                                                                                                withBlock1.Range("I_Accrued_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "frequency":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo76 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo76;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "periodMultiplier":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_PeriodMultiplier2_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "period":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_Period2_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                        case "type":
                                                                                                            {
                                                                                                                withBlock1.Range("I_Type_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "exchangeNotional":
                                                                                                            {
                                                                                                                withBlock1.Range("I_ExchangeNotional_V").NumberFormat = "@";
                                                                                                                withBlock1.Range("I_ExchangeNotional_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "interestLegPeriodSchedule":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo611 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo611;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "interestLegPeriod":
                                                                                                            {
                                                                                                                foreach (IXMLDOMNode currentNodo77 in Nodo6.ChildNodes)
                                                                                                                {
                                                                                                                    Nodo7 = currentNodo77;
                                                                                                                    if (Nodo7.HasChildNodes)
                                                                                                                    {
                                                                                                                        switch (Nodo7.BaseName)
                                                                                                                        {
                                                                                                                            case "accrualStartDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_AccrualStartDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("I_AccrualStartDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_AccrualStartDates_V").Value = withBlock1.Range("I_AccrualStartDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_AccrualStartDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "accrualEndDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_AccrualEndDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("I_AccrualEndDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_AccrualEndDates_V").Value = withBlock1.Range("I_AccrualEndDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_AccrualEndDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "fixingDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_FixingDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("I_FixingDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_FixingDates_V").Value = withBlock1.Range("I_FixingDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_FixingDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "settlementDate":
                                                                                                                                {
                                                                                                                                    withBlock1.Range("I_SettlementDates_V").NumberFormat = "@";
                                                                                                                                    if (VBA.Trim(withBlock1.Range("I_SettlementDates_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_SettlementDates_V").Value = withBlock1.Range("I_SettlementDates_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_SettlementDates_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                            case "spreadValue":
                                                                                                                                {
                                                                                                                                    if (VBA.Trim(withBlock1.Range("I_SpreadValues_V").Value) != "")
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_SpreadValues_V").Value = withBlock1.Range("I_SpreadValues_V").Value + "; " + Nodo7.nodeTypedValue;
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        withBlock1.Range("I_SpreadValues_V").Value = Nodo7.nodeTypedValue;
                                                                                                                                    }

                                                                                                                                    break;
                                                                                                                                }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }

                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                                case "costOfHedge":
                                                                    {
                                                                        foreach (IXMLDOMNode currentNodo57 in Nodo4.ChildNodes)
                                                                        {
                                                                            Nodo5 = currentNodo57;
                                                                            if (Nodo5.HasChildNodes)
                                                                            {
                                                                                switch (Nodo5.BaseName)
                                                                                {
                                                                                    case "type":
                                                                                        {
                                                                                            withBlock1.Range("CH_Type_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                    case "delta":
                                                                                        {
                                                                                            foreach (IXMLDOMNode currentNodo612 in Nodo5.ChildNodes)
                                                                                            {
                                                                                                Nodo6 = currentNodo612;
                                                                                                if (Nodo6.HasChildNodes)
                                                                                                {
                                                                                                    switch (Nodo6.BaseName)
                                                                                                    {
                                                                                                        case "type":
                                                                                                            {
                                                                                                                withBlock1.Range("CH_DeltaType_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "floor":
                                                                                                            {
                                                                                                                withBlock1.Range("CH_DeltaFloor_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "liquidityAlpha":
                                                                                                            {
                                                                                                                withBlock1.Range("CH_DeltaLiquidityAlpha_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                        case "asianTailFactor":
                                                                                                            {
                                                                                                                withBlock1.Range("CH_DeltaAsianTailFactor_V").Value = Nodo6.nodeTypedValue;
                                                                                                                break;
                                                                                                            }
                                                                                                    }
                                                                                                }
                                                                                            }

                                                                                            break;
                                                                                        }
                                                                                    case "jumps":
                                                                                        {
                                                                                            withBlock1.Range("CH_Jumps_V").Value = Nodo5.nodeTypedValue;
                                                                                            break;
                                                                                        }
                                                                                }
                                                                            }
                                                                        }

                                                                        break;
                                                                    }
                                                            }
                                                        }
                                                    }

                                                    break;
                                                }
                                        }
                                    }
                                }
                            }
                        }
                        Application.Calculation = xlCalculationAutomatic;
                        ProtectSheet(true, sHoja);

                        bLoadFMM = false;

                        Interaction.MsgBox("FMM file loaded correctly", (MsgBoxStyle)((int)Constants.vbInformation + (int)Constants.vbOKOnly), "FMM FILE");
                    }
                }
            }

            return;

        ERRORES:
            ;

            if (Information.Err().Number == 5)
            {
                return;
            }
        }

        public void ButtonResetFormulas(string sHoja, int iCol, string sProduct)
        {
            if (string.IsNullOrEmpty(Galleta))
                LoginError();

            Application.ScreenUpdating = false;

            bReset = true;

            ResetFormulas(sHoja, iCol, sProduct);

            bReset = false;
        }

        public void ResetButtons(string sHoja)
        {
            {
                var withBlock = ThisWorkbook.Sheets(sHoja);
                withBlock.ButtonLogin.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonLogout.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonLoadDates.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonGenerateFMM.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonCalculatePrice.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonGenFMMCalcPrice.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonGetResult.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonRetrieveXMLs.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonEditFMM.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonLoadStaticData.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonLoadFMM.BackColor = int.MinValue + 0x0000000F;
                withBlock.ButtonResetFormulas.BackColor = int.MinValue + 0x0000000F;
            }
        }

        public void ProtectSheet(bool bProtect, string sHoja)
        {
            ;
            Application.EnableCancelKey = xlErrorHandler;
            switch (bProtect)
            {
                case true:
                    {
                        {
                            var withBlock = ThisWorkbook;
                            withBlock.Sheets(sHoja).Protect("BBVA");
                        }

                        break;
                    }
                case false:
                    {
                        {
                            var withBlock1 = ThisWorkbook;
                            withBlock1.Sheets(sHoja).Unprotect("BBVA");
                        }

                        break;
                    }
            }
            return;
        ERRORES:
            ;

            if (Information.Err().Number == 18)
                ;
        }

        public void FormatCells(string sHoja, int iCol, string sProduct)
        {
            switch (sProduct ?? "")
            {
                case "Autocall":
                    {
                        FormatCellsAutocall(sHoja, iCol);
                        break;
                    }

            }
        }

        public void FormatDatesCells(string sHoja, int iCol, string sProduct)
        {
            switch (sProduct ?? "")
            {
                case "Autocall":
                    {
                        FormatDatesCellsAutocall(sHoja, iCol);
                        break;
                    }

            }
        }

        public void CalculateSwapSpread(string sHoja, int iCol)
        {
            int i;
            int NumFila;
            Variant mFixingDates;
            Variant mEquityLeg;
            Variant mEquitySwapLeg;
            Variant mGenerateAutocallable;
            int UltFila;
            int iDiffMonths;
            var iMinDiff = default(int);
            int iMonth;
            string sPeriod;
            Variant mData;

            {
                var withBlock = ThisWorkbook.Sheets("Aux");
                withBlock.Columns(sColMonths + ":" + sColEMTN2).Clear();
            }

            {
                var withBlock1 = ThisWorkbook.Sheets(sHoja);
                mFixingDates = Application.Run("QBS.DateGen.Param.FixingDates", VBA.CLng(withBlock1.Range("AS_StartDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("AS_EndDate_V").Offset(0, iCol).Value), withBlock1.Range("AS_Frequency_V").Offset(0, iCol).Value);
                mEquityLeg = Application.Run("QBS.DateGen.Param.EquityLeg", mFixingDates, VBA.CLng(withBlock1.Range("D_InitialPaymentDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("D_FinalAlignmentDate_V").Offset(0, iCol).Value), withBlock1.Range("EL_Frequency_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentLag_V").Offset(0, iCol).Value, withBlock1.Range("EL_FixingCalendar_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentCalendar_V").Offset(0, iCol).Value, withBlock1.Range("EL_Alignment_V").Offset(0, iCol).Value, withBlock1.Range("EL_BrokenPeriod_V").Offset(0, iCol).Value, withBlock1.Range("EL_FixingAdjustment_V").Offset(0, iCol).Value, withBlock1.Range("EL_PaymentAdjustment_V").Offset(0, iCol).Value, withBlock1.Range("EL_StickToMothEnd_V").Offset(0, iCol).Value, withBlock1.Range("EL_AdjustInputDates_V").Offset(0, iCol).Value);
                mEquitySwapLeg = Application.Run("QBS.DateGen.Param.EquitySwapLeg", withBlock1.Range("ESL_SwapFrequency_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapPaymentCalendar_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapFixingCalendar_V").Offset(0, iCol).Value, "2B", withBlock1.Range("ESL_SwapAlignment_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapBrokenPeriod_V").Offset(0, iCol).Value, withBlock1.Range("ESL_SwapPaymentAdjustment_V").Offset(0, iCol).Value, VBA.CLng(withBlock1.Range("ESL_SwapStartDate_V").Offset(0, iCol).Value), VBA.CLng(withBlock1.Range("ESL_SwapEndDate_V").Offset(0, iCol).Value), withBlock1.Range("ESL_SwapStickToMothEnd_V").Offset(0, iCol).Value, Interaction.IIf(VBA.IsEmpty(withBlock1.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value) == true, false, withBlock1.Range("ESL_SwapAdjustInputDates_V").Offset(0, iCol).Value));
                mGenerateAutocallable = Application.Run("QBS.DateGen.GenerateAutocallable", mEquityLeg, mEquitySwapLeg, withBlock1.Range("A_EarlyRedemptionFreq_V").Offset(0, iCol).Value, withBlock1.Range("A_FirstEarlyRedemptionPer_V").Offset(0, iCol).Value, withBlock1.Range("A_EarlyRedemptionAlignment_V").Offset(0, iCol).Value, withBlock1.Range("A_AllowEarlyRedemptionMat_V").Offset(0, iCol).Value, Interaction.IIf(VBA.IsEmpty(withBlock1.Range("A_Dates_V").Offset(0, iCol).Value) == true, false, withBlock1.Range("A_Dates_V").Offset(0, iCol).Value), true, true);
                ;
                UltFila = withBlock1.Range(sColSwapStartDates + withBlock1.Rows.Count).End(xlUp).row;
                i = 1;
                if (UltFila >= 3)
                {
                    var loopTo = UltFila;
                    for (NumFila = 3; NumFila <= loopTo; NumFila++)
                    {
                        iDiffMonths = VBA.DateDiff("m", withBlock1.Range(sColSwapStartDates + NumFila).Value, withBlock1.Range(sColSwapEndDates + NumFila).Value);
                        if (NumFila == 3)
                            iMinDiff = iDiffMonths;
                        switch (iDiffMonths)
                        {
                            case var @case when @case < 3:
                                {
                                    mData2(i, 1) = 1;
                                    break;
                                }
                            case var case1 when case1 < 6:
                                {
                                    mData2(i, 1) = 3;
                                    break;
                                }

                            default:
                                {
                                    mData2(i, 1) = 6;
                                    break;
                                }
                        }
                        if (i == 1)
                        {
                            mData2(i, 2) = mData2(i, 1);
                        }
                        else if (i > 1)
                        {
                            mData2(i, 2) = mData2(i - 1, 2) + mData2(i, 1);
                        }
                        if (mData2(i, 2) < 12)
                        {
                            sPeriod = mData2(i, 2) + "m";
                        }
                        else if (mData2(i, 2) == 12) // igual al año
                        {
                            sPeriod = VBA.Int(mData2(i, 2) / 12) + "y";
                        }
                        else // MÁS QUE UN AÑO
                        {
                            sPeriod = VBA.Replace(VBA.Int(mData2(i, 2) / 12) + "y" + Interaction.IIf(VBA.Int(mData2(i, 2) % 12) == 0, "", VBA.Int(mData2(i, 2) % 12) + "m"), "y0m", "");
                        }
                        ThisWorkbook.Sheets("Aux").Range(sColEMTN2 + NumFila).Formula = "=IFERROR(VLOOKUP(" + "\"" + sPeriod + "\"" + ", " + sColMonthsEMTNRAR + "1:" + sColEMTN + "121, 2, FALSE), 0)";
                        if (iDiffMonths < iMinDiff)
                            iMinDiff = iDiffMonths;
                        i = i + 1;
                    }
                }

                switch (iMinDiff)
                {
                    case var case2 when case2 < 3:
                        {
                            iMonth = 1;
                            break;
                        }
                    case var case3 when case3 < 6:
                        {
                            iMonth = 3;
                            break;
                        }

                    default:
                        {
                            iMonth = 6;
                            break;
                        }
                }

                {
                    var withBlock2 = ThisWorkbook.Sheets("Aux");
                    withBlock2.Range(sColMonthsEMTNRAR + "1:" + sColEMTN2 + "121").NumberFormat = "General";
                    if (ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value == "EUR")
                    {
                        withBlock2.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").FormulaArray = "=bbva_GetTyMatrix(\"REFERENCE\",\"IR_NOTE_SPREAD\",TODAY()," + VBA.Chr(34) + VBA.Trim(VBA.Replace(ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value, "EUR", "") + " MTN " + iMonth) + "m Callable Spread\")";
                    }
                    else // cualquier otra divisa
                    {
                        withBlock2.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").FormulaArray = "=bbva_GetTyMatrix(\"REFERENCE\",\"IR_NOTE_SPREAD\",TODAY()," + VBA.Chr(34) + ThisWorkbook.Sheets(sHoja).Range("Ac_Currency_V").Offset(0, iCol).Value + " MTN 1m Callable Spread\")";
                    }
                    if (withBlock2.Range(sColMonthsEMTNRAR + "1").Value == "You are not logged")
                    {
                        Interaction.MsgBox("You are not logged in Typhoon Add-Ins", (MsgBoxStyle)((int)Constants.vbCritical + (int)Constants.vbOKOnly), "ERROR LOGIN");
                        return;
                    }
                    else if (VBA.InStr(1, withBlock2.Range(sColMonthsEMTNRAR + "1").Value, "ERROR", Constants.vbTextCompare) > 0)
                    {
                        Interaction.MsgBox.Range(sColMonthsEMTNRAR + "1").Value(default, (int)Constants.vbCritical + (int)Constants.vbOKOnly, "ERROR");
                        return;
                    }
                    withBlock2.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").Copy();
                    withBlock2.Range(sColMonthsEMTNRAR + "1:" + sColRAR + "121").PasteSpecial(xlPasteValues);
                    Application.CutCopyMode = false;

                    UltFila = withBlock2.Range(sColEMTN2 + withBlock2.Rows.Count).End(xlUp).row;
                    mData = withBlock2.Range(sColEMTN2 + "3:" + sColEMTN2 + UltFila).Value;
                    if (UltFila >= 3)
                    {
                        var loopTo1 = UltFila - 2;
                        for (i = 1; i <= loopTo1; i++)
                            mData(i, 1) = mData(i, 1) / 10000;
                    }
                }
                if (VBA.Left(sHoja, 4) != "Bulk")
                    withBlock1.Range(sColSwapSpread + "3:" + sColSwapSpread + UltFila).Value = mData;
            }
        }
    }
}
