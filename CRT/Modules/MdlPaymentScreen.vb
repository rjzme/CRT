Module MdlPaymentScreen
    Public Property blnCancelMacro As Boolean
    Public Property ClmLines As Boolean

    Function BuildADJ(ByVal excelrow As Integer)


        DenialCode = "I4,I0,I1,I2,I3,I5,ID,IE,IG,IJ,IH,IO,FE,K1,K2,K3,K4,KG,KH,KJ,KK,KM,KQ,KU,KV,KW,KT,K0,MW,O9,O5,O3,OG,OH,OJ,Q4,Q7,Q8,Q9,QD,VD,VQ,WA,WE,Y1,Y2,Y6,LN,LY,AA,DJ,GZ,IA,KU,MP,PI,VQ,1Y,3D,2Z,09,VP,PC,R6,XS,BD,DK,G1,I0,LC,NQ,PJ,WW,2B,3T,2T,7,UW,UZ,R7,OX,BJ,DM,G6,I5,LG,OC,P6,WZ,2E,39,2R,Y6,UB,O5,SR,M9,B5,EF,G9,JA,LH,OI,Q0,XG,2H,54,2M,YP,Q9,PU,US,7V,B6,E4,HJ,JS,LJ,OM,Q3,XK,2K,58,2I,XH,3L,OJ,LQ,7U,CH,E6,H0,J0,L7,OY,YR,2P,4R,5B,W4,VB,RZ,88,7L,CQ,FA,H7,KA,MK,O6,UM,06,2S,3X,2D,WY,4N,6C,83,7C,C3,GR,H8,KL,MN,O8,VM,08,2X,3E,2A,VR,VF,M0,K5,EL"
        SSPCodes = "IT,IL,IX,IS,FV,CY"

        autECLPSObj.SendKeys("[pf12]")
        ApiChk()
        clearChk()
        ApiChk()
        fundingChk()


        ApiChk()

        verifyEDS = False

        Dim strNYSurcharge As String

        If ChkNYS = "A" Then

            strCntrLine1 = CltrLine1() & "," & Pre & Tin
            ApiChk()
            strCntrLine1 = Replace(strCntrLine1, " ", "")   '*******Removes any spaces in control line for more room*******

            If Len(strCntrLine1 & ",stao45ct,,") > 78 Then
                Dim intCtrLgth, intClmLgth
                intCtrLgth = Len(strCntrLine1 & ",stao45ct,,") - 78   '*******Determines amount of characters over
                intClmLgth = Len(ClmLines)

                If intCtrLgth = 8 Or intCtrLgth > intClmLgth Then
                    strNYSurcharge = MsgBox("Control line is over available space by '" & intCtrLgth &
                                                              "' characters. Do you want to remove the NY tax id?" &
                                                              vbCrLf & "Length of claim lines entered = " & intClmLgth, vbYesNo + vbDefaultButton1)
                    If strNYSurcharge = vbNo Then
                        respo = MsgBox("Control line is too long to include the NY tax id and the Claim Lines entered.  Do you " &
                                                  "want to remove the Claim Lines? Clicking 'No' will exit the macro.", vbYesNo + vbDefaultButton1)
                        If respo = vbNo Then
                            strCntrLine1 = strCntrLine1 & ",stao45ct"

                        Else
                            strCntrLine1 = strCntrLine1 & ",stao45ct"
                        End If
                    End If
                Else
                    strCntrLine1 = strCntrLine1 & ",stao45ct,,"
                End If
            Else
                strCntrLine1 = strCntrLine1 & ",2999999999"

            End If

            autECLPSObj.SendKeys(strCntrLine1 & "[erase eof]", 2, 2)
            autECLPSObj.SendKeys("ADJ", 2, 2)
            Enter()
            ApiChk()

        Else
            adjline = CltrLine1() & "," & Pre & Tin
            autECLPSObj.SendKeys("[erase eof]", 2, 2)
            autECLPSObj.SendKeys(adjline, 2, 2)
            autECLPSObj.SendKeys("ADJ", 2, 2)
            Enter()

        End If

        ApiChk()

        If Pre = 8 Then
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
        ElseIf Pre = 7 Then
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
        ElseIf Pre = 9 Then
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
            autECLPSObj.SendKeys("99999", 3, 7)
            Enter()
        End If

        If MainForm.MetroTextBoxSuffix.Text = "" Then

            ApiChk()
        Else
            ApiChk()
            Suffix = MainForm.MetroTextBoxSuffix.Text

        End If


        If NysSta = "A" Then
            ApiChk()
            autECLPSObj.SendKeys(Suffix, 3, 7)
            autECLPSObj.SendKeys("E", 3, 13)
            autECLPSObj.SendKeys("00035", 3, 22)
            autECLPSObj.SendKeys("E", 3, 28)
            Enter()

        Else
            autECLPSObj.SendKeys(Suffix, 3, 7)
            autECLPSObj.SendKeys("E", 3, 13)
            Enter()

        End If
        DkiChk()


        ApiChk()
        Dim arrayIndex As Integer = 0
        Dim LineSVCCode As String
        Dim LineRemarkCode As String
        For i = 10 To 22 Step 2
            ApiChk()
            LineSVCCode = Trim(autECLPSObj.GetText(i, 4, 6))
            LineRemarkCode = autECLPSObj.GetText(i, 36, 2)

            CodesCol(arrayIndex) = LineSVCCode
            RemarkColl(arrayIndex) = LineRemarkCode
            arrayIndex = arrayIndex + 1
        Next

        If MainForm.MetroRadioButtonINNYes.Checked Then
            ApiChk()
            autECLPSObj.SendKeys("W", 5, 3)
            Enter()

        End If
        If Pz1(excelrow) Then
            Return True
        End If
        ApiChk()
        autECLPSObj.SendKeys("EDS", 2, 2)
        Enter()

        Dim ReadyChk
        ReadyChk = autECLPSObj.GetText(24, 2, 13)
        ApiChk()
        If ReadyChk = "INVALID ENTRY" Then
            ApiChk()
            MsgBox("Please Input Correct Value and Continue")

        End If
        ApiChk()
        autECLPSObj.SendKeys("-", 22, 61)
        ApiChk()
        autECLPSObj.SendKeys("          ", 22, 69)
        Enter()

        Onr()
        Rsc = "01"
        If MainForm.MetroRadioButtonIntYes.Checked Then
            INT = "Yes"
        End If

        If MainForm.MetroRadioButtonIntNo.Checked Then
            INT = "No"
        End If

        If INT = "Yes" Then

            If Funding = "1" Or Funding = "2" Then


                ApiChk()
                autECLPSObj.SendKeys("Y", 23, 71)
                Enter()


            ElseIf SState = "MO" Or SState = "AR" Or SState = "NC" Or SState = "LA" Or SState = "RI" Or SState = "GA" Then


                If Erisa = "N" Then


                    ApiChk()
                        autECLPSObj.SendKeys("Y", 23, 71)
                        Enter()

                Else
                    ApiChk()
                    autECLPSObj.SendKeys("N", 23, 71)
                    Enter()

                    ApiChk()
                    RscChk()
                    Enter()
                End If
            Else
                ApiChk()
                autECLPSObj.SendKeys("N", 23, 71)
                Enter()
                ApiChk()
                RscChk()
                Enter()

            End If

        ElseIf INT = "No" Then

            ApiChk()
            autECLPSObj.SendKeys("N", 23, 71)
            Enter()
            ApiChk()
            RscChk()
            Enter()
        End If

        ApiChk()

        Dim cc As String

        If autECLPSObj.GetText(1, 67, 7) = "NXT SCR" Then
            verifyEDS = True
        Else
            verifyEDS = False
            autECLPSObj.SendKeys("MPC", 2, 2)
            Status = "EDS screens not available, unable to do ADJ."
            getDateTime(excelrow)
            Return True
        End If
        If MainForm.CheckBoxCause.Checked Then
            chkCauseCode = Trim(MainForm.MetroTextBoxCauseCode.Text)

        End If

        For EDSCount = 6 To 18 Step 3
            ApiChk()
            BLKChk = Trim(autECLPSObj.GetText(EDSCount, 11, 6))
            EdsTest = autECLPSObj.GetText(EDSCount, 2, 3)
            cc = autECLPSObj.GetText(EDSCount, 11, 1)
            RemarkChk = autECLPSObj.GetText(EDSCount, 45, 2)
            Dim isoBlk As String = autECLPSObj.GetText(EDSCount, 38, 2)
            If EdsTest = "---" Then
                ApiChk()
                Dim ProtectField As String
                Dim ProtectField1 As String


                If claimType = "UB" Then
                    ApiChk()
                    EdsOV = "--"
                    If EdsHalf(excelrow) Then
                        ApiChk()
                        Return True

                    End If
                    Exit For
                    ApiChk()

                ElseIf claimType = "HCFA" Then

                    If isoBlk = "30" Then
                        ApiChk()
                        autECLPSObj.SendKeys("Blank-", EDSCount, 4)
                        Enter()

                    Else

                        ApiChk()
                        ProtectField = FieldState(EDSCount, 59)
                        ProtectField1 = FieldState(EDSCount + 1, 6)

                        If InStr(DenialCode, RemarkChk) > 0 Then
                            ApiChk()
                            autECLPSObj.SendKeys("--21", EDSCount, 38)

                        ElseIf InStr(SSPCodes, RemarkChk) > 0 And MainForm.MetroRadioButtonSSPYes.Checked Then

                            ApiChk()
                            autECLPSObj.SendKeys("--21", EDSCount, 38)
                        Else

                            If ProtectField = False Then

                                ApiChk()
                                autECLPSObj.SendKeys("------", EDSCount, 59)
                                autECLPSObj.SendKeys("--", EDSCount, 66)

                            End If
                            If ProtectField1 = False Then

                                ApiChk()
                                autECLPSObj.SendKeys("------", EDSCount + 1, 6)
                                autECLPSObj.SendKeys("--", EDSCount + 1, 13)
                                autECLPSObj.SendKeys("---", EDSCount + 1, 20)

                            End If

                        End If

                        If MainForm.CheckBoxCause.Checked Then
                            ApiChk()
                            autECLPSObj.SendKeys("OPN", EDSCount, 2)
                            ApiChk()
                            Enter()
                            ApiChk()
                            autECLPSObj.SendKeys("Y", 10, 22)
                            ApiChk()
                            Enter()
                            ApiChk()
                            autECLPSObj.SendKeys(chkCauseCode, 10, 6)
                            Enter()
                            ApiChk()
                            autECLPSObj.SendKeys("[PF3]")
                            ApiChk()
                        End If

                    End If


                End If





                If EDSCount = 18 Then

                    If EdsTest <> "---" Then
                        Exit For
                    Else
                        Enter()
                        autECLPSObj.SendKeys("[PF8]")
                        ApiChk()

                        If autECLPSObj.GetText(24, 2, 7) = "ALREADY" Or autECLPSObj.GetText(24, 2, 7) = "NO MORE" Then
                            Exit For
                        Else
                            EDSCount = 3
                        End If
                    End If
                End If

            ElseIf BLKChk = "OI" Or BLKChk = "TDPD" Or BLKChk = "ACPD" Or BLKChk = "PRE" Or BLKChk = "OIMEDI" Then
                ApiChk()
                autECLPSObj.SendKeys("blank-", EDSCount, 11)

            ElseIf BLKChk = "OIM" Then
                ApiChk()
                'Med
                ApiChk()
                autECLPSObj.SendKeys("Blank-", EDSCount, 11)
                Enter()

            End If
        Next
        ApiChk()
        'VdD
        'split
        'CChange

        ApiChk()
        autECLPSObj.SendKeys("A", 1, 75)
        Enter()
        ApiChk()
        Onr()
        ApiChk()
        EDsChk = autECLPSObj.GetText(1, 67, 7)
        If EDsChk = "NXT SCR" Then

            ApiChk()
            autECLPSObj.SendKeys("A", 1, 75)
            Enter()
        End If
        If Pz1(excelrow) Then
            Return True
        End If

        If MainForm.CheckBoxCause.Checked Then
            ApiChk()
            autECLPSObj.SendKeys(chkCauseCode, 7, 4)
            Enter()
            ApiChk()

            If Pz1(excelrow) Then
                Return True
            End If
        End If

        ApiChk()
        Dim PendChk As String
        Dim PendRChk As String



        For x = 10 To 22 Step 2
            ApiChk()
            PendChk = Trim(autECLPSObj.GetText(x, 29, 2))
            PendRChk = autECLPSObj.GetText(x, 36, 2)
            If PendChk = "P" And PendRChk = "CV" Or PendRChk = "AD" Or PendRChk = "PM" Or PendRChk = "AM" Or PendRChk = "CM" Then
                ApiChk()
                autECLPSObj.SendKeys("09", x, 29)
                Enter()
            End If

            If x = 22 Then
                If autECLPSObj.GetText(x, 4, 6) <> "------" Then
                    Status = "To many line Manual Intervention Required"
                End If
            End If

        Next

        For i = 10 To 22 Step 2

            Dim noSvcChk As String
            Dim mpcRCchk As String
            Dim ProtectField As String

            noSvcChk = Trim(autECLPSObj.GetText(i, 4, 6))
            mpcRCchk = autECLPSObj.GetText(i, 36, 2)

            If noSvcChk = "------" Then
                Exit For
            End If

            If noSvcChk = "OI" Or noSvcChk = "OIM" Or noSvcChk = "OIMEDI" Or InStr(DenialCode, mpcRCchk) > 0 Then
                ApiChk()
                autECLPSObj.SendKeys("21", i, 32)

            Else
                If MainForm.MetroRadioButtonSSPYes.Checked And InStr(SSPCodes, mpcRCchk) > 0 Then
                    ApiChk()
                    autECLPSObj.SendKeys("21", i, 32)
                Else
                    ProtectField = FieldState(i, 50)
                    ApiChk()
                    autECLPSObj.SendKeys("21", i, 32)
                    If ProtectField = False Then
                        autECLPSObj.SendKeys("-----", i, 50)
                        autECLPSObj.SendKeys("--", i, 56)
                    End If
                End If
            End If

        Next

        If Pz1(excelrow) Then
            Return True
        End If

        arcChk()

        Dim ce As String
        Dim bc As String
        Dim bcc As String
        ApiChk()
        ce = autECLPSObj.GetText(24, 3, 5)
        If ce = "E985T" Then
            ApiChk()
            bc = autECLPSObj.GetText(24, 19, 6)
            ApiChk()
            bcc = autECLPSObj.GetText(24, 26, 2)
            ApiChk()
            autECLPSObj.SendKeys(bc, 6, 4)
            ApiChk()
            autECLPSObj.SendKeys(bcc, 6, 11)
            Enter()
        End If

        ApiChk()
        arcChk()
        ApiChk()

        For x = 10 To 22 Step 2
            LineDenialCode = autECLPSObj.GetText(x, 36, 2)
            If LineDenialCode = "VP" Or LineDenialCode = "BT" Then
                autECLPSObj.SendKeys("17", x, 29)
                EdsOV = ""
                ApiChk()
                autECLPSObj.SendKeys("EDS", 2, 2)
                Enter()
                ApiChk()

                If EdsHalf(excelrow) Then
                    Return True
                End If

            End If

            If LineDenialCode = "05" Or LineDenialCode = "TQ" Then
                Status = "Please Check For Duplicate"
                getDateTime(excelrow)
                Return True
            End If
            If LineDenialCode = "6X" Then
                ApiChk()
                MsgBox("Please check for penalty")
            End If
        Next


        If CHECK_FOR_EDIT_MPC(excelrow) Then
            Return True
        End If
        ApiChk()

        For y = 10 To 22 Step 2

            copay = autECLPSObj.GetText(y, 4, 6)
            If copay = "NCOPAY" Then
                copayNBR = Trim(autECLPSObj.GetText(y, 25, 3))
                If copayNBR = "" Or copayNBR = "0" Then
                    copayN = InputBox("Please enter the Copay No..", "Copay", "1")
                    ApiChk()
                    autECLPSObj.SendKeys(copayN, y, 27)
                    Enter()


                End If
            End If
            ApiChk()
            CopayChk = autECLPSObj.GetText(24, 3, 6)

            If CopayChk = "E1542C" Then

                ApiChk()
                autECLPSObj.SendKeys("02", y, 29)
                Enter()

            End If

        Next



        If Pz1(excelrow) Then
            Return True
        End If

        If mpcIntChk(excelrow) Then
            Return True
        End If


        If Pz1(excelrow) Then
            Return True
        End If


        If Rsc = "17" Then
            ApiChk()
            autECLPSObj.SendKeys("G81", 2, 2)
            Enter()
            If Pz1(excelrow) Then
                Return True
            End If
        End If

        If NysSta = "A" Then

            ApiChk()

            If claimType = "UB" Then
                Nysurcharge(excelrow)
                If Nysurcharge(excelrow) Then
                    Return True
                End If
            End If
        End If
        ApiChk()

        If InStr(plcSv, "20") > 0 Then
            ApiChk()
            OpnIbaag()
            MsgBox("Urgent Care Claim Check for Benefit")
        End If



        If CHECK_PaidAmount_MPC(excelrow) Then
            Return True
        End If
        ApiChk()

        payment = Decimal.Parse(PaidDoll & "." & PaidCents)

        Dim intDoll As String
        Dim intCents As String
        Dim intpaid As Decimal
        For i = 10 To 22 Step 2
            If Trim(autECLPSObj.GetText(i, 4, 6)) = "CXINT" Then
                ApiChk()
                intDoll = autECLPSObj.GetText(i + 1, 61, 5)
                intCents = autECLPSObj.GetText(i + 1, 67, 2)
                intpaid = Decimal.Parse(intDoll & "." & intCents)
                CurrentPayment = CurrentPayment - intpaid

                Exit For
            End If
        Next


        If CurrentPayment < payment Then
            ApiChk()

            MdlOverPaid.main(excelrow)

        ElseIf payment < 0.01 Then
            If isoFound = "30" Then
                ApiChk()
                BuildISO(excelrow)
            Else
                ApiChk()
                ReconeChk(excelrow)
            End If
        Else

            ApiChk()
            BuildISO(excelrow)

        End If
        ApiChk()

        'If mpcIntChk(excelrow) Then
        '    Return True

        'End If

        ApiChk()


        Enter()
        MsgBox("Check Screen")
        ApiChk()
        autECLPSObj.SendKeys("MPP", 2, 2)
        ApiChk()
        Enter()
        ApiChk()



        If MainForm.MetroRadioButtonOrsYes.Checked Then
            OrsCreation()
        End If

        ApiChk()
        GetClaimInfo(excelrow)

        Return False
    End Function



    Function NySurchargeCode(ByVal theStr As String)
        Dim SvcCode As String
        ApiChk()
        SvcCode = theStr
        Do Until InStr(SvcCode, "-") = 0
            SvcCode = Replace(SvcCode, "-", "")
        Loop
        ApiChk()
        SvcCode = Replace(SvcCode, "REV", "")
        If Left(SvcCode, 1) <> "C" Then
            If IsNumeric(SvcCode) = True Then
                SvcCode = Left(SvcCode, 1)
            End If
        End If

        Select Case SvcCode
            Case "SP"
                NySurchargeCode = "C02NYS"
            Case "IC"
                NySurchargeCode = "C05NYS"
            Case "NB"
                NySurchargeCode = "C08NYS"
            Case "IS"
                NySurchargeCode = "C04NYS"
            Case "OA"
                NySurchargeCode = "C09NYS"
            Case "OPS"
                NySurchargeCode = "C11NYS"
            Case "MISC"
                NySurchargeCode = "C09NYS"
            Case "EMERG"
                NySurchargeCode = "C12NYS"


            Case Else
                NySurchargeCode = SvcCode & "NYS"
        End Select

        If blnCancelMacro Then
            Exit Function
        End If

    End Function
    Sub Onr()
        Dim OonR As String

        ApiChk()
        OonR = autECLPSObj.GetTExt(24, 2, 18)
        ApiChk()
        If OonR = "INVALID OONR VALUE" Then
            If claimType = "UB" Then
                ApiChk()
                autECLPSObj.SendKeys("R", 3, 79)
            Else
                ApiChk()
                autECLPSObj.SendKeys("R", 3, 48)

            End If

            Enter()
        End If


    End Sub

    Function EdsHalf(ByVal excelrow As Integer)
        Dim NewRev, NewChg, NewCents, NewDos, NewRemark As String
        Dim OldRev, OldChg, OldCents, OldDos, OldRemark, LineChk, LineChk1, Modf, BLKChk, isoBlk As String
        Dim ProtectField As String


        ApiChk()
        autECLPSObj.SendKeys("[PF7]")
        ApiChk()

        For EDSCount = 6 To 18 Step 3
            ApiChk()
            LineChk = Trim(autECLPSObj.GetText(EDSCount, 2, 3))
            BLKChk = Trim(autECLPSObj.GetText(EDSCount, 11, 6))
            isoBlk = autECLPSObj.GetText(EDSCount, 38, 2)
            ApiChk()

            If isoBlk = "30" Then
                ApiChk()
                autECLPSObj.SendKeys("Blank-", EDSCount, 11)
                Enter()
            Else
                If LineChk = "---" Then
                    ApiChk()
                    autECLPSObj.SendKeys(EdsOV, EDSCount, 38)
                    autECLPSObj.SendKeys("21", EDSCount, 41)
                    autECLPSObj.SendKeys("---", EDSCount, 45)
                    ProtectField = FieldState(EDSCount + 1, 50)
                    If ProtectField = False Then
                        ApiChk()
                        autECLPSObj.SendKeys("---", EDSCount + 1, 50)
                    End If

                    Enter()
                    ApiChk()
                    autECLPSObj.SendKeys("OPN", EDSCount, 2)
                    Enter()
                    ApiChk()

                    Onr()

                    ApiChk()

                    For NCount = 8 To 17 Step 3
                        ApiChk()
                        autECLPSObj.SendKeys("Y", 2, 49)
                        LineChk1 = Trim(autECLPSObj.GetText(NCount, 2, 3))
                        Modf = Trim(autECLPSObj.GetText(NCount, 20, 2))
                        If LineChk1 = "---" Then
                            If NCount = 8 Then
                                ApiChk()
                                OldRev = Trim(autECLPSObj.GetText(8, 7, 6))
                                OldChg = autECLPSObj.GetText(8, 47, 6)
                                OldCents = autECLPSObj.GetText(8, 54, 2)
                                OldDos = autECLPSObj.GetText(8, 33, 4)
                                OldRemark = Trim(autECLPSObj.GetText(8, 67, 2))
                            End If

                            If Modf = "50" Then
                                Status = "Mod 50 Please Split The Line"
                                getDateTime(excelrow)
                                Return True
                            End If

                            ApiChk()
                            autECLPSObj.SendKeys("--", NCount, 44)
                            ApiChk()
                            autECLPSObj.SendKeys("--------", NCount, 57)
                            ApiChk()
                            autECLPSObj.SendKeys("---", NCount, 67)
                            ApiChk()
                            Enter()

                            If MainForm.CheckBoxCause.Checked Then

                                ApiChk()
                                autECLPSObj.SendKeys("Y", NCount + 2, 56)
                                ApiChk()
                                Enter()
                                ApiChk()
                                autECLPSObj.SendKeys(chkCauseCode, NCount + 2, 43)
                                Enter()
                                ApiChk()
                            End If

                        End If


                        If NCount = 17 And LineChk1 = "---" Then
                            Enter()

                            autECLPSObj.SendKeys("[PF8]")
                            ApiChk()
                            NewRev = Trim(autECLPSObj.GetText(8, 7, 6))
                            NewChg = autECLPSObj.GetText(8, 47, 6)
                            NewCents = autECLPSObj.GetText(8, 54, 2)
                            NewDos = autECLPSObj.GetText(8, 33, 4)
                            NewRemark = Trim(autECLPSObj.GetText(8, 67, 2))

                            If NewRev = OldRev And NewChg = OldChg And NewCents = OldCents And NewDos = OldDos And NewRemark = OldRemark Then

                                ApiChk()
                                Exit For
                            End If
                            NCount = 5

                        End If
                    Next
                    ApiChk()
                        autECLPSObj.SendKeys("[PF3]")

                        If EDSCount = 18 And LineChk = "---" Then

                            autECLPSObj.SendKeys("[PF8]")
                            ApiChk()

                            If autECLPSObj.GetText(24, 2, 7) = "ALREADY" Or autECLPSObj.GetText(24, 2, 7) = "NO MORE" Then
                                ApiChk()
                                Exit For
                            Else
                                EDSCount = 3
                            End If

                        End If


                    Else
                        ApiChk()
                    Exit For

                End If

            End If



        Next


    End Function

    Function Nysurcharge(ByVal excelrow As String)
        Dim chkOPN, NYCharge(), NYSCount, svctest, NYSvcCode, PrevSvcCode, Count2, EdsChk As String
        Dim Count As Integer

        DupChk = InStr(40, autECLPSObj.GetText(2, 2, 80), ",,")
        If DupChk > 0 Then
            autECLPSObj.SendKeys("[erase eof]", 2, DupChk)
            ApiChk()

            For edscount = 10 To 22 Step 2
                EdsOV = autECLPSObj.GetText(edscount, 29, 1)
                If EdsOV <> "P" And EdsOV <> "3" Then
                    autECLPSObj.SendKeys("01", edscount, 29)
                End If
            Next

            If blnCancelMacro Then
                Return True
            End If

        End If


        autECLPSObj.SendKeys("EDS", 2, 2)
        Enter()
        ApiChk()

        If blnCancelMacro Then
            Return True
        End If
        ApiChk()

        If autECLPSObj.GetText(1, 67, 7) <> "NXT SCR" Then
            autECLPSObj.SendKeys("MPC", 2, 2)
            Count = 0
            Count2 = 12
            Do
                svcFound = autECLPSObj.GetText(Count2, 4, 6)
                SvC = InStr(svcFound, "------")
                If SvC > 0 Then
                    Return True
                End If
                Count2 = Count2 + 2
                If Count2 = 24 Then
                    Status = "Please manually add the NY Surcharge line."
                    getDateTime(excelrow)
                    Return True
                End If
            Loop
            If Count2 <> 24 Then
                PrevSvcCode = Trim(autECLPSObj.GetText(Count2 - 2, 4, 6))
                NYSvcCode = NySurchargeCode(PrevSvcCode)
                'Request 1241
                If blnCancelMacro = True Then
                    Return True
                End If
                autECLPSObj.SendKeys(".-", Count2, 1)
                autECLPSObj.SendKeys(NYSvcCode, Count2, 4)
                autECLPSObj.SendKeys(".-----", Count2, 11)
                autECLPSObj.SendKeys(".-----", Count2, 18)
                autECLPSObj.SendKeys("000", Count2, 25)
                autECLPSObj.SendKeys("06", Count2, 29)
                autECLPSObj.SendKeys("2", Count2, 32)
                autECLPSObj.SendKeys("2", Count2, 34)
                autECLPSObj.SendKeys("YW", Count2, 36)
            End If
        Else
            Count = 0
            Count2 = 6
            For x = 6 To 18 Step 3
                chkOPN = autECLPSObj.GetText(x, 2, 3)
                If chkOPN = "---" Then
                    ReDim Preserve NYCharge(Count)
                    NYCharge(Count) = x
                    Count = Count + 1
                End If
            Next
            Do
                svcFound = autECLPSObj.GetText(Count2, 2, 3)
                SvC = InStr(svcFound, "---")
                If SvC > 0 Then
                    Exit Do
                End If
                Count2 = Count2 + 3
                If Count2 = 21 Then
                    Status = "Please manually add the NY Surcharge line."
                    getDateTime(excelrow)
                    clearChk()
                    ApiChk()
                    Return True
                End If
            Loop
            If Count2 <> 21 Then
                If UBound(NYCharge) <= 1 Then
                    NYSCount = NYCharge(0)
                    For Count = 0 To UBound(NYCharge)
                        autECLPSObj.SendKeys("IA", NYSCount, 2)
                        Enter()
                        ApiChk()

                        For x = 6 To 18 Step 3
                            svctest = autECLPSObj.GetText(x, 11, 6)
                            If svctest = "------" Then
                                PrevSvcCode = Trim(autECLPSObj.GetText(x - 3, 11, 6))
                                NYSvcCode = NySurchargeCode(PrevSvcCode)
                                'Request 1241
                                If blnCancelMacro = True Then
                                    Return True
                                End If
                                autECLPSObj.SendKeys(".-", x, 8)
                                autECLPSObj.SendKeys(NYSvcCode, x, 11)
                                autECLPSObj.SendKeys(".-----", x, 18)
                                autECLPSObj.SendKeys(".-----", x, 25)
                                autECLPSObj.SendKeys("00000", x, 32)
                                autECLPSObj.SendKeys("2", x, 41)
                                autECLPSObj.SendKeys("2", x, 43)
                                autECLPSObj.SendKeys("YW", x, 45)
                                Enter()
                                ApiChk()
                                Exit For
                            End If
                        Next
                        If Count = 0 Then
                            NYCharge(0) = NYCharge(0) + 3
                        ElseIf Count = 1 Then
                            NYCharge(Count) = NYCharge(Count) + 6
                        End If
                        NYSCount = NYSCount + 6
                    Next
                Else
                    clearChk()
                    ApiChk()
                    Status = "Please manually add the NY Surcharge line."
                    getDateTime(excelrow)

                    Return True
                End If
                For x = 6 To 18 Step 3
                    chkOPN = autECLPSObj.GetText(x, 2, 3)
                    If chkOPN = "---" Then
                        autECLPSObj.SendKeys("--", x, 38)
                    End If
                Next
                For Count = 0 To UBound(NYCharge)
                    autECLPSObj.SendKeys("--", NYCharge(Count), 38)
                Next
                Enter()
                ApiChk()
                autECLPSObj.SendKeys("A", 1, 75)
                Enter()
                ApiChk()
                Onr()
                ApiChk()
                EdsChk = autECLPSObj.GetText(1, 67, 7)
                If EdsChk = "NXT SCR" Then

                    ApiChk()
                    autECLPSObj.SendKeys("A", 1, 75)
                    Enter()
                End If



                If blnCancelMacro Then
                    Return True
                End If
            End If
        End If
        ApiChk()
        autECLPSObj.SendKeys("S", 7, 30)
        Enter()
        Return False

    End Function

    Sub DkiChk()
        Dim dkihold As String

        Do
            dkihold = autECLPSObj.GetText(2, 2, 3)
            If dkihold = "DKC" Then
                ApiChk()
                autECLPSObj.SendKeys("NA", 7, 14)
                Enter()
            End If
        Loop Until dkihold = "MPC"

    End Sub

    Function Pz1(ByVal excelrow As String)


        If InStr(autECLPSObj.GetText(2, 2, 80), "PZ") > 0 Then
            Status = "Duplicate claim"
            getDateTime(excelrow)

            Return True
        End If

        Return False
    End Function

    Sub ControlLine()
        Dim cnt, cnt2, colFound, endDel As String
        cnt = 24
        cnt2 = 0
        colFound = autECLPSObj.GetText(2, cnt, 1)
        Do
            colFound = autECLPSObj.GetText(2, cnt, 1)
            Select Case colFound
                Case " "
                Case ","
                    endDel = cnt + 2
                    Exit Do
            End Select
            cnt = cnt + 1
        Loop
        cntrLine1 = autECLPSObj.GetText(2, 2, endDel)

    End Sub

    Sub arcChk()
        Dim archold As String

        Do
            archold = autECLPSObj.GetText(2, 2, 3)
            If archold = "ARC" Then
                ApiChk()
                autECLPSObj.SendKeys("S", 3, 5)
                Enter()
            End If
        Loop Until archold = "MPC"

    End Sub

    Function BuildISO(ByVal excelrow As String)


        If Pz1(excelrow) Then
            Return True
        End If

        If isoFound = "30" Then
            Dim cPaid As Double = Double.Parse(Trim(PaidDoll & "." & PaidCents))
            Dim tISO As Double = cPaid + isoPaid

            MsgBox(tISO)

            'If InStr(tISO, ".") = 0 Then
            '    tISO = tISO & ".00"
            'ElseIf Len(tISO) - InStr(tISO, ".") = 1 Then
            '    tISO = tISO & "0"
            'End If
            'PaidDoll = Left(tISO, InStr(tISO, ".") - 1)
            'PaidDoll = PaidDoll.PadLeft(5, "0")
            'PaidCents = Right(tISO, 2)
            'If Left(PaidCents, 1) = "." Then
            '    PaidCents = Right(PaidCents, 1)
            'End If

            Dim testArr() As String = (tISO.ToString()).Split(".")

            PaidDoll = testArr(0)
            PaidDoll = PaidDoll.PadLeft(5, "0")
            PaidCents = testArr(1)

            'MsgBox(PaidDoll)
            'MsgBox(PaidCents)

        Else

            ApiChk()

        End If

        For i = 10 To 22 Step 2
            ApiChk()
            Dim svcCode1 As String = Trim(autECLPSObj.GetText(i, 4, 6))
            Dim MpcisoFound As String = autECLPSObj.GetText(i, 29, 2)
            If svcCode1 = "------" Or svcCode1 = "CXINT" Or MpcisoFound = "30" Then
                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(SlLine(), i, 1)
                autECLPSObj.SendKeys(PaidDoll, i, 41)
                autECLPSObj.SendKeys(PaidCents, i, 47)
                autECLPSObj.SendKeys(PaidDoll, i + 1, 30)
                autECLPSObj.SendKeys(PaidCents, i + 1, 36)
                autECLPSObj.SendKeys(PaidDoll, i + 1, 61)
                autECLPSObj.SendKeys(PaidCents, i + 1, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", i, 25)
                autECLPSObj.SendKeys("3021", i, 29)
                autECLPSObj.SendKeys("0H", i, 36)
                autECLPSObj.SendKeys("100", i + 1, 55)
                ApiChk()
                Enter()
                ApiChk()
                Exit For
            End If

        Next

        If Pz1(excelrow) Then
            Return True
        End If
        Dim Nys As String
        Dim NysPayment As String
        Dim NysPdDol As String
        Dim NysPdCents As String

        If NysSta = "A" Then

            Nys = MsgBox("NYS lLine, Do U Wanna NYS 30 Line ?", vbYesNo + vbDefaultButton1, "NYS Line")
            If Nys = vbYes Then
                ApiChk()
                NysPayment = InputBox(" Enter the Correct Amount", "Correct Payment Format00000.00 or 0.00", "")

                If NysPayment <> "" And NysPayment <> "0" And NysPayment <> "0.00" Then
                    NysPayment = CStr(NysPayment)
                    If InStr(1, NysPayment, ".") = 0 Then
                        NysPayment = NysPayment & ".00"
                    ElseIf Len(NysPayment) - InStr(1, NysPayment, ".") = 1 Then
                        NysPayment = NysPayment & "0"
                    End If
                    NysPdDol = Left(NysPayment, InStr(1, NysPayment, ".") - 1)
                    NysPdDol = NysPdDol.PadLeft(5, "0")
                    NysPdCents = Right(NysPayment, 2)
                End If
                For i = 10 To 22 Step 2
                    ApiChk()
                    Dim svcCode1 As String = autECLPSObj.GetText(i, 4, 6)
                    If svcCode1 = "------" Then
                        ApiChk()
                        ApiChk()
                        autECLPSObj.SendKeys(SL(line, 1), i, 1)
                        autECLPSObj.SendKeys("NYS", i, 7)
                        ApiChk()
                        autECLPSObj.SendKeys(NysPdDol, i, 41)
                        autECLPSObj.SendKeys(NysPdCents, i, 47)
                        autECLPSObj.SendKeys(NysPdDol, i + 1, 30)
                        autECLPSObj.SendKeys(NysPdCents, i + 1, 36)
                        autECLPSObj.SendKeys(NysPdDol, i + 1, 61)
                        autECLPSObj.SendKeys(NysPdCents, i + 1, 67)
                        ApiChk()
                        autECLPSObj.SendKeys("002", i, 25)
                        autECLPSObj.SendKeys("3021", i, 29)
                        autECLPSObj.SendKeys("0H", i, 36)
                        autECLPSObj.SendKeys("100", i + 1, 55)
                        ApiChk()
                        Enter()
                        ApiChk()
                        Exit For
                    End If

                Next

            End If
        End If

        ApiChk()
        arcChk()
        ReconeChk(excelrow)

    End Function
    Function rcFlot(ByVal excelrow As String)
        If InStr(autECLPSObj.GetText(2, 2, 80), "G8") > 0 Then

            ApiChk()

        Else

            ApiChk()



            If claimType = "UB" Then
                ApiChk()
                autECLPSObj.SendKeys("GD2", 2, 2)
                ApiChk()
                Enter()
                If Pz1(excelrow) Then
                    Return True
                End If

            Else
                ApiChk()
                autECLPSObj.SendKeys("GD1", 2, 2)
                ApiChk()
                Enter()
                If Pz1(excelrow) Then
                    Return True
                End If
            End If
        End If


    End Function


    Function FieldState(Rowx, Coly)
        Dim myField

        autECLPSObj.autECLFieldList.Refresh
        myField = autECLPSObj.autECLFieldList.FindFieldByRowCol(Rowx, Coly)
        If (myField.Protected) Then
            FieldState = True
        Else
            FieldState = False
        End If

    End Function

    Sub RscChk()

        If MainForm.MetroTextBoxRsc.Text = "" Then
            ApiChk()
            autECLPSObj.SendKeys("01", 22, 49)

        Else

            ApiChk()
            Rsc = Trim(MainForm.MetroTextBoxRsc.Text)
            ApiChk()
            autECLPSObj.SendKeys(Rsc, 22, 49)
        End If

    End Sub

    Function CHECK_FOR_EDIT_MPC(ByVal excelrow As String)


        Dim num As String = 1
        Dim claimEdit As String = ""
        Dim editNum As String = ""
        Dim IsoLine As String
        Try



            If InStr(autECLPSObj.gettext(2, 2, 2), "MP") > 0 Then


                Do


                    If InStr(autECLPSObj.GetText(24, 1, 78), "WE") > 0 Then
                        editNum = InStr(autECLPSObj.GetText(24, 1, 78), "WE")
                        editNum = editNum + 1

                        claimEdit = autECLPSObj.GetText(24, CInt(editNum), 5)

                        If claimEdit = "E1434" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)

                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    autECLPSObj.SendKeys("07", i, 29)
                                    Exit For
                                End If

                            Next
                        ElseIf claimEdit = "E1654" Or claimEdit = "E1655" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("07", i, 29)
                                    End If

                                End If

                            Next
                        ElseIf claimEdit = "E2497" Or claimEdit = "E2806" Or claimEdit = "E185 " Then

                            ApiChk()

                            MsgBox("Ck Manual Spi please check and Continue")
                        ElseIf claimEdit = "E1748" Or claimEdit = "E2193" Or claimEdit = "E2193" Or claimEdit = "E2843" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("08", i, 29)
                                    End If

                                End If

                            Next
                        ElseIf claimEdit = "E2565" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("09", i, 29)
                                    End If

                                End If

                            Next

                        ElseIf claimEdit = "E348 " Then

                            Status = "Please follow New Coins "
                            getDateTime(excelrow)
                            Return True
                        ElseIf claimEdit = "E1500" Then

                            Status = "Manual Intervantion Require "
                            getDateTime(excelrow)
                            Return True

                        ElseIf claimEdit = "E2117" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("17", i, 29)
                                    End If

                                End If

                            Next

                        ElseIf claimEdit = "E2666" Then

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("07", i, 29)
                                    End If

                                End If

                            Next

                        ElseIf claimEdit = "E1876" Then
                            ApiChk()
                            Rsc = "17"
                        ElseIf claimEdit = "E1498" Then
                            ApiChk()

                            For i = 10 To 22 Step 2
                                Dim LineDoll As String = autECLPSObj.GetText(i, 41, 5)
                                Dim LineCent As String = autECLPSObj.GetText(i, 47, 2)
                                IsoLine = autECLPSObj.GetText(i, 36, 2)
                                If LineDoll <> "-----" And LineCent <> "--" Then
                                    If IsoLine <> "0H" Or IsoLine <> "B9" Then
                                        autECLPSObj.SendKeys("09", i, 29)
                                    End If

                                End If

                            Next

                        End If


                    End If



                    num += 1
                    Enter()
                Loop Until num = 10
            End If

            '/---------------------------------------------------------------------------


        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Function

    Function CHECK_PaidAmount_MPC(ByVal excelrow As String)

        Dim repeat_line As String
        Dim strMpcLine As String = ""
        Dim arrMpcLine() As String
        Dim iterationCounter As Integer = 0
        Dim chkPaid As String = ""



        Try
            ApiChk()
            ApiChk()

            If InStr(autECLPSObj.gettext(2, 2, 2), "MP") = 0 Then

            End If



RePickMpcLine:


            If iterationCounter <> 15 Then

                ApiChk()

                strMpcLine = strMpcLine & "," & Trim(autECLPSObj.gettext(24, 1, 80))

                iterationCounter = iterationCounter + 1

                Enter()

                GoTo RePickMpcLine

            End If


            Erase arrMpcLine
            arrMpcLine = strMpcLine.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)


            For T As Integer = 0 To arrMpcLine.Length - 1

                repeat_line = ""
                repeat_line = arrMpcLine(T)

                If Microsoft.VisualBasic.Left(repeat_line, 7) = "W169P1-" Or Microsoft.VisualBasic.Left(repeat_line, 7) = "W168P1-" Then

                    repeat_line = repeat_line.Replace(",", "")
                    repeat_line = Trim(repeat_line)

                    chkPaid = Left(repeat_line, 7)

                    claimPaid = Trim(Mid(repeat_line, 8))
                    CurrentPayment = Double.Parse(claimPaid)
                    If CurrentPayment = 0 Then
                        CurrentPayment = 0.00
                    End If
                    Exit For

                End If


            Next


            '/---------------------------------------------------------------------------


        Catch ex As Exception

            Status = "Payment Screen Not Found"
            getDateTime(excelrow)
            Return True
        End Try
        If chkPaid <> "W169P1-" Then
            Status = "Paid Amount Not Found"
            getDateTime(excelrow)
            Return True
        End If

    End Function

    Sub GetMainDraftNo()

        Dim FlagValidDraft As Boolean = False
        Dim FlagRC30 As Boolean = False
        Dim findObject1
        Dim arrDraftNumberMain(), arrVoidDraftNORepeat() As String
        Dim flagExit69_70_71Main As Boolean = False
        Dim FlagAdjNotSignOn As Boolean = False
        Dim flagDraft As Integer
        Dim dxcode
        Dim draft_sci As String
        arrDraftNumberMain = Nothing
        arrVoidDraftNORepeat = Nothing
        draft_sci = ""


        If InStr(autECLPSObj.gettext(24, 1, 79), "DOC PROCESSED") > 0 Then

            Exit Sub
        ElseIf InStr(autECLPSObj.gettext(24, 1, 79), "INVALID") > 0 Then

            Exit Sub
        End If


        If InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "NOT AUTHORIZED") > 0 Then


            GoTo EndOfSub
        ElseIf InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "ADJSTR NOT SIGNED ON,") > 0 Then

            If EmuLogin() Then
                Exit Sub
            End If

        End If

        '''''''''''''''''pull MHI Screen end.........................
        flagDraft = 0
        Dim row
        Dim counter
        Dim i
        Dim servicecode, pos, temp, claim_date, pdamount, strRC, draft_void, strOV, splitDraft As String
        splitDraft = ""

        i = 0
        counter = 0
        strRC = ""
        draft_void = ""
        strOV = ""

        While InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "E065NO MORE CLMS ON FILE,") = 0 Or i = 0
            i = 0

            i = i + 1

            For row = 11 To 23 Step 4

                autECLPSObj.auteclfieldlist.refresh()
                findObject1 = autECLPSObj.auteclfieldlist.FindFieldByRowCol(row, 1)
                If findObject1.display = True Then
                    autECLPSObj.auteclfieldlist.refresh()

                    servicecode = " "
                    pos = " "

                    temp = Trim(autECLPSObj.gettext(row, 1, 1))

                    If temp = "%" Then

                        strRC = strRC & "," & Trim(autECLPSObj.gettext(row - 1, 36, 2))
                        strOV = strOV & "," & Trim(autECLPSObj.gettext(row - 1, 29, 2))


                        dxcode = Trim(autECLPSObj.gettext(7, 4, 5))
                        servicecode = Trim(autECLPSObj.gettext(row - 1, 4, 6))
                        'MsgBox("servicecode = " & servicecode)
                        pos = Trim(autECLPSObj.gettext(row - 1, 1, 2))
                        'MsgBox("pos = " & pos)

                    End If

                    If autECLPSObj.gettext(row + 1, 1, 3) = "ICN" Then

                        If InStr(strRC, "69") > 0 Or InStr(strRC, "70") > 0 Or InStr(strRC, "71") > 0 Or InStr(strRC, "74") > 0 Or InStr(strRC, "75") > 0 Or InStr(strRC, "77") > 0 Or InStr(strRC, "87") > 0 Then
                            strRC = ""
                            strOV = ""
                            Dim flagExit69_70_71 As Boolean = True
                            draft_void = draft_void & Trim(autECLPSObj.gettext(row - 1, 28, 10))
                            GoTo nextline


                        ElseIf InStr(strRC, "74") > 0 Then
                            strRC = ""
                            strOV = ""
                            draft_void = draft_void & Trim(autECLPSObj.gettext(row - 1, 28, 10))
                            GoTo nextline

                        ElseIf InStr(strOV, "C") > 0 Then
                            strRC = ""
                            strOV = ""
                            draft_void = draft_void & Trim(autECLPSObj.gettext(row - 1, 28, 10))
                            GoTo nextline

                        ElseIf InStr(strRC, "30") > 0 Or InStr(strRC, "42") > 0 Or InStr(strRC, "87") > 0 Or InStr(strRC, "WY") > 0 Or InStr(strRC, "W9") > 0 Then
                            FlagRC30 = True
                            Exit Sub
                        End If

                        strRC = ""
                        strOV = ""

                        claim_date = autECLPSObj.gettext(row - 1, 39, 6)

                        pdamount = Val(autECLPSObj.gettext(row, 17, 8))


                        'MsgBox(draft_sci)
                        If InStr(draft_sci, Trim(autECLPSObj.gettext(row - 1, 28, 10))) = 0 Then
                            draft_sci = Trim(autECLPSObj.gettext(row - 1, 28, 10))
                        End If
                        If draft_sci <> "0000000000" Then
                            If InStr(draft_void, draft_sci) > 0 Then
                                GoTo nextline

                            End If
                            FlagValidDraft = True
                            'MsgBox(draft_sci)
                            draft_sci = draft_sci & " "
                            splitDraft = splitDraft & "," & draft_sci
                            ' MsgBox(splitDraft)



                            arrDraftNumber = splitDraft.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)

                        End If


                    End If
                    If InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "E065NO MORE CLMS ON FILE,") > 0 And row = 23 And i = 1 Then
                        'MsgBox("Exit While")
                        Exit While

                    ElseIf row = 23 Then

                        autECLPSObj.sendkeys("[pf8]")

                        ApiChk()
                        If InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "741") > 0 Or InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "CMI REC") > 0 Then
                            autECLPSObj.sendkeys("[pf8]")
                        End If
                        row = 7
                    End If

                Else
                    If InStr(Trim(autECLPSObj.gettext(24, 2, 78)), "E065NO MORE CLMS ON FILE,") > 0 And row = 23 And i = 1 Then
                        'MsgBox("Exit While")
                        Exit While

                    ElseIf row = 23 Then
                        autECLPSObj.sendkeys("[pf8]")
                        ApiChk()
                        row = 7
                    End If
                End If
nextline:
            Next



            If counter <> 2 Then
                counter = 1
            End If

            If i = 100 Then
                Exit While
            End If

        End While



EndOfSub:



    End Sub

    Function multiDraftScrub(ByVal excelrow As String)

        Dim TotalClaimpaid As Decimal = "0.00"
        Dim draftPaid1 As String
        Dim draftPaid As Decimal

        For i As Integer = 0 To arrDraftNumber.Length - 1

            ApiChk()
            autECLPSObj.SendKeys("[Clear]")
            ApiChk()
            autECLPSObj.SendKeys(CltrLine1(), 1, 2)
            ApiChk()
            autECLPSObj.SendKeys("MHI", 1, 2)
            Enter()
            ApiChk()

            For j = 10 To 20 Step 1
                ApiChk()

                If InStr(autECLPSObj.GetText(j, 1, 3), "ICN") > 0 Then
                    ApiChk()
                    draftPaid1 = Trim(autECLPSObj.GetText(j - 1, 16, 9))
                    If draftPaid1 = ".00" Then
                        draftPaid = Convert.ToDecimal("0.00")
                    Else
                        draftPaid = Convert.ToDecimal(draftPaid1)
                    End If

                End If
                ApiChk()
                Enter()
                ApiChk()
            Next

            TotalClaimpaid = TotalClaimpaid + draftPaid
        Next

        For d As Integer = 0 To arrDraftNumber.Length - 1

            Draft = Trim(arrDraftNumber(d))

            ApiChk()
            autECLPSObj.SendKeys("[Clear]")
            ApiChk()
            autECLPSObj.SendKeys(CltrLine1(), 1, 2)
            ApiChk()
            autECLPSObj.SendKeys("MHI", 1, 2)
            Enter()
            ApiChk()

            For i = 10 To 24 Step 1
                If InStr(autECLPSObj.GetText(i, 1, 1), "%") > 0 Then
                    ApiChk()
                    Dim SvC1 As String = autECLPSObj.gettext(i - 1, 4, 6)
                    If SvC1 <> "NCOPAY" Then
                        SvC = autECLPSObj.gettext(i - 1, 4, 6)
                        PoS = autECLPSObj.gettext(i - 1, 1, 2)
                        FDate = autECLPSObj.gettext(i - 1, 11, 6)
                        LDate = autECLPSObj.gettext(i - 1, 18, 6)
                        Remark = Trim(autECLPSObj.gettext(i - 1, 36, 2))
                        isoFound = autECLPSObj.GetText(i - 1, 29, 2)
                        ApiChk()

                    End If
                    If InStr(autECLPSObj.GetText(i, 1, 3), "ICN") > 0 Then
                        If autECLPSObj.GetText(i - 2, 28, 10) <> "0000000000" Then
                            PaidDoll = autECLPSObj.gettext(i - 1, 17, 5)
                            PaidCents = autECLPSObj.gettext(i - 1, 23, 2)
                        End If
                    End If
                End If
            Next

            Dim claimPaid As String = Trim(PaidDoll & "." & PaidCents)

            If claimPaid = "0.00" Or claimPaid = "." Or claimPaid = "0" Or claimPaid = ".00" Then
                ApiChk()

                If MdlZeroPaid.Main(excelrow) Then
                    Return True

                End If

            Else
                ApiChk()

                If MdlUnderPaid.Main(excelrow) Then
                    Return True
                End If

            End If

        Next
        ApiChk()
        fundingChk()
        ApiChk()

        For d As Integer = 0 To arrDraftNumber.Length - 1
            ApiChk()
            clearChk()
            clearChk()

            ApiChk()
            autECLPSObj.SendKeys(CltrLine1(), 1, 2)
            ApiChk()
            autECLPSObj.SendKeys("ADJ", 1, 2)
            Enter()
            ApiChk()
            autECLPSObj.SendKeys("EDS", 2, 2)
            ApiChk()
            Enter()


        Next

    End Function

    Function mpcIntChk(ByVal excelrow As String)
        Dim cxint As String
        If INT = "Yes" And Funding = "1" Or Funding = "2" Then

            For Z = 10 To 22 Step 2
                ApiChk()
                cxint = autECLPSObj.GetText(Z, 4, 5)


                If cxint = "CXINT" Then

                    Exit For
                End If

            Next

            If cxint <> "CXINT" Then


                ApiChk()
                autECLPSObj.SendKeys("EDS", 2, 2)
                Enter()

                ApiChk()
                autECLPSObj.SendKeys("N", 23, 71)
                Enter()
                ApiChk()
                autECLPSObj.SendKeys(Rsc, 22, 49)
                Enter()

                ApiChk()
                autECLPSObj.SendKeys("A", 1, 75)
                Enter()
                ApiChk()
                Onr()
                ApiChk()
                EDsChk = autECLPSObj.GetText(1, 67, 7)
                If EDsChk = "NXT SCR" Then

                    ApiChk()
                    autECLPSObj.SendKeys("A", 1, 75)
                    Enter()
                End If
                If Rsc = "17" Then
                    ApiChk()
                Else

                    If claimType = "UB" Then
                        ApiChk()
                        autECLPSObj.SendKeys("GD2", 2, 2)
                        ApiChk()
                        Enter()
                        If Pz1(excelrow) Then
                            Return True
                        End If

                    Else
                        ApiChk()
                        autECLPSObj.SendKeys("GD1", 2, 2)
                        ApiChk()
                        Enter()
                        If Pz1(excelrow) Then
                            Return True
                        End If
                    End If

                End If

                For y = 10 To 22 Step 2

                    ApiChk()
                    copay = autECLPSObj.GetText(y, 4, 6)

                    If copay = "NCOPAY" Then
                        copayNBR = Trim(autECLPSObj.GetText(y, 25, 3))
                        If copayNBR = "" Or copayNBR = "0" Then
                            copayN = InputBox("Please enter the Copay No..", "Copay", "1")
                            ApiChk()
                            autECLPSObj.SendKeys(copayN, y, 27)
                            Enter()


                        End If
                    End If
                    ApiChk()
                    CopayChk = autECLPSObj.GetText(24, 3, 6)

                    If CopayChk = "E1542C" Then

                        ApiChk()
                        autECLPSObj.SendKeys("02", y, 29)
                        Enter()

                    End If

                Next


            End If

        End If


    End Function

End Module
