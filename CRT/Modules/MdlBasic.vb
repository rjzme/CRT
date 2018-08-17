Module MdlBasic

    Sub main()

        MdlMain.main()


        Dim ExcelRow As Integer
        Dim oldICN As String
        ExcelRow = InputBox("Please Enter the line No. to start", "Row No.", "2")
beHere:

        While worksheet.Cells(ExcelRow, 2).Value <> Nothing

            OVAmount = "0.00"
            MainForm.MetroGridExcelImport.Refresh()
            MainForm.MetroGridUnetInfo.Rows.Clear()
            MainForm.MetroGridHardcopyInfo.Rows.Clear()
            MainForm.MetroTextBoxOCINDB.Clear()
            strICN = Trim(Left(worksheet.Cells(ExcelRow, 2).Value, 10))

            If strICN = oldICN Then
                Status = "Refer to Previous Line "
                getDateTime(ExcelRow)
                ExcelRow += 1
                Continue While
            End If
            oldICN = strICN
            If strICN = "" Then
                MsgBox("DONE")
                Exit Sub
            End If


            strEngine = UCase(Trim(worksheet.Cells(ExcelRow, 3).Value))
            saveCount = saveCount + 1
            If saveCount = 50 Then
                workbook.Save()
                saveCount = 0
            End If

            ApiChk()
            autECLPSObj.SendKeys("[pA1]")
            ApiChk()

            GetIntoEngine(strEngine)

            clearChk()
            ApiChk()
            clearChk()
            ApiChk()


            strCommand = "RET,I" & strICN & ",M"
            autECLPSObj.SendKeys(strCommand, 1, 2)
            Enter()
            ApiChk()


            If InStr(autECLPSObj.gettext(24, 7, 21), "ADJSTR NOT SIGNED ON,") > 0 Then
                ApiChk()
                ApiChk()
                If EmuLogin() Then
                    Exit Sub
                End If
                oldICN = ""
                GoTo beHere
            End If
            ApiChk()
            If InStr(autECLPSObj.GetText(24, 3, 20), "E028NOT AUTHORIZED") Then
                ApiChk()
                Status = "E028NOT AUTHORIZED"
                getDateTime(ExcelRow)
                ExcelRow += 1
                Continue While
            End If

            GetMainDraftNo()

            Try
                If arrDraftNumber.Length > 1 Then
                    Status = "Multi Draft"
                    getDateTime(ExcelRow)
                    ExcelRow += 1
                    Continue While
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            If InStr(autECLPSObj.gettext(2, 2, 2), "MH") > 0 Then
                autECLPSObj.SendKeys("I", 2, 4)
                Enter()
                ApiChk()
                CltrLine = autECLPSObj.gettext(2, 6, 28)
                ApiChk()

                Do

                    For i = 9 To 22 Step 1
                        If InStr(autECLPSObj.gettext(i, 1, 1), "%") > 0 Then
                            ApiChk()
                            Dim PayeeCode As String
                            Dim SvC1 As String = autECLPSObj.gettext(i - 1, 4, 6)
                            If SvC1 <> "NCOPAY" Then
                                SvC = autECLPSObj.gettext(i - 1, 4, 6)
                                PoS = autECLPSObj.gettext(i - 1, 1, 2)
                                FDate = autECLPSObj.gettext(i - 1, 11, 6)
                                LDate = autECLPSObj.gettext(i - 1, 18, 6)
                                Remark = Trim(autECLPSObj.gettext(i - 1, 36, 2))
                                isoFound = autECLPSObj.GetText(i - 1, 29, 2)
                                PayeeCode = Trim(autECLPSObj.gettext(i - 1, 31, 1))
                                ApiChk()

                                If PayeeCode = "Z" Or PayeeCode = "1" Then
                                    ApiChk()
                                    Status = "Claim paid to member Please check"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If

                                If SvC = "OI" Or SvC = "OIM" Or SvC = "OIMEDI" Then
                                    ApiChk()
                                    Status = SvC & " Claim Please review"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If
                                If isoFound = "30" Then
                                    ApiChk()
                                    isoPaid = Trim(autECLPSObj.GetText(i, 61, 8))
                                End If
                                If Remark = "69" Or Remark = "70" Then
                                    ApiChk()
                                    Status = "Already Voided Claim Please check"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If
                                If Remark = "YP" Then
                                    ApiChk()
                                    Status = "URN Claim"
                                End If
                                If Remark = "7Y" Then
                                    ApiChk()
                                    Status = "Closed Claim"
                                End If
                                If Remark = "OV" Then
                                    ApiChk()
                                    Status = "Manual Rov in this Claim Please review"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If
                                If Remark = "75" Or Remark = "77" Then
                                    ApiChk()
                                    Status = Remark & "In this claim please review"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If
                                If Remark = "E5" Or Remark = "CA" Or Remark = "LO" Or Remark = "05" Then
                                    ApiChk()
                                    Status = "Corrected Claim/ Duplicate Claim"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If
                                If Remark = "07" Or Remark = "WY" Or Remark = "06" Or Remark = "08" Then
                                    ApiChk()
                                    Status = "Claim Correctly denied for Member Term"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While
                                End If

                                billedChargeDollar = Trim(autECLPSObj.GetText(i - 1, 41, 5))
                                billedChargeCent = autECLPSObj.GetText(i - 1, 47, 2)
                            End If

                        End If
                        ApiChk()


                        If InStr(autECLPSObj.GetText(i, 1, 3), "ICN") > 0 Then

                            If autECLPSObj.GetText(i - 2, 28, 10) <> "0000000000" Then
                                Draft = autECLPSObj.gettext(i - 2, 28, 10)
                                Tin = autECLPSObj.gettext(i - 2, 5, 9)
                                Pre = autECLPSObj.gettext(i - 2, 3, 1)
                                ClaimSuff = autECLPSObj.gettext(i - 2, 15, 5)
                                Suffix = ClaimSuff
                                PaidDoll = autECLPSObj.gettext(i - 1, 17, 5)
                                PaidCents = autECLPSObj.gettext(i - 1, 23, 2)


                                Policy = autECLPSObj.gettext(2, 6, 6)
                                UBH = autECLPSObj.GetText(8, 42, 1)

                                SSN = autECLPSObj.gettext(2, 14, 9)
                                Name = autECLPSObj.GetText(2, 24, 10)
                                Rel = autECLPSObj.GetText(2, 35, 2)
                                Dx1 = Trim(autECLPSObj.gettext(7, 4, 6))

                                Dx2 = Trim(autECLPSObj.gettext(7, 15, 6))
                                Fln = autECLPSObj.gettext(i, 34, 10)
                                DCC = autECLPSObj.gettext(i, 49, 3)
                                FlnDcc = Fln & DCC

                                MainForm.MetroGridUnetInfo.Rows.Add(New String() {" ", " ", SvC, FDate, LDate, dayUnit, billedChargeDollar & "." & billedChargeCent})

                                If correctedInfo = "7" Or correctedInfo = "137" Then
                                    Status = "Corrected Claim Please Review"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While

                                End If

                                ApiChk()
                                If UBH = "1" Or UBH = "6" Then
                                    ApiChk()
                                    Status = "UBH Claim"
                                    getDateTime(ExcelRow)
                                    ExcelRow += 1
                                    Continue While

                                End If
                                Exit Do
                            End If

                        End If

                    Next
                    Enter()
                Loop

            End If
            ApiChk()


            GoToCei()

            Try
                MdlPullHardCopyInside.OpenHardCopy(MainForm.MetroTextBoxEDSSID.Text, MainForm.MetroTextBoxEDSSPassword.Text, FlnDcc)
                MdlScrapHardCopyData.GetHcfa(FrmCRT)
                MainForm.MetroTabControlMain.SelectedTab = MainForm.MetroTabPageUnetInfo
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ApiChk()
            autECLPSObj.SendKeys("[pf12]")
            ApiChk()

            clearChk()
            ApiChk()
            clearChk()
            ApiChk()



            If ChkPMI(ExcelRow) Then
                ApiChk()
                ExcelRow += 1
                Continue While

            Else
                If MainForm.MetroRadioButtonAYes.Checked Then

                    If claimType = "UB" Then
                        If RT = "R" Then
                            ApiChk()
                            Status = "Need HSS Pricing Allowable TO Process This Claim"
                            getDateTime(ExcelRow)
                            ExcelRow += 1
                            Continue While
                        Else
                            ApiChk()
                            clearChk()
                            ApiChk()
                            autECLPSObj.SendKeys(CltrLine1(), 2, 2)
                            autECLPSObj.SendKeys("ADJ", 2, 2)
                            Enter()
                            ApiChk()
                            MdlOCI.OpenOCI(FlnDcc, MainForm.MetroTextBoxEmulator.Text)
                            ApiChk()
                            autECLPSObj.SendKeys("OCC", 2, 2)
                            Enter()
                            ApiChk()
                            ClaimAllowable = autECLPSObj.GetText(22, 60, 10)
                            ApiChk()
                            MainForm.MetroTabControlMain.SelectedTab = MainForm.MetroTabPageOCINDB

                            MainForm.MetroTextBoxOCINDB.Text = GetUnetScreen()


                        End If

                    Else

                        ApiChk()
                        MdlNDB.NDBFee()

                    End If
                End If
                clearChk()
                ApiChk()
                clearChk()
                ApiChk()
                ApiChk()



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

                        If MdlZeroPaid.Main(ExcelRow) Then

                            ExcelRow += 1
                            Continue While
                        End If

                    Else
                        ApiChk()


                        If MdlUnderPaid.Main(ExcelRow) Then
                            ExcelRow += 1
                            Continue While
                        End If

                    End If

                    ApiChk()

                    ApiChk()
                    If MdlPaymentScreen.BuildADJ(ExcelRow) Then
                        ExcelRow += 1
                        Continue While

                    End If

                Next


            End If




            ExcelRow += 1

        End While
        ApiChk()
        MsgBox("Done")
        workbook.Save()
        'CreateDatabase()
        workbook.Close()
        APP.Quit()


    End Sub
    Sub GoToCei()
        Dim NameInfo As String
        Dim RelInfo As String
        Dim Cei As String
        Dim NextChk As String
        Dim DateChk As String
        Dim Dos As String
        ApiChk()
        clearChk()
        Cei = "Cei," & Policy & ",S" & SSN
        autECLPSObj.SendKeys(Cei, 2, 2)
        Enter()
        ApiChk()

        For RowChk = 9 To 22 Step 4
            ApiChk()
            NameInfo = Trim(autECLPSObj.GetText(RowChk, 3, 12))
            RelInfo = Trim(autECLPSObj.GetText(RowChk, 16, 2))
            DateChk = autECLPSObj.GetText(RowChk + 2, 23, 2)
            Dos = Right(FDate, 2)
            If NameInfo = Name And RelInfo = Rel Then
                ApiChk()
                If DateChk >= Dos Then
                    ApiChk()
                    RowChk = RowChk + 2
                Else
                    RowChk = RowChk + 1
                End If
                SetNo = autECLPSObj.GetText(RowChk, 41, 3)
                If SetNo = "000" Then
                    SetNo = autECLPSObj.GetText(RowChk, 45, 3)
                    If SetNo = "000" Then
                        SetNo = autECLPSObj.GetText(RowChk, 49, 3)
                    End If
                End If
                Exit For
            ElseIf RowChk = 22 Then
                ApiChk()
                Enter()
                NextChk = autECLPSObj.GetText(24, 8, 7)
                If NextChk = "No More" Then
                    Exit For
                End If
            End If

        Next

        'MsgBox(SetNo)
    End Sub

    Sub OpnIbaag()
        Dim OpenIbaagStr As String
        ApiChk()

        DateEff = Left(FDate, 2) & "/" & Mid(FDate, 3, 2) & "/20" & Right(FDate, 2)

        OpenIbaagStr = "http://utsa.uhc.com/srch_rslts.asp?policy_number=" & Policy & " &set_number=" & SetNo & " &eff_date=" & DateEff & "&show_revision=&doc_type=T&sys_id=ORS"

        OpenIEWindow(OpenIbaagStr)

    End Sub
    Sub OpenIEWindow(ByVal strlocation As String)
        Dim objIE
        'Create an instance of the Internet explorer object also works with Excel, Word, Powerpoint etc.
        objIE = CreateObject("InternetExplorer.Application")
        'Make object visible
        objIE.Visible = True

        objIE.Navigate(strlocation)
    End Sub


End Module
