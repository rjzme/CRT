Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Module MdlMain
    Public autECLPSObj As Object
    Public autECLOIAObj As Object
    Public autECLConnList As Object
    Public strEngine As String
    Public strICN As String
    Public strCommand As String
    Public line, i, j As Integer
    Public UnetSSN, UnetReady, WinTitle, void_found, isoFound As String
    Public Ors, FlnDcc, Pol, Fln, DCC, cpt, modCode, charges, dayUnit, plcSv, dos_from, rev_code, dos_to, UBH As String
    Public Dx1, Dx2, Dx3, Dx4, Dx5 As String
    Public strEngineNo As String
    Public strpmi, ChkPTI, ChkNYS, DupChk, EdsOV, svcFound, cntrLine1, strEngineN As String
    Public PaidDoll, PaidCents, Policy, SetNo, billedChargeDollar, billedChargeCent As String
    Public Draft, Tin, Suffix, Pre, Flag, SSN, var3, var4 As String
    Public CltrLine, Rel, Name As String
    Public APP As New Excel.Application
    Public worksheet As Excel.Worksheet
    Public workbook As Excel.Workbook
    Public saveCount As String
    Public claimType, correctedInfo, ClaimAllowable, claimPaid, copayN, CopayChk, copay, copayNBR As String
    Public SlLine1 As String
    Public RemarkColl(7), CodesCol(7), arrDraftNumber() As String
    Public CodesColl, DraftColl(), Remark As String
    Public LC, CVoid, CDone, PoS, SvC, FDate, LDate, DateEff As String
    Public respo, cntrLine, notCovDol, notCovCnts, notBal, notBalDol, notBalChg, IPHClaim As String
    Public PdDol, PdCents, adj, totdol, totchg, Test, adjline, Rdy4MPP, wrscr, TestEdit As String
    Public ClaimNum, SL(7, 6), Prov835, PatNum, verifyEDS, TinRespo As String
    Public adjNumber, strNTID, myTime, StartTime, EndTime As String
    Public BLKChk, INT, INThold, INTuse, IRC, LineDenialCode As String
    Public Plan, Plan1, Policy1, Interest, Funding, SState, Erisa As String
    Public NysSta, strCntrLine1, VDenial, EdsTest As String
    Public LineRemark, RconChk As String
    Public SvCode, Sline, FDline, LDline, Overpaid, OverPaidDollars, OverPaidCents As String
    Public isoPaid As Double
    Public payment, CurrentPayment As Decimal
    Public RMCode, RCCode, Rsc, OonR, Flt, ModF, EDsChk, chkCauseCode As String
    Public RemarkChk, DenialCode, strDOS, strTIN, strTINsfx, strDraft, FEEschd, RT As String
    Public strCtrlline, strPPIctrlline, strProduct, strType, strMarket, strIPA, strGroupTab As String
    Public dateDOS, datePPIeffdt, datePPIcancdt As Date
    Public proDate As String = ""
    Public proTime As String = ""
    Public ClaimPaidAmount As String = ""
    Public ClaimIntAmount As String = ""
    Public connString, dbFile, SSPCodes, voidEdit As String
    Public ClaimSuff, OVAmount, ExitReason, Status As String
    Public excelFile As String
    Public conn As OleDbConnection

    Function main() As Boolean
        Try
            workbook = APP.Workbooks.Open(excelFile)
            worksheet = workbook.Worksheets("sheet1")
        Catch ex As Exception
            MsgBox("Please Insert Data")
            Return False
        End Try

        Return True

    End Function
    ' Sub ApiChk()

    '    autECLOIAObj.WaitForInputReady

    'End Sub
    Sub ApiChk()
        Do
            System.Threading.Thread.Sleep(50)

        Loop Until autECLOIAObj.Inputinhibited = 0

    End Sub


    Sub GetIntoEngine(engine)
        autECLPSObj.SendKeys("[pa1]")
        ApiChk()

        Select Case UCase(engine)

            Case "A"

                autECLPSObj.SendKeys("01", 17, 42)
                ApiChk()
            Case "B"

                autECLPSObj.SendKeys("02", 17, 42)
                ApiChk()
            Case "E"

                autECLPSObj.SendKeys("03", 17, 42)
                ApiChk()
            Case "N"
                autECLPSObj.SendKeys("04", 17, 42)
                ApiChk()
            Case "S"

                autECLPSObj.SendKeys("05", 17, 42)
                ApiChk()
            Case "W"

                autECLPSObj.SendKeys("06", 17, 42)
                ApiChk()
            Case "C"

                autECLPSObj.SendKeys("07", 17, 42)
                ApiChk()
            Case "G"

                autECLPSObj.SendKeys("08", 17, 42)
                ApiChk()
            Case "O"

                autECLPSObj.SendKeys("09", 17, 42)
                ApiChk()
            Case "Q"

                autECLPSObj.SendKeys("10", 17, 42)
                ApiChk()
            Case "Y"

                autECLPSObj.SendKeys("11", 17, 42)
                ApiChk()
            Case "Z"

                autECLPSObj.SendKeys("12", 17, 42)
                ApiChk()
            Case "X"

                autECLPSObj.SendKeys("13", 17, 42)
                ApiChk()
            Case "K"

                autECLPSObj.SendKeys("14", 17, 42)
                ApiChk()
            Case "D"

                autECLPSObj.SendKeys("15", 17, 42)
                ApiChk()
            Case "F"

                autECLPSObj.SendKeys("16", 17, 42)
                ApiChk()
            Case "M"

                autECLPSObj.SendKeys("17", 17, 42)
                ApiChk()
            Case "U"

                autECLPSObj.SendKeys("18", 17, 42)
                ApiChk()
            Case "H"
                autECLPSObj.SendKeys("19", 17, 42)
                ApiChk()
            Case "J"

                autECLPSObj.SendKeys("20", 17, 42)
                ApiChk()
            Case "R"

                autECLPSObj.SendKeys("21", 17, 42)
                ApiChk()
            Case "L"

                autECLPSObj.SendKeys("22", 17, 42)
                ApiChk()
        End Select
        Enter()
        ApiChk()

    End Sub

    Sub Connection()

        Dim Session As String = MainForm.MetroTextBoxEmulator.Text

        autECLPSObj = CreateObject("PCOMM.auteclps")
        autECLConnList = CreateObject("PCOMM.auteclconnlist")
        autECLOIAObj = CreateObject("PCOMM.autecloia")

        If autECLConnList.FindConnectionByName(Session) Is Nothing Then

            MetroFramework.MetroMessageBox.Show(MainForm, "Emulator not found")

            Exit Sub
        End If

        autECLPSObj.SetConnectionByName(Session)
        autECLOIAObj.SetConnectionByName(Session)
        autECLPSObj.auteclfieldlist.Refresh()
        autECLConnList.Refresh()
        'EmuLogin()
    End Sub


    Function ChkPMI(ByVal excelrow As String)
        Dim Provider As String
        Dim Suffix1 As String

        strpmi = "PMI,"
        autECLPSObj.SendKeys(strpmi, 1, 2)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys(Pre, 3, 6)
        autECLPSObj.SendKeys(Tin, 3, 8)
        autECLPSObj.SendKeys(Suffix, 3, 22)
        Enter()
        ApiChk()
        Flag = autECLPSObj.gettext(3, 33, 1)
        ApiChk()
        ChkNYS = autECLPSObj.gettext(12, 38, 1)
        Suffix1 = MainForm.MetroTextBoxSuffix.Text

        If Suffix1 = "" Then

            If Flag = "1" Then
                Suffix = Trim(InputBox("Flag 1 Please Input the Correct Suffix",, ""))
                ApiChk()
                If Suffix = "" Then
                    ApiChk()
                    Status = "Flag 1 Please Check"
                    getDateTime(excelrow)
                    Return True

                End If

                Return False

            End If


        Else
            ApiChk()
            Suffix = Suffix1

        End If
        clearChk()
        strpmi = "PMI,"
        autECLPSObj.SendKeys(strpmi, 1, 2)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys(Pre, 3, 6)
        autECLPSObj.SendKeys(Tin, 3, 8)
        autECLPSObj.SendKeys(Suffix, 3, 22)
        Enter()
        ApiChk()
        ChkPTI = autECLPSObj.gettext(9, 63, 1)
        Provider = autECLPSObj.gettext(8, 7, 1)
        If ChkPTI = "G" Then
            If claimType = "HCFA" Then
                If Provider = ";" Then
                    If Suffix1 = "" Then

                        Suffix = Trim(InputBox("Group Provider With PTI G, Please Input Correct Suffix ", Suffix))

                        If Suffix = "" Then

                            Status = "Group Provider With PTI G please review."
                            getDateTime(excelrow)

                            Return True

                        End If
                    End If


                End If

            End If
        End If
    End Function


    Sub ImportExcel(excelfile)

        Dim ds As DataSet = New DataSet
        Dim conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelfile + ";Extended Properties=Excel 12.0;")
        Dim dta As OleDbDataAdapter = New OleDbDataAdapter("Select * From [Sheet1$]", conn)


        dta.Fill(ds, "[Sheet1$]")


        MainForm.MetroGridExcelImport.DataSource = ds
        MainForm.MetroGridExcelImport.DataMember = "[Sheet1$]"
        conn.Close()


    End Sub

    Function EmuLogin()

        Dim Invalid As String
        Dim Engine As String
        Dim ChkScrn As String

        UnetSSN = "Soi," & MainForm.MetroTextBoxUnetID.Text


        UnetReady = autECLPSObj.gettext(2, 1, 7)
        ApiChk()
        ApiChk()

        If UnetReady <> "UHC0010" Then
            ApiChk()
            clearChk()
            ApiChk()
        End If

        If UnetReady <> "UHC0010" Then
            ApiChk()
            autECLPSObj.SendKeys("clear-transid", 1, 2)
            Enter()

            ApiChk()
            autECLPSObj.SendKeys("off", 1, 2)
            Enter()
            ApiChk()

        End If

        ApiChk()
        autECLPSObj.SendKeys("topssea", 3, 1)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys("sea1", 1, 1)
        Enter()
        ApiChk()

        For X = 1 To 22

            Engine = getEngine(X)
            ApiChk()
            autECLPSObj.SendKeys(X, 17, 42)
            Enter()

            ApiChk()
            Invalid = autECLPSObj.GetText(21, 12, 7)
            If Invalid = "INVALID" Or Invalid = "SYSTEM " Then

                autECLPSObj.SendKeys(X, 17, 42)
            Else

                ApiChk()
                autECLPSObj.SendKeys(UnetSSN, 2, 2)
                Enter()
                ApiChk()
                autECLPSObj.SendKeys(MainForm.MetroTextBoxUnetPassword.Text, 6, 19)
                autECLPSObj.SendKeys("323", 6, 46)
                autECLPSObj.SendKeys("323", 8, 24)
                autECLPSObj.SendKeys(Engine, 8, 41)
                Enter()
                ApiChk()

                If X < 3 Then
                    ChkScrn = autECLPSObj.GetText(24, 3, 5)
                    If ChkScrn <> "W018A" Then
                        MsgBox("Something is wrong with the password or SSN - Please start over")
                        Return True
                    End If
                End If
                ApiChk()
                autECLPSObj.SendKeys("[PA1]", 8, 41)
                ApiChk()

                If X = 22 Then
                    ApiChk()
                    autECLPSObj.SendKeys("A", 17, 42)
                    Enter()
                    ApiChk()
                End If
                ApiChk()

            End If

        Next

    End Function

    Function getEngine(y)
        Select Case y
            Case 1
                getEngine = "A"
            Case 2
                getEngine = "B"
            Case 3
                getEngine = "E"
            Case 4
                getEngine = "N"
            Case 5
                getEngine = "S"
            Case 6
                getEngine = "W"
            Case 7
                getEngine = "C"
            Case 8
                getEngine = "G"
            Case 9
                getEngine = "O"
            Case 10
                getEngine = "Q"
            Case 11
                getEngine = "Y"
            Case 12
                getEngine = "Z"
            Case 13
                getEngine = "X"
            Case 14
                getEngine = "K"
            Case 15
                getEngine = "D"
            Case 16
                getEngine = "F"
            Case 17
                getEngine = "M"
            Case 18
                getEngine = "U"
            Case 19
                getEngine = "H"

            Case 20
                getEngine = "J"

            Case 21
                getEngine = "R"

            Case 22
                getEngine = "L"

            Case Else
                getEngine = "A"
        End Select
    End Function
    Function CltrLine1() As String
        ApiChk()
        Dim CltrLine12 As String
        CltrLine12 = "MPI," & CltrLine & "," & Rel & "," & Draft & ",I" & strICN
        Return CltrLine12
    End Function
    Function SlLine() As String
        ApiChk()

        SlLine1 = PoS & SvC & FDate & LDate & "001"

        Return SlLine1
    End Function


    Sub ChkPPI()


        clearChk()
        ApiChk()
        ApiChk()
        autECLPSObj.SendKeys(CltrLine1(), 2, 2)
        autECLPSObj.SendKeys("ADJ", 2, 2)
        Enter()
        ApiChk()

        If autECLPSObj.GetText(2, 2, 2) = "MP" Then
            If autECLPSObj.GetText(10, 11, 2) <> "--" Then
                strDOS = autECLPSObj.GetText(10, 11, 6)
            Else
                strDOS = autECLPSObj.GetText(12, 11, 6)
            End If
            Gather_PPI_control_line_info()
            strCtrlline = autECLPSObj.GetText(2, 1, 78)

        End If


        If autECLPSObj.GetText(2, 2, 2) = "MH" Then
            strDOS = autECLPSObj.GetText(10, 11, 6)

            For LnCnt = 10 To 22 Step 2
                If autECLPSObj.GetText(LnCnt + 2, 1, 3) = "ICN" Then
                    strDraft = autECLPSObj.GetText(LnCnt, 28, 10)
                    strICN = autECLPSObj.GetText(LnCnt + 2, 5, 10)
                    strTIN = autECLPSObj.GetText(LnCnt, 3, 1) & autECLPSObj.GetText(LnCnt, 5, 9)
                    strTINsfx = autECLPSObj.GetText(LnCnt, 15, 5)
                    Exit For
                End If
                If LnCnt >= 22 And autECLPSObj.GetText(24, 3, 1) <> "E" Then
                    LnCnt = 8
                    Enter()
                    ApiChk()
                End If
            Next
            autECLPSObj.SendKeys("adj", 2, 2)
            Dim colx As Integer
            For colx = 24 To 34
                If autECLPSObj.GetText(2, colx, 1) = "," Then
                    Exit For
                End If
            Next
            autECLPSObj.SendKeys("," & strDraft & ",i" & strICN & "[erase eof]", 2, colx + 3)
            Enter()
            ApiChk()
            Gather_PPI_control_line_info()
        End If


        dateDOS = CDate(Left(strDOS, 2) & "/" & Mid(strDOS, 3, 2) & "/" & Right(strDOS, 2))

        autECLPSObj.SendKeys("[PF12]")
        ApiChk()
        autECLPSObj.SendKeys("PRI," & strTIN & strTINsfx & "[erase eof]", 2, 2)
        ApiChk()
        Enter()
        ApiChk()
        'strProduct = InputBox("What is the product?",,strProduct)
        'strType = InputBox("What is the type?  If no type please enter 00",,strType)
        'strMarket = InputBox("What is the market Number?(3 digits please)",,strMarket)
        'strIPA = InputBox("What is the IPA Number?",,strIPA)
        'strGroupTab = InputBox("What is the group tab #? (blank if none)")

        strPPIctrlline = "PPI," & strTIN & strTINsfx & "," & strProduct & "," & strType & "," & strMarket & "," & strIPA & "," & strGroupTab
        autECLPSObj.SendKeys(strPPIctrlline & "[erase eof]", 2, 2)
        Enter()
        ApiChk()

        Do
            If autECLPSObj.GetText(5, 9, 2) = "--" Or autECLPSObj.GetText(5, 9, 2) = "99" Then
                datePPIeffdt = CDate("01/01/1970")
            Else
                datePPIeffdt = CDate(autECLPSObj.GetText(5, 9, 2) & "/" & autECLPSObj.GetText(5, 12, 2) & "/" & autECLPSObj.GetText(5, 15, 4))
            End If
            If autECLPSObj.GetText(5, 30, 2) = "--" Or autECLPSObj.GetText(5, 30, 2) = "99" Then
                datePPIcancdt = CDate("12/31/2025")
            Else
                datePPIcancdt = CDate(autECLPSObj.GetText(5, 30, 2) & "/" & autECLPSObj.GetText(5, 33, 2) & "/" & autECLPSObj.GetText(5, 36, 4))
            End If

            If dateDOS >= datePPIeffdt And dateDOS <= datePPIcancdt Then

                ApiChk()

                ApiChk()
                If claimType = "UB" Then
                    ApiChk()
                    RT = autECLPSObj.GetText(6, 61, 1)
                    FEEschd = autECLPSObj.GetText(3, 75, 5)
                Else
                    ApiChk()
                    RT = autECLPSObj.GetText(9, 9, 1)
                    FEEschd = autECLPSObj.GetText(9, 12, 5)
                End If

                Exit Do
            End If
            ApiChk()


            If autECLPSObj.GetText(2, 2, 3) = "PPU" Then
                MsgBox("Unable to find correct PPI record")
                Exit Do
            End If
            Enter()


        Loop



    End Sub

    Sub Gather_PPI_control_line_info()
        Dim strEditHold
        Dim numW1616loc As Integer
        strEditHold = ""
        Do
            If InStr(1, autECLPSObj.GetText(24, 1, 78), "W1616") > 0 Then
                numW1616loc = InStr(1, autECLPSObj.GetText(24, 1, 78), "W1616")

                numW1616loc = numW1616loc + 9   'beginning of product text
                If numW1616loc > 77 Then
                    Enter()
                    ApiChk()
                    numW1616loc = 12
                End If
                strProduct = Trim(autECLPSObj.GetText(24, numW1616loc, 4))

                numW1616loc = numW1616loc + 6   'beginning of type text
                If numW1616loc > 79 Then
                    Enter()
                    ApiChk()
                    numW1616loc = 4
                End If
                strType = autECLPSObj.GetText(24, numW1616loc, 2)

                numW1616loc = numW1616loc + 4   'beginning of market text
                If numW1616loc > 74 Then
                    Enter()
                    ApiChk()
                    numW1616loc = 4
                End If
                strMarket = autECLPSObj.GetText(24, numW1616loc, 7)

                numW1616loc = numW1616loc + 9   'beginning of IPA text
                If numW1616loc > 76 Then
                    Enter()
                    ApiChk()
                    numW1616loc = 4
                End If
                strIPA = autECLPSObj.GetText(24, numW1616loc, 5)

                Exit Do
            End If
            If autECLPSObj.GetText(2, 2, 2) = "AR" Then
                autECLPSObj.SendKeys("S", 3, 5)
            End If
            If InStr(1, autECLPSObj.GetText(2, 1, 76), "PZ1") > 0 Then
                'prevent dup pend
                autECLPSObj.SendKeys("[erase eof]", 2, InStr(1, autECLPSObj.GetText(2, 1, 76), "PZ1"))
                Dim SvcLine As String
                For x = 10 To 22 Step 2

                    SvcLine = autECLPSObj.GetText(x, 4, 6)
                    If SvcLine <> "------" Then

                        autECLPSObj.SendKeys("01", x, 29)

                    End If


                Next
            End If
            If autECLPSObj.GetText(10, 29, 1) = "P" And (autECLPSObj.GetText(10, 36, 2) <> "80" And autECLPSObj.GetText(10, 36, 2) <> "82") Then
                'prevent system auto pend (P-T1, P-PH, etc), allow PISL lines
                Exit Do
            End If
            If autECLPSObj.GetText(12, 29, 1) = "P" And (autECLPSObj.GetText(12, 36, 2) <> "80" And autECLPSObj.GetText(12, 36, 2) <> "82") Then
                Exit Do
            End If
            If autECLPSObj.GetText(24, 3, 5) = strEditHold Or InStr(1, autECLPSObj.GetText(24, 1, 76), "W118") > 0 Then
                'exit if edit/warn line is repeating due to edit or has OK to pay
                Exit Do
            End If
            If autECLPSObj.GetText(24, 3, 1) = "E" Then
                strEditHold = autECLPSObj.GetText(24, 3, 5)
            End If
            Enter()
            ApiChk()
        Loop
    End Sub

    Function GetUnetScreen() As String

        Dim strResult As String
        strResult = ""

        For i = 1 To 24

            If strResult = "" Then
                strResult = autECLPSObj.GetText(i, 1, 80) & vbCrLf
            Else
                strResult = strResult & autECLPSObj.GetText(i, 1, 80) & vbCrLf
            End If



        Next

        Return strResult

    End Function

    Sub Enter()

        autECLPSObj.SendKeys("[enter]")
        ApiChk()

    End Sub



    Function VoidChk()

        Dim i As Integer = 1

        Do
            If InStr(autECLPSObj.GetText(24, 1, 78), "118") > 0 Then

                Return True
                Exit Do
            End If

            Enter()
            i += 1
        Loop Until i = 5

        Return False
    End Function

    Sub GetClaimInfo(ByVal ExcelRow As String)

        ClaimIntAmount = "0.00"

        ApiChk()
        clearChk()
        ApiChk()

        autECLPSObj.SendKeys(strCommand, 1, 2)
        Enter()
        Dim inces As String = 1
        Dim intInd As String
        Do
            For Info = 9 To 23 Step 1

                If InStr(autECLPSObj.GetText(Info, 1, 3), "ICN") > 0 Then
                    ApiChk()
                    If autECLPSObj.GetText(Info - 2, 28, 10) <> "0000000000" Then
                        ClaimPaidAmount = Trim(autECLPSObj.GetText(Info - 1, 16, 9))
                        intInd = Trim(autECLPSObj.GetText(Info, 28, 1))
                    End If
                End If

                If intInd = "Y" Then

                    If InStr(autECLPSObj.GetText(Info, 4, 5), "CXINT") > 0 Then
                        ApiChk()
                        ClaimIntAmount = Trim(autECLPSObj.GetText(Info, 41, 8))
                        Exit Do
                    End If
                ElseIf intInd = "N" Then
                    ApiChk()
                    Exit Do
                End If

            Next
            inces += 1
            Enter()
        Loop Until inces = 3

        worksheet.Cells(ExcelRow, 6).Value = ClaimPaidAmount
        worksheet.Cells(ExcelRow, 7).Value = ClaimIntAmount
        worksheet.Cells(ExcelRow, 8).Value = OVAmount
        Status = "Claim Processed Correctly"
        getDateTime(ExcelRow)

    End Sub

    Sub getDateTime(ByVal ExcelRow As String)

        Try

            Dim dateTime As DateTime = DateTime.UtcNow
            Dim IndianTime As DateTime = dateTime.AddHours(5.5)

            Dim proDate As String = Trim(Left(IndianTime, 10))
            Dim proTime As String = Trim(Right(IndianTime, 10))

            worksheet.Cells(ExcelRow, 4).Value = Policy
            worksheet.Cells(ExcelRow, 5).Value = SSN
            worksheet.Cells(ExcelRow, 9).Value = Status
            worksheet.Cells(ExcelRow, 10).Value = proDate
            worksheet.Cells(ExcelRow, 11).Value = proTime

            CreateDatabase(proDate, proTime)
            workbook.Save()
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Public Function CreateAccessDatabase(ByVal DatabaseFullPath As String) As Boolean
        Dim bAns As Boolean
        Dim cat As New ADOX.Catalog()
        Try

            Dim sCreateString As String
            sCreateString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseFullPath
            cat.Create(sCreateString)

            bAns = True

        Catch Excep As System.Runtime.InteropServices.COMException
            bAns = False

        Finally
            cat = Nothing
        End Try
        Return bAns
    End Function

    Sub CreateDatabase(ByRef strdate As String, ByRef strtime As String)
        Dim wsShell = CreateObject("WScript.Shell")
        Dim citrixUsername = wsShell.ExpandEnvironmentStrings("%USERNAME%")

        Dim dateTime As DateTime = DateTime.UtcNow
        Dim cstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
        Dim CSTdateTime As DateTime = TimeZoneInfo.ConvertTimeFromUtc(dateTime, cstZone)
        Dim IndianTime As DateTime = dateTime.AddHours(5.5)

        Dim dbPath As String = "\\unpiox53pn\Kolkata_Claim_Drive\Sharique\CRT\DB\" & CSTdateTime.Date.ToString("dd-MM-yyyy") & "\"

        If (Not System.IO.Directory.Exists(dbPath)) Then
            System.IO.Directory.CreateDirectory(dbPath)
        End If

        dbFile = dbPath & citrixUsername & ".mdb"

        If System.IO.File.Exists(dbFile) Then
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & dbFile & ";Jet OLEDB:Database Password='magic'"
            InsertData(strdate, strtime)
        Else
            CreateAccessDatabase(dbFile)
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & dbFile & ";Jet OLEDB:Database Password='magic'"
            createTable(strdate, strtime)

        End If

    End Sub

    Sub createTable(ByRef strdate As String, ByRef strtime As String)
        conn = New OleDbConnection(connString)
        conn.Open()

        Dim qryStr As String = "CREATE TABLE CRTWork (ID COUNTER, [ICN] TEXT, [ENGINE] TEXT, [POLICY] TEXT, [SSN] TEXT, [PaidAmount] TEXT, [InterestAmount] TEXT, [OverPaidAmount] TEXT, [Status] TEXT, [Date] TEXT, [Time] TEXT)"
        Dim cmd As OleDbCommand = New OleDbCommand(qryStr, conn)


        Try
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        conn.Close()

        InsertData(strdate, strtime)
        '----------------------'

        'Dim excelConn As OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFile & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1""")

        'Dim insStr As String = "INSERT INTO [MS Access;Database=" & dbFile & ";].[CRTWork] Select * FROM [Sheet1$]"

        'excelConn.Open()

        'Dim cmdIns As OleDbCommand = New OleDbCommand(insStr, excelConn)

        'Try
        '    cmdIns.ExecuteNonQuery()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        'excelConn.Close()

    End Sub
    Sub InsertData(ByRef strdate As String, ByRef strtime As String)
        conn = New OleDbConnection(connString)
        conn.Open()

        Dim qryStr1 As String = "INSERT INTO CRTWork ([ICN], [ENGINE], [POLICY], [SSN], [PaidAmount], [InterestAmount], [OverPaidAmount], [Status], [Date], [Time]) VALUES (@strICN, @strEngine, @Policy, @SSN, @ClaimPaidAmount, @ClaimIntAmount, @OVAmount, @Status, @proDate, @proTime)"
        Dim cmd1 As OleDbCommand = New OleDbCommand(qryStr1, conn)
        cmd1.Parameters.Add(New OleDbParameter("@strICN", strICN))
        cmd1.Parameters.Add(New OleDbParameter("@strEngine", strEngine))
        cmd1.Parameters.Add(New OleDbParameter("@Policy", Policy))
        cmd1.Parameters.Add(New OleDbParameter("@SSN", SSN))
        cmd1.Parameters.Add(New OleDbParameter("@ClaimPaidAmount", ClaimPaidAmount))
        cmd1.Parameters.Add(New OleDbParameter("@ClaimIntAmount", ClaimIntAmount))
        cmd1.Parameters.Add(New OleDbParameter("@OVAmount", OVAmount))
        cmd1.Parameters.Add(New OleDbParameter("@Status", Status))
        cmd1.Parameters.Add(New OleDbParameter("@proDate", strdate))
        cmd1.Parameters.Add(New OleDbParameter("@proTime", strtime))

        Try

            cmd1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        conn.Close()


    End Sub


    Function ReconeChk(ByVal excelrow As String)
        Dim strRecCode As String = ""
        Dim servCode As String = ""
        Dim RemarkCode As String
        Dim RemarkCHK As String

        For i = 10 To 22 Step 2
            ApiChk()
            servCode = Trim(autECLPSObj.GetText(i, 4, 6))
            ApiChk()
            RemarkCHK = Trim(autECLPSObj.GetText(i, 36, 2))
            For index = 0 To 6
                If CodesCol(index) IsNot Nothing Then
                    If CodesCol(index) = servCode Then
                        RemarkCode = RemarkColl(index)
                        strRecCode = GetReconCode(RemarkCode)
                        ApiChk()
                        If servCode = "------" Then
                            ApiChk()
                            Exit Function
                        End If
                        If strRecCode = "" Then
                            If servCode = "NCOPAY" Or servCode = "COPAY" Or servCode = "CXINT" Or RemarkCHK = "80" Or RemarkCHK = "0H" Then
                                ApiChk()
                            Else
                                ApiChk()
                                strRecCode = InputBox("Enter Recon Code", "Recon Code", "")
                            End If

                        End If

                        If RemarkCHK = "UG" Or RemarkCHK = "W1" Or RemarkCHK = "0H" Or RemarkCHK = "YM" Or InStr(DenialCode, RemarkCHK) > 0 Then
                            ApiChk()

                        Else
                            ApiChk()
                            autECLPSObj.SendKeys(strRecCode, i, 36)

                        End If
                    End If
                End If


            Next

            If RemarkCHK = "UG" Or RemarkCHK = "W1" Then
                ApiChk()
                autECLPSObj.SendKeys("G" & strRecCode, 2, 2)
                Enter()
            Else
                ApiChk()
                rcFlot(excelrow)
            End If
        Next


    End Function

    Sub clearChk()
        ApiChk()
        autECLPSObj.SendKeys("[Clear]")
    End Sub

    Sub Get_Edit()
        Dim editNum As String
        Dim MpcLine() As String

        editNum = autECLPSObj.gettext(24, 1, 80)
        MpcLine = editNum.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)

        For i As Integer = 0 To MpcLine.Length - 1
            Dim editChk As String = ""
            editChk = MpcLine(i)
            If InStr(editChk, "WE") Then
                voidEdit = editChk
                Exit For
            End If

        Next

    End Sub

    Sub fundingChk()

        Dim Mxi As String
        Dim Mmi As String
        ApiChk()
        clearChk()
        ApiChk()
        adjline = CltrLine1() & ","
        autECLPSObj.SendKeys("[erase eof]", 2, 2)
        autECLPSObj.SendKeys(adjline, 2, 2)
        autECLPSObj.SendKeys("MRI", 2, 2)
        Enter()

        ApiChk()
        Policy = autECLPSObj.GetText(7, 8, 6)
        ApiChk()
        Plan = autECLPSObj.GetText(7, 15, 4)

        ApiChk()
        Mxi = "MXI," & Policy & "," & Plan & ",,"
        autECLPSObj.SendKeys("[erase eof]", 2, 2)
        autECLPSObj.SendKeys(Mxi, 2, 2)
        Enter()

        ApiChk()
        Policy1 = autECLPSObj.GetText(6, 2, 6)
        ApiChk()
        Plan1 = autECLPSObj.GetText(6, 10, 4)

        ApiChk()
        Mmi = "MMI," & Policy1 & "," & Plan1 & ",,"
        autECLPSObj.SendKeys("[erase eof]", 2, 2)
        autECLPSObj.SendKeys(Mmi, 2, 2)
        Enter()

        ApiChk()
        Funding = autECLPSObj.GetText(5, 13, 1)
        ApiChk()
        SState = autECLPSObj.GetText(11, 51, 2)
        ApiChk()
        Erisa = autECLPSObj.GetText(6, 53, 1)

    End Sub


End Module
