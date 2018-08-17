Imports System.Threading
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Text

Module MdlScrapHardCopyData

    Public FlagExit As Boolean = False

    Public strExitReason As String

    Public Function ClearText(ByVal Text As String) As String

        Return Text.Trim().Replace(vbCr, "").Replace(vbLf, "").Replace("=", "").Replace("  ", "")

    End Function

    Dim strId, pName, iName, billingProvider, billingNpi, physName, billingProviderAddress, billingTin, policy As String
    Dim priorNo, cliaNo, typeSv, diagCode, anesTime, patAccount, provAccptAssgn, totCharge, totPat As String
    Dim servPhyName, servPhyNpi, servPhyPhone As String

    '// Get step action for HCFA // 
    Sub GetHcfa(ByRef cntrlUsrImport As FrmCRT)

        Dim i As Integer
        Dim intLnCount As Integer
        Dim dxptr, dxptr1
        intLnCount = 0


        Dim LineCount As Integer = 0
        Dim TempLineCount As Integer = 0

        Dim TempLine As String = ""
        Dim TempLine2 As String = ""

        Dim newArr() As String



        Dim arrEdss(2000) As String

        For Each strLine In cntrlUsrImport.rtfEdss.Lines


            arrEdss(intLnCount) = strLine
            intLnCount = intLnCount + 1

        Next


        Dim lines() As String = cntrlUsrImport.rtfEdss.Lines

        ' dx code
        'Dim DXCode As String = ManualDXCodeFromHC(lines)

        For Each strLine As String In lines

            If InStr(strLine, "VENDOR ID") Then
                claimType = Mid(strLine, (InStr(strLine, "VENDOR ID")) + 20, 20).Trim

                If InStr(claimType, "HOSPITAL BILLING") Then
                    claimType = "UB"
                Else
                    claimType = "HCFA"
                End If
            End If

            If claimType = "HCFA" Then

                If InStr(strLine, "12A") > 0 Then

                    correctedInfo = Mid(lines(LineCount + 2), (InStr(strLine, "12A")), 20).Trim

                End If

                If InStr(strLine, "7B") > 0 Then

                    policy = Mid(lines(LineCount + 2), (InStr(strLine, "7B")), 40).Trim
                    FrmCRT.PolicyHTextBox.Text = policy

                End If

                ' patient name and insured name
                If InStr(strLine, "PATIENTS NAME (LFM)") > 0 Then

                    pName = ClearText(Mid(lines(LineCount + 2), LineCount, 60).Trim)
                    FrmCRT.NameHTextBox.Text = pName

                    TempLine = lines(LineCount + 1)
                    newArr = TempLine.Split("|")
                    iName = Left(newArr(2), 60).Trim

                    TempLine = lines(LineCount + 2)
                    newArr = TempLine.Split("|")
                    iName = iName & " " & Left(newArr(2), 50).Trim

                End If


                ' billing provider 
                If InStr(strLine, "11 BILLING PROVIDER NAME") > 0 Then

                    billingProvider = Mid(lines(LineCount + 1), InStr(strLine, "11 BILLING PROVIDER NAME"), 60).Trim

                    billingProviderAddress = Mid(lines(LineCount + 3), InStr(strLine, "11 BILLING PROVIDER NAME"), 60).Trim

                    billingProviderAddress = billingProviderAddress & "," & Mid(lines(LineCount + 4), InStr(strLine, "11 BILLING PROVIDER NAME"), 60).Trim

                    billingProviderAddress = billingProviderAddress & "," & Mid(lines(LineCount + 5), InStr(strLine, "11 BILLING PROVIDER NAME"), 60).Trim

                    billingNpi = GetBillingNPI(LineCount, lines)

                    billingTin = GetBillingTIN(LineCount, lines)

                End If

                ' refering phys

                If InStr(strLine, "17 REFERRING PHYS") > 0 Then

                    TempLine = lines(LineCount + 1)
                    newArr = TempLine.Split("|")
                    physName = Left(newArr(1), 60).Trim

                    TempLine = lines(LineCount + 2)
                    newArr = TempLine.Split("|")
                    physName = physName & " " & Left(newArr(1), 30).Trim

                End If

                ' prior no
                If InStr(strLine, "23 PRIOR AUTHORIZATION NUMBER") > 0 Then

                    priorNo = Mid(lines(LineCount + 1), (InStr(strLine, "23 PRIOR AUTHORIZATION NUMBER")), 50).Trim

                End If



                ' date of service
                If InStr(strLine, "DATES OF SERVICE") > 0 Then

                    Dim checkLine1 As Integer = 3



                    TempLine = lines(LineCount + 3).Trim
                    newArr = TempLine.Split("|")

                    dos_from = newArr(1).Split("  ")(1)
                    dos_to = newArr(1).Split("  ")(3)


                    plcSv = newArr(2)

                    typeSv = newArr(3)

                    cpt = newArr(4).Split(" ")(0)


                    modCode = newArr(4).Split(" ")(1)

                    diagCode = newArr(5)

                    charges = newArr(6)
                    'FrmCRT.BChargeHRTextBox.Text = FrmCRT.BChargeHRTextBox.Text & vbCrLf & charges

                    dayUnit = newArr(7)

                    anesTime = newArr(8)





                End If

                ' patients account
                If InStr(strLine, "26 PATIENTS ACCOUNT#") > 0 Then

                    patAccount = Mid(lines(LineCount + 2), (InStr(strLine, "26 PATIENTS ACCOUNT#")), 20).Trim

                End If

                ' accpts assgnmnt
                If InStr(strLine, "27 PROV ACCPTS ASGNMT") > 0 Then

                    provAccptAssgn = Mid(lines(LineCount + 2), (InStr(strLine, "27 PROV ACCPTS ASGNMT")), 20).Trim

                End If

                ' total charge
                If InStr(strLine, "28 TOT CHARGE") > 0 Then

                    totCharge = Mid(lines(LineCount + 2), (InStr(strLine, "28 TOT CHARGE")), 13).Trim

                End If

                ' pad pd
                If InStr(strLine, "29 TOT PAT PD") > 0 Then

                    totPat = Mid(lines(LineCount + 2), (InStr(strLine, "29 TOT PAT PD")), 13).Trim

                End If

                If InStr(strLine, "33 SERVICING PHYSICIAN/SUPPLIER NAME") > 0 Then

                    servPhyName = Mid(lines(LineCount + 2), InStr(strLine, "33 SERVICING PHYSICIAN/SUPPLIER NAME"), 105).Trim

                    servPhyNpi = GetBillingNPI(LineCount, lines)
                    servPhyPhone = GetPhoneNo(LineCount, lines)

                End If

                MainForm.MetroGridHardcopyInfo.Rows.Add(New String() {plcSv, " ", cpt, dos_from, dos_to, dayUnit, charges})

            Else        'for UB claim



                If InStr(strLine, "PROVIDER NM") > 0 Then

                    billingProvider = Mid(strLine, (InStr(strLine, "PROVIDER NM")) + 11, 60).Trim

                    billingProviderAddress = Mid(lines(LineCount + 1), InStr(strLine, "PROVIDER NM") + 15, 60).Trim

                    billingProviderAddress = billingProviderAddress & "," & Mid(lines(LineCount + 2), InStr(strLine, "PROVIDER NM") + 15, 20).Trim

                    billingNpi = Mid(lines(LineCount + 3), InStr(strLine, "PROVIDER NM") + 57, 20).Trim


                End If

                If InStr(strLine, "FEDERAL TAX ID") > 0 Then

                    billingTin = Mid(strLine, (InStr(strLine, "FEDERAL TAX ID")) + 15, 20).Trim

                End If

                If InStr(strLine, "1PAYER NAME") > 0 Then

                    policy = Mid(lines(LineCount + 5), InStr(strLine, "1PAYER NAME") + 50, 7).Trim

                    FrmCRT.PolicyHTextBox.Text = policy

                End If

                If InStr(strLine, "PATIENTS NAME") > 0 Then

                    pName = Mid(lines(LineCount + 1), InStr(strLine, "PATIENTS NAME"), 60).Trim & " " & Mid(lines(LineCount + 2), InStr(strLine, "PATIENTS NAME"), 60).Trim

                    FrmCRT.NameHTextBox.Text = pName

                End If

                If InStr(strLine, "TYPE BILL: HOSPITAL/OUTPATIENT/ADM-DSCH CL") > 0 Then

                    correctedInfo = Mid(strLine, (InStr(strLine, "TYPE BILL: HOSPITAL/OUTPATIENT/ADM-DSCH CL")) + 49, 5).Trim


                End If

                ' date of service
                If InStr(strLine, "LINE REV   RATE") > 0 Then

                    Dim checkLine As Integer = 0

                    While InStr(Mid(lines(LineCount + checkLine), InStr(strLine, "MOD"), 3).Trim, "MOD")

                        dos_from = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 10, 5).Trim

                        rev_code = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 6, 4).Trim

                        dayUnit = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 54, 7).Trim

                        charges = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 64, 10).Trim

                        cpt = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 75, 6).Trim

                        modCode = Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 91, 4).Trim & "," & Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 99, 4).Trim & "," & Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 107, 4).Trim & "," & Mid(lines(LineCount + 2), InStr(strLine, "LINE REV   RATE") + 115, 4).Trim

                        'FrmCRT.DosRTextBox.Text = FrmCRT.DosRTextBox.Text & vbCrLf & dos_from

                        'FrmCRT.CodesHRTextBox.Text = FrmCRT.CodesHRTextBox.Text & vbCrLf & rev_code & "  " & cpt

                        'FrmCRT.BChargeHRTextBox.Text = FrmCRT.BChargeHRTextBox.Text & vbCrLf & charges

                        checkLine += 1
                    End While

                End If
                MainForm.MetroGridHardcopyInfo.Rows.Add(New String() {" ", rev_code, cpt, dos_from, dos_from, dayUnit, charges})
            End If

            LineCount += 1

        Next

    End Sub

    Function GetBillingNPI(ByVal LinNo As Integer, ByVal Lines As Array) As String

        Dim billingNpi As String



        Do

            If InStr(Lines(LinNo), "-----") > 0 Then
                Exit Do
            End If

            If InStr(Lines(LinNo), "NPI") > 0 Then

                billingNpi = Mid(Lines(LinNo), (InStr(Lines(LinNo), "NPI")) + 5, 12).Trim

                'MsgBox(billingNpi)

            End If

            LinNo = LinNo + 1

        Loop

        Return billingNpi

    End Function

    Function GetBillingTIN(ByVal LinNo As Integer, ByVal Lines As Array) As String

        Dim billingTIN As String



        Do

            If InStr(Lines(LinNo), "-----") > 0 Then
                Exit Do
            End If

            If InStr(Lines(LinNo), "TIN#") > 0 Then

                billingTIN = Mid(Lines(LinNo), (InStr(Lines(LinNo), "TIN#")) + 5, 15).Trim

            End If

            LinNo = LinNo + 1

        Loop

        Return billingTIN

    End Function

    Function GetPhoneNo(ByVal LinNo As Integer, ByVal Lines As Array) As String

        Dim phoneNo As String



        Do

            If InStr(Lines(LinNo), "-----") > 0 Then
                Exit Do
            End If

            If InStr(Lines(LinNo), "PHONE#") > 0 Then

                phoneNo = Mid(Lines(LinNo), (InStr(Lines(LinNo), "PHONE#")) + 7, 12).Trim

                'MsgBox(billingNpi)

            End If

            LinNo = LinNo + 1

        Loop

        Return phoneNo

    End Function


    Function ManualDXCodeFromHC(ByVal lines As Array) As String

        Dim i As Integer = 0
        Dim dxCodes As String = ""
        Dim intLnCount As Integer = 0
        'Dim lines(2000) As String

        For Each strLine In lines

            intLnCount = intLnCount + 1

        Next


        For i = 0 To intLnCount Step 1

            If InStr(lines(i), "21 DIAGNOSIS Or NATURE") > 0 And InStr(lines(i + 2), "1|A") = 0 Then

                dxCodes = dxCodes & Trim(Mid(lines(i + 2), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 2), 42, 10))


                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 4), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 4), 42, 10))



                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 6), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 6), 42, 10))



                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 7), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 7), 42, 10))



                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 8), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 8), 42, 10))



                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 9), 6, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 9), 42, 10))

            ElseIf InStr(lines(i), "21 DIAGNOSIS Or NATURE") > 0 And InStr(lines(i + 2), "1|A") > 0 Then

                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 2), 7, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 2), 32, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 2), 56, 10))

                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 4), 7, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 4), 32, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 4), 56, 10))


                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 6), 7, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 6), 32, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 6), 56, 10))

                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 7), 7, 10))
                dxCodes = dxCodes & ", " & Trim(Mid(lines(i + 7), 32, 10))
                dxCodes = dxCodes & "," & Trim(Mid(lines(i + 7), 56, 10))

            End If

            i = i + 1


        Next

        Return dxCodes

    End Function

    Function StartCollectValueForBillingProvider(ByRef cntrlUsrImport As FrmCRT, ByVal RowCount As Integer, ByVal LineCount As Integer, ByVal FieldCount As Integer, ByVal strTempLine As String) As String

        '   nMsgBox("check gvSvc")
        'cntrlUsrImport.rtfEdss.Text = strEdss




        Dim strResult As New StringBuilder

        Dim arrTemp(), arrTempField() As String
        Dim strTempField As String
        Dim count As Integer = 0

        Do

            strTempLine = cntrlUsrImport.rtfEdss.Lines(LineCount)

            If InStr(strTempLine, "-----") > 0 Then
                Exit Do
            End If
            'arrTemp(count) = strTempLine
            LineCount += 1
            strResult.AppendLine(strTempLine)

        Loop



        Return strResult.ToString


    End Function
    '//Collect value from Hardcopy //


End Module
