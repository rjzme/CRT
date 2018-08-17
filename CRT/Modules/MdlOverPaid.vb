Module MdlOverPaid

    Function main(ByVal excelrow As String)



        ApiChk()

        ApiChk()
        autECLPSObj.SendKeys("[pf12]")
        ApiChk()
        clearChk()
        ApiChk()
        autECLPSObj.SendKeys(CltrLine1(), 1, 2)
        Enter()

        ApiChk()
        autECLPSObj.SendKeys("69", 3, 78)
        autECLPSObj.SendKeys("GZH", 2, 2)
        Enter()
        Enter()
        Enter()
        ApiChk()
        Dim svcPos As String = PoS
        Dim svcCode As String = SvC
        Dim svcFD As String = FDate
        Dim svcLD As String = LDate

        'For i = 10 To 22 Step 2
        '    ApiChk()
        '    svcCode = autECLPSObj.GetText(i, 4, 6)

        '    If svcCode <> "------" Then
        '        ApiChk()
        '        svcCode = autECLPSObj.GetText(i, 4, 6)
        '        svcPos = autECLPSObj.GetText(i, 1, 2)
        '        svcFD = autECLPSObj.GetText(i, 11, 6)
        '        svcLD = autECLPSObj.GetText(i, 18, 6)
        '        Exit For
        '    End If


        'Next


        ApiChk()


        Try
            If CurrentPayment.ToString <> "" And CurrentPayment.ToString <> "0" And CurrentPayment.ToString <> "0.00" And CurrentPayment <> payment Then
                Dim Ovp As Decimal
                Ovp = (payment - CurrentPayment)
                Overpaid = Convert.ToString(Ovp)

                Dim OvPaidArr() As String = Overpaid.Split(".")
                OverPaidDollars = OvPaidArr(0)
                OverPaidDollars = OverPaidDollars.PadLeft(5, "0")
                If OvPaidArr(1) = "" Then
                    OvPaidArr(1) = "00"
                End If
                OverPaidCents = OvPaidArr(1)

                claimPaid = Convert.ToString(CurrentPayment)

                Dim CuPaidArr() As String = claimPaid.Split(".")
                PdDol = CuPaidArr(0)
                PdDol = PdDol.PadLeft(5, "0")
                If CuPaidArr(1) = "" Then
                    CuPaidArr(1) = "00"
                End If
                PdCents = CuPaidArr(1)



            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 10, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 12, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 14, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 16, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 18, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 20, 1)
        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 22, 1)
        Enter()


        If CurrentPayment.ToString = "0.00" Or CurrentPayment.ToString = "0" Then
            If Flag = "1" Then

                ApiChk()
                autECLPSObj.SendKeys(svcPos, 18, 1)
                autECLPSObj.SendKeys(svcCode, 18, 4)
                autECLPSObj.SendKeys(svcFD, 18, 11)
                autECLPSObj.SendKeys(svcLD, 18, 18)

                autECLPSObj.SendKeys(PaidDoll, 18, 41)
                autECLPSObj.SendKeys(PaidCents, 18, 47)
                autECLPSObj.SendKeys(PaidDoll, 19, 30)
                autECLPSObj.SendKeys(PaidCents, 19, 36)
                autECLPSObj.SendKeys(PaidDoll, 19, 61)
                autECLPSObj.SendKeys(PaidCents, 19, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 18, 25)
                autECLPSObj.SendKeys("1321", 18, 29)
                autECLPSObj.SendKeys("100", 19, 55)
                Enter()

            Else

                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(svcPos, 18, 1)
                autECLPSObj.SendKeys(svcCode, 18, 4)
                autECLPSObj.SendKeys(svcFD, 18, 11)
                autECLPSObj.SendKeys(svcLD, 18, 18)
                ApiChk()
                autECLPSObj.SendKeys(PaidDoll, 18, 41)
                autECLPSObj.SendKeys(PaidCents, 18, 47)
                autECLPSObj.SendKeys(PaidDoll, 19, 30)
                autECLPSObj.SendKeys(PaidCents, 19, 36)
                autECLPSObj.SendKeys(PaidDoll, 19, 61)
                autECLPSObj.SendKeys(PaidCents, 19, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 18, 25)
                autECLPSObj.SendKeys("2021", 18, 29)
                autECLPSObj.SendKeys("100", 19, 55)
                Enter()
                Dim newSuff As String = Trim(MainForm.MetroTextBoxSuffix.Text)


                If newSuff = "" Or newSuff = ClaimSuff Then

                    ApiChk()
                    autECLPSObj.SendKeys("063", 10, 25)
                    Enter()

                Else
                    ApiChk()
                    autECLPSObj.SendKeys("038", 10, 25)
                    Enter()

                End If

            End If

        Else

            If Flag = "1" Then
                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(svcPos, 22, 1)
                autECLPSObj.SendKeys(svcCode, 20, 4)
                autECLPSObj.SendKeys(svcFD, 20, 11)
                autECLPSObj.SendKeys(svcLD, 20, 18)
                ApiChk()
                autECLPSObj.SendKeys(PdDol, 20, 41)
                autECLPSObj.SendKeys(PdCents, 20, 47)
                autECLPSObj.SendKeys(PdDol, 21, 30)
                autECLPSObj.SendKeys(PdCents, 21, 36)
                autECLPSObj.SendKeys(PdDol, 21, 61)
                autECLPSObj.SendKeys(PdCents, 21, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 20, 25)
                autECLPSObj.SendKeys("1321", 20, 29)
                autECLPSObj.SendKeys("100", 21, 55)
                Enter()

                ApiChk()
                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(svcPos, 18, 1)
                autECLPSObj.SendKeys(svcCode, 18, 4)
                autECLPSObj.SendKeys(svcFD, 18, 11)
                autECLPSObj.SendKeys(svcLD, 18, 18)
                ApiChk()
                autECLPSObj.SendKeys(OverPaidDollars, 18, 41)
                autECLPSObj.SendKeys(OverPaidCents, 18, 47)
                autECLPSObj.SendKeys(OverPaidDollars, 19, 30)
                autECLPSObj.SendKeys(OverPaidCents, 19, 36)
                autECLPSObj.SendKeys(OverPaidDollars, 19, 61)
                autECLPSObj.SendKeys(OverPaidCents, 19, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 18, 25)
                autECLPSObj.SendKeys("1321", 18, 29)
                autECLPSObj.SendKeys("100", 19, 55)
                Enter()
                ApiChk()
            Else
                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(svcPos, 20, 1)
                autECLPSObj.SendKeys(svcCode, 20, 4)
                autECLPSObj.SendKeys(svcFD, 20, 11)
                autECLPSObj.SendKeys(svcLD, 20, 18)
                ApiChk()
                autECLPSObj.SendKeys(PdDol, 20, 41)
                autECLPSObj.SendKeys(PdCents, 20, 47)
                autECLPSObj.SendKeys(PdDol, 21, 30)
                autECLPSObj.SendKeys(PdCents, 21, 36)
                autECLPSObj.SendKeys(PdDol, 21, 61)
                autECLPSObj.SendKeys(PdCents, 21, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 20, 25)
                autECLPSObj.SendKeys("1321", 20, 29)
                autECLPSObj.SendKeys("100", 21, 55)
                Enter()
                ApiChk()
                ApiChk()
                ApiChk()
                autECLPSObj.SendKeys(svcPos, 18, 1)
                autECLPSObj.SendKeys(svcCode, 18, 4)
                autECLPSObj.SendKeys(svcFD, 18, 11)
                autECLPSObj.SendKeys(svcLD, 18, 18)
                ApiChk()
                autECLPSObj.SendKeys(OverPaidDollars, 18, 41)
                autECLPSObj.SendKeys(OverPaidCents, 18, 47)
                autECLPSObj.SendKeys(OverPaidDollars, 19, 30)
                autECLPSObj.SendKeys(OverPaidCents, 19, 36)
                autECLPSObj.SendKeys(OverPaidDollars, 19, 61)
                autECLPSObj.SendKeys(OverPaidCents, 19, 67)
                ApiChk()
                autECLPSObj.SendKeys("002", 18, 25)
                autECLPSObj.SendKeys("2021", 18, 29)
                autECLPSObj.SendKeys("100", 19, 55)
                Enter()
                ApiChk()
                Dim newSuff As String = Trim(MainForm.MetroTextBoxSuffix.Text)


                If newSuff = "" Or newSuff = ClaimSuff Then

                    ApiChk()
                    autECLPSObj.SendKeys("063", 10, 25)
                    Enter()

                Else
                    ApiChk()
                    autECLPSObj.SendKeys("038", 10, 25)
                    Enter()

                End If



            End If


        End If




        MsgBox("Check Screen")

        ApiChk()
        autECLPSObj.SendKeys("MPP", 2, 2)
        Enter()



        ApiChk()
        CVoid = autECLPSObj.GetText(24, 3, 1)
        If CVoid = "E" Then
            ApiChk()
            Enter()
            ApiChk()
            Get_Edit()
            Status = "Void Not Completed Due To Edit: " & voidEdit
            getDateTime(excelrow)
            Return True

        End If


        ApiChk()
        autECLPSObj.SendKeys("[pf12]")

        ApiChk()
        autECLPSObj.SendKeys("EDS", 2, 2)

        Enter()
        ApiChk()
        ApiChk()
        autECLPSObj.SendKeys("N", 12, 71)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys(Rsc, 22, 49)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys("A", 1, 75)
        Enter()

        Dim copay As String
        Dim copayNBR As String
        Dim copayN As String
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
        Next

        If CurrentPayment.ToString = "0.00" Or CurrentPayment.ToString = "0" Then

            ApiChk()
            ReconeChk(excelrow)


        Else
            For i = 10 To 22 Step 2
                ApiChk()
                Dim svcCode1 As String = autECLPSObj.GetText(i, 4, 6)
                If svcCode1 = "------" Then
                    ApiChk()
                    ApiChk()
                    ApiChk()
                    autECLPSObj.SendKeys(svcPos, i, 1)
                    autECLPSObj.SendKeys(svcCode, i, 4)
                    autECLPSObj.SendKeys(svcFD, i, 11)
                    autECLPSObj.SendKeys(svcLD, i, 18)
                    ApiChk()
                    autECLPSObj.SendKeys(PdDol, i, 41)
                    autECLPSObj.SendKeys(PdCents, i, 47)
                    autECLPSObj.SendKeys(PdDol, i + 1, 30)
                    autECLPSObj.SendKeys(PdCents, i + 1, 36)
                    autECLPSObj.SendKeys(PdDol, i + 1, 61)
                    autECLPSObj.SendKeys(PdCents, i + 1, 67)
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

            ApiChk()


        End If


        If claimType = "UB" Then
            ApiChk()
            autECLPSObj.SendKeys("GD2", 2, 2)
            ApiChk()
            autECLPSObj.SendKeys("[enter]")

        Else
            ApiChk()
            autECLPSObj.SendKeys("GD1", 2, 2)
            ApiChk()
            autECLPSObj.SendKeys("[enter]")
        End If
        ApiChk()
        ReconeChk(excelrow)

        OVAmount = OverPaidDollars & "." & OverPaidCents

        If OVAmount = "." Then
            OVAmount = "0.00"
        End If
    End Function




End Module
