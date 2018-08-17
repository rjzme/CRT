Module MdlUnderPaid

    Function Main(ByVal excelrow As String)

        ApiChk()
        autECLPSObj.SendKeys("[Clear]")
        ApiChk()
        autECLPSObj.SendKeys(CltrLine1(), 1, 2)
        Enter()
        ApiChk()

        ApiChk()
        autECLPSObj.SendKeys("69", 3, 78)
        autECLPSObj.SendKeys("GZH", 2, 2)
        Enter()
        Enter()
        Enter()
        ApiChk()


        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 10, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 12, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 14, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 16, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 18, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 20, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 22, 1)
        Enter()

        ApiChk()
        autECLPSObj.SendKeys(SlLine(), 20, 1)
        autECLPSObj.SendKeys(PaidDoll, 20, 41)
        autECLPSObj.SendKeys(PaidCents, 20, 47)
        autECLPSObj.SendKeys(PaidDoll, 21, 30)
        autECLPSObj.SendKeys(PaidCents, 21, 36)
        autECLPSObj.SendKeys(PaidDoll, 21, 61)
        autECLPSObj.SendKeys(PaidCents, 21, 67)

        autECLPSObj.SendKeys("1321", 20, 29)
        autECLPSObj.SendKeys("100", 21, 55)
        Enter()
        Enter()
        ApiChk()
        Dim lineRemark As String
        For x = 10 To 24 Step 2
            ApiChk()
            lineRemark = Trim(autECLPSObj.GetText(x, 29, 2))
            If lineRemark = "P" Then
                autECLPSObj.SendKeys("13", x, 29)
                Enter()
                Status = autECLPSObj.GetText(24, 3, 25)
                getDateTime(excelrow)
                Return True
            End If
        Next

        ApiChk()
        'MsgBox("Check Screen")
        autECLPSObj.SendKeys("MPP", 2, 2)
        Enter()

        ApiChk()
        CVoid = autECLPSObj.gettext(24, 3, 1)
        ApiChk()
        If CVoid = "E" Then
            ApiChk()
            Enter()
            ApiChk()
            Get_Edit()
            Status = "Void Not Completed Due To Edit: " & voidEdit
            getDateTime(excelrow)
            Return True

        End If


    End Function

End Module
