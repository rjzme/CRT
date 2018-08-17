Module MdlZeroPaid
    Function Main(ByVal excelrow As String)

        ApiChk()
        autECLPSObj.SendKeys("[Clear]")
        ApiChk()
        autECLPSObj.SendKeys(CltrLine1(), 1, 2)
        autECLPSObj.SendKeys("[enter]")

        ApiChk()
        autECLPSObj.SendKeys("69", 3, 78)
        autECLPSObj.SendKeys("GZH", 2, 2)
        autECLPSObj.SendKeys("[enter]")
        autECLPSObj.SendKeys("[enter]")
        autECLPSObj.SendKeys("[enter]")
        ApiChk()

        ApiChk()
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 10, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 12, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 14, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 16, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 18, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 20, 1)
        autECLPSObj.SendKeys("-------------------------------------------------------------------------------------------------------------", 22, 1)
        autECLPSObj.SendKeys("[enter]")

        ApiChk()
        autECLPSObj.SendKeys(SlLine(), 22, 1)
        autECLPSObj.SendKeys("RXRX----", 22, 1)
        autECLPSObj.SendKeys("0000001", 22, 41)
        autECLPSObj.SendKeys("0000001", 22, 50)

        ApiChk()
        autECLPSObj.SendKeys("1321", 22, 29)
        autECLPSObj.SendKeys("100", 23, 55)
        autECLPSObj.SendKeys("[enter]")
        autECLPSObj.SendKeys("[enter]")
        ApiChk()


        ApiChk()
        'MsgBox("Check Screen")
        ApiChk()
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


        Return False

    End Function
End Module
