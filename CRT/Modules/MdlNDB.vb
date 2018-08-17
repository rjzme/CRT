Module MdlNDB

    Sub NDBFee()

        MdlMain.ChkPPI()

        ApiChk()

        autECLPSObj.SendKeys("[PF11]")
        ApiChk()
        ApiChk()
        autECLPSObj.SendKeys("s", 21, 22)
        autECLPSObj.SendKeys("[Enter]")
        ApiChk()
        autECLPSObj.SendKeys("i", 22, 12)
        autECLPSObj.SendKeys("c", 22, 24)
        autECLPSObj.SendKeys("[Enter]")
        ApiChk()
        Dim CPT_Split() As String
        CPT_Split = cpt.Split(",")

        Dim Doll As String
        Dim Cents As String
        Dim charges1 As String
        ApiChk()
        Doll = charges.Split(".")(0)
        Cents = charges.Split(".")(1)
        charges1 = Doll & Cents
        For i = 0 To CPT_Split.Length - 1

            ApiChk()
            autECLPSObj.SendKeys("t", 2, 6)
            autECLPSObj.SendKeys(strMarket, 2, 12)
            autECLPSObj.SendKeys(FEEschd, 2, 24)
            autECLPSObj.SendKeys("Y", 2, 48)
            autECLPSObj.SendKeys("i", 2, 63)
            autECLPSObj.SendKeys(Left(dos_from, 2), 9, 4)
            autECLPSObj.SendKeys(Mid(dos_from, 3, 2), 9, 7)
            autECLPSObj.SendKeys("20" & Right(dos_from, 2), 9, 10)
            autECLPSObj.SendKeys(Trim(plcSv), 9, 15)
            autECLPSObj.SendKeys(charges1, 9, 18)
            autECLPSObj.SendKeys(CPT_Split(i), 9, 30)
            autECLPSObj.SendKeys(dayUnit, 9, 40)
            autECLPSObj.SendKeys(modCode, 9, 53)
            autECLPSObj.SendKeys("[PF6]")
            ApiChk()
            ClaimAllowable = autECLPSObj.GetText(19, 68, 10)
            ApiChk()
            FrmCRT.OCIRTextBox.Text = FrmCRT.OCIRTextBox.Text & vbCrLf & GetUnetScreen()

        Next
        ApiChk()
        autECLPSObj.SendKeys("[PF2]")
        ApiChk()
        autECLPSObj.SendKeys("[PF2]")
        ApiChk()
        autECLPSObj.SendKeys("[clear]")

    End Sub


End Module
