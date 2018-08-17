Module MdlOrsCreation

    Sub OrsCreation()
        Dim orssts As String
        Dim ORSN As String

        ApiChk()

        Draft = autECLPSObj.GetText(8, 59, 10)

        ApiChk()
        autECLPSObj.SendKeys("[pf9]")
        ApiChk()

        Do
            ApiChk()
        Loop Until autECLPSObj.GetText(2, 25, 6) = "ONLINE"

        ApiChk()

        ApiChk()
        autECLPSObj.SendKeys("34", 22, 5)
        ApiChk()
        Enter()

        ApiChk()
        orssts = autECLPSObj.GetText(6, 6, 4)

        If orssts = "PNSP" Then

            ORSN = MsgBox("Do You Want Create new ORS?", vbYesNo + vbDefaultButton1, "Claim ORS")
            If ORSN = vbYes Then
                ApiChk()
                autECLPSObj.SendKeys("NI", 22, 5)
                ApiChk()
                Enter()

            ElseIf ORSN = vbNo Then

                ApiChk()
                autECLPSObj.SendKeys("[pf1]")
                Exit Sub


            End If

        Else
            ApiChk()
            autECLPSObj.SendKeys("NI", 22, 5)
            ApiChk()
            Enter()

        End If

        ApiChk()
        If autECLPSObj.GetText(23, 2, 2) = "FI" Then
            autECLPSObj.SendKeys("[pf2]")
            GoTo ComeHere
        End If
ComeHere:
        ApiChk()
        autECLPSObj.SendKeys("CRT", 3, 14)
        autECLPSObj.SendKeys("GD", 7, 10)
        autECLPSObj.SendKeys("N", 7, 68)
        autECLPSObj.SendKeys("C", 17, 64)
        autECLPSObj.SendKeys("PNSP", 20, 6)
        autECLPSObj.SendKeys("ADJ/CLM BASED ON CRT PROJECT ", 18, 9)
        autECLPSObj.SendKeys("AD", 22, 5)
        Enter()
        ApiChk()
        ApiChk()
        autECLPSObj.SendKeys(MainForm.MetroTextBoxProjectNumber.Text)
        autECLPSObj.SendKeys("[tab]")
        autECLPSObj.SendKeys(MainForm.MetroTextBoxProjectInstruction.Text)
        Enter()
        autECLPSObj.SendKeys(MainForm.MetroTextBoxProjectExtraInstruction.Text)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys("[tab]")
        ApiChk()
        autECLPSObj.SendKeys(" Draft No. " & Draft)
        Enter()
        ApiChk()
        autECLPSObj.SendKeys("The Claim has Been Adjusted under the Same Icn: ")
        ApiChk()
        autECLPSObj.SendKeys("[pf7]")

        ApiChk()
        autECLPSObj.SendKeys(" ", 11, 73)

        ApiChk()
        autECLPSObj.SendKeys("ff", 22, 5)
        Enter()

        ApiChk()
        ApiChk()
        autECLPSObj.SendKeys("[pf1]")

    End Sub

End Module
