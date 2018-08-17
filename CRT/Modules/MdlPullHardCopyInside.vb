Imports System.Configuration.ConfigurationManager

Module MdlPullHardCopyInside

#Region "All Varriables"

    Dim strEncryptedOutput As String = ""
    Dim strkeyValues As String

    Dim strEncryptedText As String = ""

    Dim EDSS_Doc As String
    Dim EDSS_Error As Boolean = False
    Dim oShell
    Dim WShell, objShell, objShellWindows, objEDSS As Object
    Public EDSS_URLFound As Boolean = False
    Public OnlyHardCopyOpen As Boolean = False
#End Region


    Public Sub WaitForInput(ByVal seconds)

        Dim startTime, endTime As Date
        startTime = DateAndTime.Now
        endTime = startTime.AddSeconds(seconds)

        'MsgBox(startTime)

        'MsgBox(endTime)

        While DateAndTime.Now < endTime
            'MsgBox(DateAndTime.Now)
            Application.DoEvents()
        End While

    End Sub

    '**********************************

    '// Encrypt spcifically formated text //
    Function Encrypt_Text(ByVal eText As String) As String
        '**********************************
        Dim Encrypted_Text As String = ""
        '// Parameters to pass to jar
        Dim Jar_Parameters = eText

        Dim javapath = AppSettings.Get("jarFilePath")
        'EDSS_Local_Folder & Jar_File_Name
        Dim javaclass = "com.uhg.edss.om.util.ISETURLEncryptor "

        '// Set command line for executing jar file
        Dim cmdline = "java -cp " & javapath & " " & javaclass & Jar_Parameters


        '// Run jar file, and get output encryption/decryption
        With CreateObject("WScript.Shell").Exec(cmdline)
            'System.Threading.Thread.Sleep(100)
            While Not .StdOut.AtEndOfStream
                Encrypted_Text = Encrypted_Text & .StdOut.ReadAll
            End While
        End With

        ' blnEncrypt_Error = False

        On Error Resume Next
        strEncryptedOutput = Trim(Split(Encrypted_Text, "Output:")(1))    '// Get Encrypted/Decrypted value
        '// Check for Error
        If Err.Number <> 0 Then
            'blnEncrypt_Error = True
        End If
        Return strEncryptedOutput

    End Function


    Public CountCreateIEObjec As Integer

    '// Launch EDSS with the given URL //
    Sub Launch_Edss(ByVal strEDSS_URL As String)
        '**************************

        WShell = CreateObject("WScript.Shell")
        '// Open new IE Instance if EDSS is NOT open
CreateIEObject:

        If CountCreateIEObjec > 3 Then
            MsgBox("Hardcopy not opening")
            Exit Sub
        End If

        Try
            If EDSS_URLFound = False Then
                '//objEDSS = cntrlUsrImport.WebEdss1
                objEDSS = CreateObject("InternetExplorer.Application")   'InternetExplorer.Application
            Else
                '// Activate Existing EDSS Window
                WShell.AppActivate(objEDSS.document.title)
            End If

            objEDSS.Navigate(strEDSS_URL)
            objEDSS.Visible = True
        Catch ex As Exception
            CountCreateIEObjec += 1
            objEDSS.quit
            EDSS_URLFound = False
            GoTo CreateIEObject
        End Try






        '// Interaction.AppActivate("Grand_Daddy")
        WaitForInput(3)
        'System.Threading.Thread.Sleep(3000)
        '//  Interaction.AppActivate("Unet Wrapper")

        If InStr(objEDSS.Document.Body.InnerText, "Screen not valid") > 0 Then
            '//MsgBox("Screen not valid")
            MsgBox("Hardcopy not opening")
            objEDSS.quit
            Exit Sub

        End If

        If InStr(objEDSS.Document.Body.InnerText, "Retrieved 1 document") = 0 Then
            '//MsgBox("Screen not valid")
            'ExitReason = "Hardcopy not opening"
            'Exit Function

            '// MsgBox("Multiple hard copy found!!!" & Chr(13) & "Open correct hardcopy manually and click OK", vbOK, "Multiple Hard Copy!!!")

            MsgBox("Multiple Hardcopy OOS")
            objEDSS.quit
            Exit Sub

        End If

        EDSS_Doc = objEDSS.Document.GetElementById("edmsUIDocumentViewForm:documentTextContent").InnerText

        ' // OpenHardCopy()

        '//MainGD.rtfEdss.Text = EDSS_Doc
        '//   MsgBox(MainGD.rtfEdss.Text)
        FrmCRT.rtfEdss.Text = EDSS_Doc



    End Sub

    '// Check whether EDSS document are already opened. //
    Sub Determine_EDSS_Open(strEDSS_URL)
        '**************************
        WShell = CreateObject("WScript.Shell")
        EDSS_URLFound = False
        objShell = CreateObject("Shell.Application")
        objShellWindows = objShell.Windows

        For Each objEDSS In objShell.Windows
            On Error Resume Next
            If InStr(UCase(objEDSS.LocationURL), UCase(strEDSS_URL)) Then

                If InStr(UCase(objEDSS.FullName), "IEXPLORE.EXE") Then
                    EDSS_URLFound = True
                    Exit For
                End If
            End If

        Next


    End Sub

    Sub OpenHardCopy(edss_id As String, edss_password As String, flndcc As String)

        flndcc = flndcc

        Dim i As Long
        Dim IE As Object
        Dim objElement As Object
        Dim objCollection As Object

        ' Create InternetExplorer Object
        IE = CreateObject("InternetExplorer.Application")

        ' Send the form data To URL As POST binary request
        IE.Navigate("http://edsscontent.uhc.com:8080/edssui")

        ' Show IE
        IE.Visible = True

        ' Wait while IE loading...
        Do While IE.Busy Or IE.ReadyState <> 4
            Application.DoEvents()
        Loop

        objCollection = IE.Document.getElementsByTagName("input")

        i = 0
        While i < objCollection.Length
            If objCollection(i).Name = "loginForm:userID" Then

                ' Set text for search
                objCollection(i).Value = edss_id
            Else
                If objCollection(i).Name = "loginForm:password" Then
                    objCollection(i).Value = edss_password
                End If

            End If
            i = i + 1
        End While

        IE.Document.getElementById("loginForm:logonButton").Click

        ' Wait while IE re-loading...
        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementById("edmsUIForm:datagroupDropdownList_label") Is Nothing
            Application.DoEvents()
        Loop

        IE.Document.getElementById("edmsUIForm:datagroupDropdownList_label").Click

        ' Wait while IE re-loading...
        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementsByTagName("li") Is Nothing
            Application.DoEvents()
        Loop

        For Each elem In IE.Document.getElementsByTagName("li")

            If elem.GetAttribute("data-label") = "EDI Claims - FLNDCC" Then
                elem.Click
                Exit For
            End If
        Next

        ' Wait while IE re-loading...
        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementById("tabView:0:searchForm:j_id_1i:0:key") Is Nothing
            Application.DoEvents()
        Loop

        IE.Document.getElementById("tabView:0:searchForm:j_id_1i:0:key").Value = flndcc

        ' Wait while IE re-loading...
        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementById("tabView:0:searchForm:submitButton") Is Nothing
            Application.DoEvents()
        Loop

        IE.Document.getElementById("tabView:0:searchForm:submitButton").Click

        Do While IE.Busy Or IE.ReadyState <> 4
            Application.DoEvents()
        Loop

        System.Threading.Thread.Sleep(2000)

        If IE.Document.getElementById("tabView:0:resultForm:j_id_2k:0:viewDocumentEven") Is Nothing Then
            IE.Navigate("http://edsscontent.uhc.com:8080/edssui/edmsuidatagroups.jsf")
        End If

        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementById("tabView:0:searchForm:j_id_1i:0:key") Is Nothing
            Application.DoEvents()
        Loop

        IE.Document.getElementById("tabView:0:searchForm:j_id_1i:0:key").Value = flndcc
        IE.Document.getElementById("tabView:0:searchForm:submitButton").Click

        Do While IE.Busy Or IE.ReadyState <> 4
            Application.DoEvents()
        Loop

        System.Threading.Thread.Sleep(2000)

        If IE.Document.getElementById("tabView:0:resultForm:j_id_2k:0:viewDocumentEven") Is Nothing Then
            MsgBox("Unable To Open Hard Copy :" & flndcc)
            IE.Quit
            Exit Sub
        End If

        IE.Document.getElementById("tabView:0:resultForm:j_id_2k:0:viewDocumentEven").Click

        Do While IE.Busy Or IE.ReadyState <> 4 Or IE.Document.getElementById("edmsUIDocumentViewForm:documentTextContent") Is Nothing
            Application.DoEvents()
        Loop

        EDSS_Doc = IE.Document.getElementById("edmsUIDocumentViewForm:documentTextContent").InnerHtml

        FrmCRT.rtfEdss.Text = EDSS_Doc

        IE.Quit
        ' Clean up
        IE = Nothing
        objElement = Nothing
        objCollection = Nothing

    End Sub

End Module