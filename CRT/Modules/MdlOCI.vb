Module MdlOCI
    Sub OpenOCI(ByVal flndcc As String, ByVal Session As String)

        Dim IE As Object

        Dim objCollection As Object

        ' Create InternetExplorer Object
        IE = CreateObject("InternetExplorer.Application")

        ' Send the form data To URL As POST binary request
        IE.Navigate("http://uhc102.uhc.com/sites/dept7/NAT/NAT_Dev_Site/Webpage%20Library/WebForms/Oci_IriP/Oci_IriP.htm")

        ' Show IE
        IE.Visible = True

        ' Wait while IE loading...
        Do While IE.Busy Or IE.ReadyState <> 4
            Application.DoEvents()
        Loop

        objCollection = IE.Document.getElementsByTagName("input")

        i = 0
        While i < objCollection.Length
            If objCollection(i).Value = "oci" Then
                objCollection(i).Click
            End If
            i = i + 1
        End While

        IE.Document.getElementById("Emulator").Value = Session

        IE.Document.getElementById("FLNDCC").Value = flndcc

        IE.Document.getElementById("btnEDSS").Click

        IE.Document.getElementById("btnWrite").Click

        IE.Quit
        ' Clean up
        IE = Nothing
        objCollection = Nothing

    End Sub
End Module
