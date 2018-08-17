Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration.ConfigurationManager

Module mdlData

    Function GetReconCode(ByVal RemarkCode As String) As String
        Try

            Dim strResult As String = ""
            Dim strCmd As String
            Dim dr As String
            'strCmd = "SELECT ReconCode.ReconCode FROM ReconCode WHERE (((ReconCode.RemarkCode)=[?]));"
            strCmd = "SELECT * FROM ReconCode WHERE RemarkCode = '" & RemarkCode & "'"
            Dim strConn As String

            strConn = My.Settings.connectionString


            conn = New OleDbConnection(strConn)
            conn.Open()

            Dim cmd As New OleDbCommand(strCmd, conn)

            'strResult = cmd.ExecuteReader

            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            While reader.Read()

                strResult = reader("ReconCode")

            End While

            conn.Close()
            Return strResult


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


End Module
