Imports LAU.NormingExtension
Imports System.IO

Module Startup
    Public AppName As String = "Leave Account Auto Update"
    Public NormingDB As DataConnector

    Dim tbl As DataTable

    Sub ReadConfigFile()
        Try
            Dim reader As New IO.StreamReader(".\config.conf")

            Dim nmgServer As String = reader.ReadLine()
            Dim nmgDb As String = reader.ReadLine()
            Dim nmgLogin As String = reader.ReadLine()
            Dim nmgPwd As String = reader.ReadLine()

            reader.Close()

            NormingDB = New DataConnector(nmgServer, nmgDb, nmgLogin, nmgPwd)
        Catch ex As Exception
            GetLog(ex.Message)
        End Try

    End Sub

    Sub ALAutoUpdate()
        tbl = New DataTable

        tbl = NormingDB.SelectField("HREMP_EMPID", "HREMP", "HREMP_STATUS = 1 AND DATEDIFF(DAY, HREMP_LHIREDAY, GETDATE()) = 365", DataConnector.SelectedReturnType.DataTable)

        For i As Integer = 0 To tbl.Rows.Count - 1
            NormingDB.InitUpdateCommand("LVEMPLEAVE")
            NormingDB.SetWhereField("LVEMPLEAVE_EMPID", tbl(i)("HREMP_EMPID").ToString.Trim)
            NormingDB.SetWhereField("LVEMPLEAVE_LEAVECODE", "AL")
            NormingDB.SetUpdateField("LVEMPLEAVE_CARRYDAYS", 15)
            NormingDB.ExecuteUpdateCommand()
        Next

    End Sub

    Sub GetLog(ByVal errTxt As String)
        File.AppendAllText(".\Log\log.txt", String.Format("{0}{1}", Environment.NewLine, "---------------------------------------------------" & Environment.NewLine & Format(Now, "dd-MM-yyyy HH:mm:ss tt") & "" & Environment.NewLine & errTxt))
    End Sub

    Sub Main()
        ReadConfigFile()

        ALAutoUpdate()
    End Sub

End Module
