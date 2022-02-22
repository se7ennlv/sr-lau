
Imports System.Data.OleDb
Imports System.Collections

Namespace NormingExtension
    Public Class DataConnector
        Private cnn As OleDbConnection
        Private cmd As OleDbCommand
        Private dap As OleDbDataAdapter
        Private insertfields As ArrayList
        Private insertvalues As ArrayList
        Private updatefields As ArrayList
        Private updatevalues As ArrayList
        Private inserttable As String

        Private _arrFieldWhere As ArrayList
        Private _arrValueWhere As ArrayList
        Sub New(ByVal Server As String, ByVal Database As String, ByVal UserID As String, ByVal Password As String)
            cnn = New OleDbConnection("Provider=SQLOledb.1; Data Source=" + Server + "; User ID=" + UserID + "; Password=" + Password + "; Initial Catalog=" + Database + "; Connection Timeout = 100")
            cnn.Open()
            cmd = New OleDbCommand()
            cmd.Connection = cnn

            dap = New OleDbDataAdapter()
        End Sub
        Sub Close()
            cnn.Close()
        End Sub
        Public Enum SelectedReturnType
            SignleValue = 0
            DataTable = 1
        End Enum
        Function SelectField(ByVal FieldName As String, ByVal TableName As String, ByVal Condition As String, ByVal ReturnType As SelectedReturnType) As Object
            Dim qry As String = If(FieldName.Length > 0 And TableName.Length > 0, "SELECT " + FieldName + " FROM " + TableName, "") + If(Condition.Length > 0 And FieldName.Length > 0 And TableName.Length > 0, " WHERE " + Condition, "")

            If qry = "" Then
                Return Nothing
            Else

                dap = New OleDbDataAdapter(qry, cnn)

                Dim tbl As New DataTable
                Dim li As Integer = dap.Fill(tbl)
                If tbl.Rows.Count = 0 Then
                    If ReturnType = SelectedReturnType.SignleValue Then
                        Return Nothing
                    Else
                        Return tbl
                    End If
                Else
                    If ReturnType = SelectedReturnType.SignleValue Then
                        Return tbl.Rows(0)(0)
                    Else
                        Return tbl
                    End If
                End If

            End If
        End Function

        Sub InitInsertCommand(ByVal TableName As String)
            cmd = New OleDbCommand()
            cmd.Connection = cnn
            insertfields = New ArrayList
            insertvalues = New ArrayList
            inserttable = TableName
        End Sub
        Sub InsertField(ByVal FieldName As String, ByVal FieldValue As Object)
            insertfields.Add(FieldName)
            insertvalues.Add(FieldValue)
        End Sub
        Sub ExecuteInsertCommand()
            Dim strfield As String = ""
            Dim strvalue As String = ""
            For i As Integer = 0 To insertfields.Count - 1
                strfield = strfield + insertfields(i).ToString() + ", "
                strvalue = strvalue + "?, "
            Next
            strfield = strfield.Substring(0, strfield.Length - 2)
            strvalue = strvalue.Substring(0, strvalue.Length - 2)
            Dim sqlstring As String = "insert into " + inserttable + " (" + strfield + ") values (" + strvalue + ")"
            cmd.CommandText = sqlstring
            For i As Integer = 0 To insertfields.Count - 1
                cmd.Parameters.AddWithValue("@" + insertfields(i).ToString(), insertvalues(i))
            Next
            Dim k As Integer = cmd.ExecuteNonQuery()
        End Sub

        Sub InitDeleteCommand(ByVal TableName As String)
            cmd = New OleDbCommand()
            cmd.Connection = cnn
            inserttable = TableName
            _arrFieldWhere = New ArrayList
            _arrValueWhere = New ArrayList
        End Sub
        Sub SetWhereField(ByVal FieldName As String, ByVal FieldValue As Object)
            _arrFieldWhere.Add(FieldName)
            _arrValueWhere.Add(FieldValue)
        End Sub
        Sub ExecuteDeleteCommand()

            Dim strwhere As String = ""
            If _arrFieldWhere.Count > 0 Then
                For i As Integer = 0 To _arrFieldWhere.Count - 1
                    strwhere = strwhere + _arrFieldWhere(i).ToString() + "=? and "
                Next
                strwhere = " where " + strwhere.Substring(0, strwhere.Length - 5)
                Dim sqlstring As String = "delete from " + inserttable + strwhere
                cmd.CommandText = sqlstring
                For i As Integer = 0 To _arrFieldWhere.Count - 1
                    cmd.Parameters.AddWithValue("@" + _arrFieldWhere(i).ToString(), _arrValueWhere(i))
                Next
                Dim k As Integer = cmd.ExecuteNonQuery()
            End If
        End Sub

        Sub InitUpdateCommand(ByVal TableName As String)
            cmd = New OleDbCommand()
            cmd.Connection = cnn
            updatefields = New ArrayList
            updatevalues = New ArrayList
            _arrFieldWhere = New ArrayList
            _arrValueWhere = New ArrayList
            inserttable = TableName
        End Sub
        Sub SetUpdateField(ByVal FieldName As String, ByVal FieldValue As Object)
            updatefields.Add(FieldName)
            updatevalues.Add(FieldValue)
        End Sub

        Sub ExecuteUpdateCommand()
            Dim strset As String = ""
            For i As Integer = 0 To updatefields.Count - 1
                strset = strset + updatefields(i).ToString() + "=?, "
            Next
            strset = strset.Substring(0, strset.Length - 2)
            Dim strwhere As String = ""
            If _arrFieldWhere.Count > 0 Then
                For i As Integer = 0 To _arrFieldWhere.Count - 1
                    strwhere = strwhere + _arrFieldWhere(i).ToString() + "=? and "

                Next
                strwhere = " where " + strwhere.Substring(0, strwhere.Length - 5)

            End If

            Dim sqlstring As String = "update " + inserttable + " set " + strset + strwhere
            cmd.CommandText = sqlstring
            For i As Integer = 0 To updatefields.Count - 1
                cmd.Parameters.AddWithValue("@" + updatefields(i).ToString(), updatevalues(i))

            Next
            If _arrFieldWhere.Count > 0 Then
                For i As Integer = 0 To _arrFieldWhere.Count - 1
                    cmd.Parameters.AddWithValue("@" + _arrFieldWhere(i).ToString(), _arrValueWhere(i))

                Next

            End If
            Dim k As Integer = cmd.ExecuteNonQuery()

        End Sub

    End Class

End Namespace

