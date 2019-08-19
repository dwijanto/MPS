Imports DJLib
Imports DJLib.Dbtools
Imports Npgsql
Public Class WorkingDaysParam
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Private DataTable As DataTable

    Private Sub WorkingDaysParam_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not DesignMode Then
            Call loadData()
        End If
    End Sub

    Private Sub loadData()
        Dim sqlstr As String = "select ivalue from paramhd where paramname ='workingdays'"
        DataTable = dbtools1.getData(sqlstr)
        Try
            TextBox1.Text = DataTable.Rows(0).Item(0).ToString
        Catch ex As Exception
            TextBox1.Text = String.Empty
        End Try
    End Sub

    Public Sub UpdateData()
        Dim sqlstr = "Update paramhd set ivalue = @data where paramname = 'workingdays'"
        Dim conn As NpgsqlConnection = New NpgsqlConnection(dbtools1.getConnectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        Try
            conn.Open()
            cmd.Parameters.Add("@data", NpgsqlTypes.NpgsqlDbType.Integer).Value = CInt(TextBox1.Text)
            Dim myRet = cmd.ExecuteNonQuery
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57) Or e.KeyChar = ChrW(Keys.Back)) Then
            Beep()
            e.KeyChar = ChrW(0)
        End If
    End Sub
End Class
