Imports DJLib
Imports DJLib.Dbtools
Imports Npgsql
Public Class Monthly
    Private dbtools1 As New Dbtools(myUserid, myPassword)
    Private ConnectionString As String = dbtools1.getConnectionString
    Private DataAdapter As New NpgsqlDataAdapter
    Private monthlyDataTable As DataTable
    Private bindingsource1 As New BindingSource
    Private CurrencyManager1 As CurrencyManager
    Private conn As NpgsqlConnection
    Private sqlstr As String

    Public AllowValidate As Boolean = False
    Enum MonthlyEnum
        period
        mydate
        sspmonthlyid
    End Enum

    'Private Class sspmonthlywd
    '    Public Property mydate As Date
    '    Public Property workingdays As String
    '    Public Property sspmonthlywdid As String
    'End Class

    Private Sub Monthly_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AllowValidate = False
        If Not DesignMode Then
            Dim sqlstr = "Select period,mydate,sspmonthlyid from sspmonthly order by mydate desc"
            monthlyDataTable = New DataTable
            monthlyDataTable = getData(sqlstr)
            Call LoadDataGrid(monthlyDataTable)
            'set Primary Key
            Dim keys(0) As DataColumn
            keys(0) = monthlyDataTable.Columns(MonthlyEnum.sspmonthlyid.ToString)
            monthlyDataTable.PrimaryKey = keys
        End If

    End Sub
    Private Function GetData(ByVal sqlstr As String) As DataTable
        Dim DataTable = New DataTable()
        Try
            DataAdapter = New NpgsqlDataAdapter(sqlstr, ConnectionString)
            monthlyDataTable.Locale = System.Globalization.CultureInfo.InvariantCulture
            DataAdapter.Fill(DataTable)

        Catch ex As NpgsqlException
        End Try
        Return DataTable
    End Function
    'Private Sub DeleteCmd(ByVal id As Long)
    '    conn = New NpgsqlConnection(ConnectionString)
    '    Try
    '        conn.Open()
    '        sqlstr = "delete from  sspmonthlywd  where sspmonthlywdid=@id"

    '        Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
    '        command.Parameters.Add(New Npgsql.NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = id
    '        Dim lnewid As Long = command.ExecuteNonQuery
    '        conn.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    'Private Function Insertcmd(ByVal sspmonthlywd As sspmonthlywd) As Long
    '    conn = New NpgsqlConnection(ConnectionString)
    '    Dim lnewid As Long
    '    Try
    '        conn.Open()
    '        sqlstr = "insert into sspmonthlywd(workingdays) values(null);select currval('sspmonthlywd_sspmonthlywdid_seq');"
    '        Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

    '        lnewid = command.ExecuteScalar
    '        'Return lnewid
    '    Catch ex As Exception
    '    Finally
    '        conn.Close()
    '    End Try
    '    Return lnewid

    'End Function

    'Private Sub UpdateCmd(ByVal monthly As sspmonthlywd)
    '    conn = New NpgsqlConnection(ConnectionString)

    '    Try
    '        conn.Open()
    '        sqlstr = "Update sspmonthlywd set mydate=@mydate,workingdays=@workingdays where sspmonthlywdid=@sspmonthlywdid"
    '        Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
    '        command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspmonthlywdid", NpgsqlTypes.NpgsqlDbType.Integer)).Value = monthly.sspmonthlywdid
    '        command.Parameters.Add(New Npgsql.NpgsqlParameter("@mydate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CDate(monthly.mydate.Year & "-" & monthly.mydate.Month & "-1")
    '        command.Parameters.Add(New Npgsql.NpgsqlParameter("@workingdays", NpgsqlTypes.NpgsqlDbType.Integer)).Value = IIf(monthly.workingdays = "", DBNull.Value, monthly.workingdays)
    '        Dim lnewid As Long = command.ExecuteScalar
    '        conn.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub LoadDataGrid(ByVal DataTable As DataTable)
        DataGridView1.DataSource = Nothing
        CurrencyManager1 = Nothing
        bindingsource1.DataSource = DataTable
        DataGridView1.DataSource = bindingsource1
        'With DataGridView1
        '    .DataSource = bindingsource1
        '    .RowsDefaultCellStyle.BackColor = Color.Bisque
        '    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        '    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        '    .AutoSize = True
        'End With
        With DataGridView1
            .AutoSize = True
            .TopLeftHeaderCell.Value = "Monthly"
            '.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
            '.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            '.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader)
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .RowsDefaultCellStyle.BackColor = Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            .Columns.Item(MonthlyEnum.mydate).ReadOnly = True
            .Columns.Item(MonthlyEnum.period).ReadOnly = True
            .Columns.Item(MonthlyEnum.mydate).DefaultCellStyle.Format = "dd-MMM-yyyy"
        End With
        'Hide Columns
        DataGridView1.Columns.Item(MonthlyEnum.sspmonthlyid).Visible = False

        DataGridView1.Columns.Item(MonthlyEnum.mydate.ToString).HeaderText = "Monthly"
        DataGridView1.Columns.Item(MonthlyEnum.period.ToString).HeaderText = "Monthly Period"

        'DataGridView1.Columns.Item(monthlyTable.mydate.ToString).DefaultCellStyle.Format = "MMMM-yyyy"
        'AddDateTimePickerColumns()

        CurrencyManager1 = CType(Me.BindingContext(bindingsource1), CurrencyManager)
    End Sub
    Private Sub AddDateTimePickerColumns()
        Dim col As New CalendarColumn1()
        col.ShowCheckBox = False
        col.CustomFormat = "MMMM-yyyy"
        col.Format = "MMMM-yyyy"
        col.DataPropertyName = "mydate"
        col.HeaderText = "Monthly Period"
        col.Width = 120
        col.SortMode = DataGridViewColumnSortMode.Automatic
        Dim style As DataGridViewCellStyle = New DataGridViewCellStyle
        style.Format = "MMMM-yyyy"
        DataGridView1.Columns.Insert(1, col)
        DataGridView1.Columns.Item(3).DefaultCellStyle = style
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        e.Cancel = False
    End Sub

    'Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
    '    Dim mylist() As String = Split(e.Control.ToString, ",")
    '    If mylist(0) = "System.Windows.Forms.DataGridViewTextBoxEditingControl" Then
    '        Dim tb As TextBox = e.Control
    '        tb.Multiline = True
    '        tb.WordWrap = True
    '    End If
    '    AllowValidate = True
    'End Sub

    'Private Sub DataGridView1_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowValidated
    '    Dim DataRow1 As DataRow = Nothing
    '    If CheckRowStateChanged(DataRow1) Then
    '        Call updateRow(DataRow1)
    '    End If
    '    AllowValidate = False
    'End Sub

    'Private Sub DataGridView1_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.RowValidating
    '    'Validating after EditingControlShowing 
    '    If AllowValidate Then
    '        If Not CurrencyManager1 Is Nothing Then
    '            If DataGridView1.Item(monthlyTable.sspmonthlywdid.ToString, CurrencyManager1.Position).Value.ToString.Length > 0 Then
    '                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
    '                Dim workingdayscell As DataGridViewCell = row.Cells(monthlyTable.workingdays.ToString)
    '                Dim mydatecell As DataGridViewCell = row.Cells(monthlyTable.mydate.ToString)


    '                e.Cancel = Not (isCellGood(workingdayscell, errormessage:="Missing Working Days.") And
    '                                isCellDateGood(mydatecell))

    '            End If
    '        End If
    '    End If
    'End Sub
    'Public Sub updateRow(ByVal DataRow As DataRow)
    '    If Not CurrencyManager1 Is Nothing Then
    '        Try
    '            Call UpdateCmd(New sspmonthlywd With {.sspmonthlywdid = DataRow.Item(monthlyTable.sspmonthlywdid.ToString).ToString,
    '                                            .mydate = DataRow.Item(monthlyTable.mydate.ToString).ToString,
    '                                            .workingdays = DataRow.Item(monthlyTable.workingdays.ToString).ToString})
    '            'Commit transaction
    '            DataRow.AcceptChanges()
    '        Catch ex As Exception
    '            'MsgBox(ex.Message)
    '        End Try
    '    End If
    'End Sub
    'Public Function CheckRowStateChanged(ByRef DataRow As DataRow) As Boolean
    '    If Not CurrencyManager1 Is Nothing Then
    '        Try
    '            Dim pkey(0) As Object
    '            pkey(0) = DataGridView1.Item(monthlyTable.sspmonthlywdid.ToString, CurrencyManager1.Position).Value
    '            DataRow = monthlyDataTable.Rows.Find(pkey)
    '            'check any rowchanges
    '            If Not (DataRow.RowState = DataRowState.Unchanged) Then
    '                Return True
    '            End If
    '        Catch ex As Exception

    '        End Try

    '    End If
    '    Return False
    'End Function
    'Private Function isCellGood(ByVal cell As DataGridViewCell, ByVal errormessage As String) As Boolean
    '    cell.ErrorText = errormessage
    '    DataGridView1.Rows(cell.RowIndex).ErrorText = errormessage
    '    If cell.Value Is Nothing Or IsDBNull(cell.Value) Then
    '        Return False
    '    ElseIf Not Integer.TryParse(cell.Value.ToString(), New Integer()) Then
    '        cell.ErrorText = "Must be a number"
    '        DataGridView1.Rows(cell.RowIndex).ErrorText = "Must be a number"
    '        Return False
    '    End If
    '    cell.ErrorText = ""
    '    DataGridView1.Rows(cell.RowIndex).ErrorText = ""
    '    Return True
    'End Function
    'Private Function isCellDateGood(ByVal cell As DataGridViewCell) As Boolean
    '    If cell.Value Is Nothing Or IsDBNull(cell.Value) Then
    '        cell.ErrorText = "Missing Date"
    '        DataGridView1.Rows(cell.RowIndex).ErrorText = "Missing Date"
    '        Return False
    '    Else
    '        Try
    '            DateTime.Parse(cell.Value.ToString())
    '        Catch ex As Exception
    '            cell.ErrorText = "Invalid format"
    '            DataGridView1.Rows(cell.RowIndex).ErrorText = "invalid format"
    '        End Try
    '    End If
    '    cell.ErrorText = ""
    '    DataGridView1.Rows(cell.RowIndex).ErrorText = ""
    '    Return True
    'End Function
    'Private Sub DataGridView1_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.UserAddedRow

    '    Dim myIdx As Long = Insertcmd(New sspmonthlywd With {.workingdays = DataGridView1.Item(monthlyTable.workingdays.ToString, CurrencyManager1.Position).Value.ToString})
    '    DataGridView1.Item(monthlyTable.sspmonthlywdid.ToString, CurrencyManager1.Position).Value = myIdx

    'End Sub

    'Private Sub DataGridView1_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
    '    If MsgBox("Delete this record?", vbYesNoCancel) = DialogResult.Yes Then
    '        MsgBox("Record Deleted")
    '    Else
    '        e.Cancel = True
    '    End If
    'End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Cursor.Current = Cursors.WaitCursor
        DataGridView1.DataSource = Nothing

        Monthly_Load(Me, e)
        Cursor.Current = Cursors.Default
    End Sub

    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    'Dim myIdx As Long = Insertcmd(New sspmonthlywd With {.workingdays = DataGridView1.Item(monthlyTable.workingdays.ToString, CurrencyManager1.Position).Value.ToString})
    '    Dim myIdx As Long = Insertcmd(New sspmonthlywd)
    '    bindingsource1.AddNew()
    '    DataGridView1.Item(monthlyTable.sspmonthlywdid.ToString, CurrencyManager1.Position).Value = myIdx
    '    AllowValidate = True
    'End Sub

    'Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If MessageBox.Show("Are you sure you wish to delete the selected row?", "Delete Row?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
    '        Dim pkey(0) As Object
    '        pkey(0) = DataGridView1.Item(monthlyTable.sspmonthlywdid.ToString, CurrencyManager1.Position).Value
    '        Dim DataRow As DataRow = monthlyDataTable.Rows.Find(pkey)
    '        Try
    '            DeleteCmd(DataRow.Item(monthlyTable.sspmonthlywdid))
    '            DataRow.Delete()
    '        Catch ex As Exception

    '        End Try


    '    End If
    'End Sub
End Class
