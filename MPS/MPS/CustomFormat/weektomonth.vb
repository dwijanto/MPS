Imports DJLib
Imports DJLib.Dbtools
Imports Npgsql
Imports SSP.classes.DBClass
Public Class weektomonth

    Private dbtools1 As New Dbtools(myUserid, myPassword)
    Private ConnectionString As String = dbtools1.getConnectionString
    Private DataAdapter As New NpgsqlDataAdapter
    Private WeeklyDataTable As DataTable
    Private bindingsource1 As New BindingSource
    Private CurrencyManager1 As CurrencyManager
    Private conn As NpgsqlConnection
    Private sqlstr As String
    Private myIdx As Long
    Public AllowValidate As Boolean = False

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub weektomonth_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AllowValidate = False
        If Not DesignMode Then
            Dim sqlstr = "Select * from sspweekly order by yearweek desc"
            WeeklyDataTable = New DataTable
            WeeklyDataTable = GetData(sqlstr)
            Call LoadDataGrid(WeeklyDataTable)
            'set Primary Key
            Dim keys(0) As DataColumn
            keys(0) = WeeklyDataTable.Columns(WeeklyEnum.sspweeklyid.ToString)
            WeeklyDataTable.PrimaryKey = keys
        End If
    End Sub
    
    Private Function GetData(ByVal sqlstr As String) As DataTable
        Dim DataTable = New DataTable()
        Try
            DataAdapter = New NpgsqlDataAdapter(sqlstr, ConnectionString)
            WeeklyDataTable.Locale = System.Globalization.CultureInfo.InvariantCulture
            DataAdapter.Fill(DataTable)
            
        Catch ex As NpgsqlException
        End Try
        Return DataTable
    End Function

    Private Sub LoadDataGrid(ByVal DataTable As DataTable)
        DataGridView1.DataSource = Nothing
        CurrencyManager1 = Nothing
        bindingsource1.DataSource = DataTable
        DataGridView1.DataSource = bindingsource1
        With DataGridView1
            .AutoSize = True
            .TopLeftHeaderCell.Value = "Weekly"
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .RowsDefaultCellStyle.BackColor = Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            .Columns.Item(WeeklyEnum.monthly).DefaultCellStyle.Format = "dd-MMM-yyyy"
            .Columns.Item(WeeklyEnum.StartDate).DefaultCellStyle.Format = "dd-MMM-yyyy"
            .Columns.Item(WeeklyEnum.StartDate).ReadOnly = True
            .Columns.Item(WeeklyEnum.monthly).ReadOnly = True
            .Columns.Item(WeeklyEnum.YearWeek).ReadOnly = True
            .Columns.Item(WeeklyEnum.monthly).Visible = False
            .Columns.Item(WeeklyEnum.crossmonth).HeaderText = "Crossmonth and (Not all week isHoliday)"
        End With
        DataGridView1.Columns.Item(WeeklyEnum.sspweeklyid.ToString).Visible = False
        
        'DataGridView1.Columns.Item(WeekTableEnum.Comments.ToString).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'AddComboBoxColumns()
        'AddDateTimePickerColumns1()
        'AddDateTimePickerColumns()

        'DataGridView1.Columns.Item(3).DefaultCellStyle.Format = "dd-MMM-yyyy"

        CurrencyManager1 = CType(Me.BindingContext(bindingsource1), CurrencyManager)
    End Sub
    Private Sub AddDateTimePickerColumns1()
        Dim col As New CalendarColumn1()
        col.ShowCheckBox = False
        col.DataPropertyName = "monthly"
        col.CustomFormat = "MMMM-yyyy"
        col.Format = "MMMM-yyyy"
        col.HeaderText = "Monthly"
        col.Width = 120

        DataGridView1.Columns.Insert(2, col)
        
    End Sub
    Private Sub AddDateTimePickerColumns()
        Dim col As New CalendarColumn()
        col.DataPropertyName = "startdate"
        col.CustomFormat = "dd-MMM-yyyy"
        col.HeaderText = "Starting Date"
        col.Format = "dd-MMM-yyyy"
        col.Width = 120
        col.SortMode = DataGridViewColumnSortMode.Automatic
        DataGridView1.Columns.Insert(3, col)

    End Sub

    Private Sub AddComboBoxColumns()
        Dim comboboxColumn As DataGridViewComboBoxColumn
        comboboxColumn = CreateComboBoxColumn()
        SetAlternateChoicesUsingDataSource(comboboxColumn)
        comboboxColumn.HeaderText = "Monthly Period"
        DataGridView1.Columns.Insert(2, comboboxColumn)
        'DataGridView1.Columns.Add(comboboxColumn)
    End Sub
    Private Function CreateComboBoxColumn() As DataGridViewComboBoxColumn
        Dim column As New DataGridViewComboBoxColumn()
        With column
            .DataPropertyName = "mydate"
            .HeaderText = "mydate"
            .DropDownWidth = 100
            .Width = 120
            .MaxDropDownItems = 4
            .FlatStyle = FlatStyle.Flat
        End With
        Return column
    End Function

    Private Class monthly
        Public Property id As Long
        Public Property mydate As String
    End Class
    Private Sub SetAlternateChoicesUsingDataSource(ByVal comboboxColumn As DataGridViewComboBoxColumn)
        Dim sqlstr As String = "select sspmonthlywdid,mydate::date from sspmonthlywd order by mydate desc"
        Dim mytable As DataTable = GetData(sqlstr)
        For Each e As Object In mytable.Rows
            comboboxColumn.Items.Add(New monthly With {.id = e.item(0),
                                                       .mydate = Format(e.item(1), "MMMM-yyyy").ToString})
        Next
        comboboxColumn.Items.Add(New monthly With {.id = 0,
                                                       .mydate = ""})
        comboboxColumn.DataPropertyName = "sspmonthlywdid"
        comboboxColumn.DisplayMember = "mydate"
        comboboxColumn.ValueMember = "id"
        'comboboxColumn.Items.Add(DBNull.Value)
        'comboboxColumn.DefaultCellStyle.NullValue = ""
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        e.Cancel = False
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim mylist() As String = Split(e.Control.ToString, ",")
        If mylist(0) = "System.Windows.Forms.DataGridViewTextBoxEditingControl" Then
            Dim tb As TextBox = e.Control
            tb.Multiline = True
            tb.WordWrap = True
        End If
        AllowValidate = True
    End Sub

    Private Sub DataGridView1_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowValidated
        Dim DataRow1 As DataRow = Nothing
        If CheckRowStateChanged(DataRow1) Then
            Call updateRow(DataRow1)
        End If
        AllowValidate = False
    End Sub

    Private Sub DataGridView1_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.RowValidating
        'Validating after EditingControlShowing 
        If AllowValidate Then
            If Not CurrencyManager1 Is Nothing Then
                If DataGridView1.Item(WeeklyEnum.sspweeklyid.ToString, CurrencyManager1.Position).Value.ToString.Length > 0 Then
                    Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                    Dim yearweekcell As DataGridViewCell = row.Cells(WeeklyEnum.YearWeek)
                    Dim startdatecell As DataGridViewCell = row.Cells(WeeklyEnum.StartDate)
                    e.Cancel = Not (isCellGood(yearweekcell, errormessage:="Missing Yearweek.") And
                                    isCellDateGood(startdatecell)
                                   )

                End If
            End If
        End If
    End Sub
    Public Sub updateRow(ByVal DataRow As DataRow)
        If Not CurrencyManager1 Is Nothing Then
            Try
                Call WeeklyModel.UpdateWeekly(New WeeklyModel With {.weeklyid = DataRow.Item(WeeklyEnum.sspweeklyid.ToString).ToString,
                                                .label = DataRow.Item(WeeklyEnum.label.ToString).ToString,
                                                .crossmonth = DataRow.Item(WeeklyEnum.crossmonth.ToString)})
                'Commit Transaction
                DataRow.AcceptChanges()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Function CheckRowStateChanged(ByRef DataRow As DataRow) As Boolean
        If Not CurrencyManager1 Is Nothing Then
            Try
                Dim pkey(0) As Object
                pkey(0) = DataGridView1.Item(WeeklyEnum.sspweeklyid.ToString, CurrencyManager1.Position).Value
                DataRow = WeeklyDataTable.Rows.Find(pkey)
                'check any rowchanges
                If Not (DataRow.RowState = DataRowState.Unchanged) Then
                    Return True
                End If
            Catch ex As Exception

            End Try

        End If
        Return False
    End Function
    Private Function isCellGood(ByVal cell As DataGridViewCell, ByVal errormessage As String) As Boolean
        cell.ErrorText = errormessage
        DataGridView1.Rows(cell.RowIndex).ErrorText = errormessage
        If cell.Value Is Nothing Or IsDBNull(cell.Value) Then
            Return False
        ElseIf Not Integer.TryParse(cell.Value.ToString(), New Integer()) Then
            cell.ErrorText = "Must be a number"
            DataGridView1.Rows(cell.RowIndex).ErrorText = "Must be a number"
            Return False
        End If
        cell.ErrorText = ""
        DataGridView1.Rows(cell.RowIndex).ErrorText = ""
        Return True
    End Function
    Private Function isCellDateGood(ByVal cell As DataGridViewCell) As Boolean
        If cell.Value Is Nothing Or IsDBNull(cell.Value) Then
            cell.ErrorText = "Missing Date"
            DataGridView1.Rows(cell.RowIndex).ErrorText = "Missing Date"
            Return False
        Else
            Try
                DateTime.Parse(cell.Value.ToString())
            Catch ex As Exception
                cell.ErrorText = "Invalid format"
                DataGridView1.Rows(cell.RowIndex).ErrorText = "invalid format"
            End Try
        End If
        cell.ErrorText = ""
        DataGridView1.Rows(cell.RowIndex).ErrorText = ""
        Return True
    End Function
    'Private Sub DataGridView1_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.UserAddedRow

    '    myIdx = Insertcmd(New Weekly With {.comments = DataGridView1.Item(WeekTableEnum.Comments.ToString, CurrencyManager1.Position).Value.ToString})
    '    DataGridView1.Item(WeekTableEnum.weektomonthid.ToString, CurrencyManager1.Position).Value = myIdx

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

        weektomonth_Load(Me, e)
        Cursor.Current = Cursors.Default
    End Sub

    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
    '    bindingsource1.AddNew()
    '    myIdx = Insertcmd(New Weekly With {.comments = DataGridView1.Item(WeekTableEnum.Comments.ToString, CurrencyManager1.Position).Value.ToString})
    '    DataGridView1.Item(WeekTableEnum.weektomonthid.ToString, CurrencyManager1.Position).Value = myIdx
    '    AllowValidate = True
    'End Sub

    'Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
    '    If MessageBox.Show("Are you sure you wish to delete the selected row?", "Delete Row?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
    '        Dim pkey(0) As Object
    '        pkey(0) = DataGridView1.Item(WeekTableEnum.weektomonthid.ToString, CurrencyManager1.Position).Value
    '        Dim DataRow As DataRow = WeeklyDataTable.Rows.Find(pkey)
    '        Try
    '            DeleteCmd(DataRow.Item(WeekTableEnum.weektomonthid))
    '            DataRow.Delete()
    '        Catch ex As Exception
    '        End Try
    '    End If
    'End Sub

    
    Enum WeeklyEnum
        sspweeklyid
        YearWeek
        StartDate
        monthly
        label
        crossmonth
    End Enum

End Class
