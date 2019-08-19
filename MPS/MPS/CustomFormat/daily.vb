Imports DJLib
Imports DJLib.Dbtools
Imports Npgsql
Imports SSP.classes.DBClass

Public Class daily
    Enum dailyenum
        dailydate
        holiday
        isholiday
        description
        dailyid
        weekperiod
        monthperiod
    End Enum

    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim conn As NpgsqlConnection
    Dim DataTable As DataTable
    Dim binding
    Private bindingsource1 As New BindingSource
    Private CurrencyManager1 As CurrencyManager
    Dim AllowValidate As Boolean = False

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MessageBox.Show("Do you want to add 1 year Calendar?", "Create Calendar", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            If DataTable.Rows.Count = 0 Then
                'create for 1 year with current year
                Call GenerateData(DateTime.Today)
            Else
                'create for 1 year with latest year
                Dim mydate As Date = DataTable.Rows(0).Item(dailyenum.dailydate)
                mydate = mydate.AddDays(1)
                Call GenerateData(mydate)
            End If
            daily_Load(Me, e)
        End If
    End Sub

    Private Sub daily_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not DesignMode Then
            Call LoadData()
            Call BindData(DataTable)
            DateTimePicker1.Value = DateTime.Today
            DataTable.DefaultView.RowFilter = "dailydate >= #" & DateTimePicker1.Value.Year & "-1-1# and dailydate <= #" & DateTimePicker1.Value.Year & "-12-31#"
        End If
    End Sub

    Private Sub LoadData()
        Dim sqlstr As String = "Select dailydate,holiday,isholiday,description,dailyid,weekperiod,monthperiod from sspdaily order by dailydate desc "
        DataTable = New DataTable
        DataTable = dbtools1.getData(sqlstr)
        Dim keys(0) As DataColumn
        keys(0) = DataTable.Columns(dailyenum.dailyid.ToString)
        DataTable.PrimaryKey = keys

    End Sub

    Private Sub BindData(ByVal datatable As DataTable)
        Try
            DataGridView1.DataSource = Nothing
            CurrencyManager1 = Nothing
            bindingsource1.DataSource = datatable
            DataGridView1.DataSource = bindingsource1
            With DataGridView1
                .AutoSize = True
                .TopLeftHeaderCell.Value = "Daily"
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige

                'Hide Columns
                .Columns.Item(dailyenum.dailyid).Visible = False
                .Columns.Item(dailyenum.monthperiod).Visible = False
                .Columns.Item(dailyenum.weekperiod).Visible = False
                .Columns.Item(dailyenum.holiday).Visible = False
                .Columns.Item(dailyenum.description).Width = 200
                .Columns.Item(dailyenum.dailydate).DefaultCellStyle.Format = "dd-MMM-yyyy"
                .Columns.Item(dailyenum.dailydate).ReadOnly = True

            End With
            DataGridView1.Columns.Item(dailyenum.description.ToString).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            'DataGridView1.Columns.Item(3).DefaultCellStyle.Format = "dd-MMM-yyyy"
            CurrencyManager1 = CType(Me.BindingContext(bindingsource1), CurrencyManager)
            'addCheckboxColumn()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub addCheckboxColumn()
        Dim checkBoxColumn As DataGridViewCheckBoxColumn
        checkBoxColumn = createCheckboxColumn()
        DataGridView1.Columns.Insert(1, checkBoxColumn)
    End Sub
    Private Function createCheckboxColumn() As DataGridViewCheckBoxColumn
        Dim column As New DataGridViewCheckBoxColumn
        With column
            .DataPropertyName = "holiday"
            .HeaderText = "isHoliday"
            .Width = 100
            .FlatStyle = FlatStyle.Flat
            .SortMode = DataGridViewColumnSortMode.Automatic
        End With
        Return column

    End Function

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
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
        'If AllowValidate Then
        '    If Not CurrencyManager1 Is Nothing Then
        '        If DataGridView1.Item(dailyenum.dailyid.ToString, CurrencyManager1.Position).Value.ToString.Length > 0 Then
        '            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
        '            Dim holidaycell As DataGridViewCell = row.Cells(dailyenum.holiday.ToString)
        '            Dim dailydatecell As DataGridViewCell = row.Cells(dailyenum.dailydate.ToString)


        '            e.Cancel = Not (isCellGood(holidaycell, errormessage:="Missing Working Days.") And
        '                            isCellDateGood(dailydatecell))

        '        End If
        '    End If
        'End If
    End Sub
    Public Sub updateRow(ByVal DataRow As DataRow)
        If Not CurrencyManager1 Is Nothing Then
            Try
                Call DailyModel.UpdateCmd(New DailyModel With {.dailyid = DataRow.Item(dailyenum.dailyid.ToString).ToString,
                                                    .holiday = DataRow.Item(dailyenum.holiday.ToString).ToString,
                                                    .description = DataRow.Item(dailyenum.description.ToString).ToString,
                                                               .isholiday = DataRow.Item(dailyenum.isholiday)})
                'Commit transaction
                DataRow.AcceptChanges()
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End If
    End Sub
    Public Function CheckRowStateChanged(ByRef DataRow As DataRow) As Boolean
        If Not CurrencyManager1 Is Nothing Then
            Try
                Dim pkey(0) As Object
                pkey(0) = DataGridView1.Item(dailyenum.dailyid.ToString, CurrencyManager1.Position).Value
                DataRow = DataTable.Rows.Find(pkey)
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

    Private Sub GenerateData(ByVal p1 As Date)
        Cursor.Current = Cursors.WaitCursor
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim EndDate As Date = CDate(p1.Year & "-12-31")
        Dim StartDate As Date = CDate(p1.Year & "-1-1")
        Dim myweek As String = String.Empty
        Dim mymonth As Date
        While True
            Dim myweekresult As Integer = dbtools1.getweekperiod(StartDate)
            mymonth = CDate(StartDate.Year & "-" & StartDate.Month & "-1")

           

            If myweekresult > 50 And StartDate.Month = 1 Then
                myweek = StartDate.Year - 1 & Format(myweekresult, "00")
            Else
                myweek = StartDate.Year & Format(myweekresult, "00")
            End If

            If StartDate.Month = 1 And StartDate.Day = 1 Then
                If StartDate.DayOfWeek <> 1 Then
                    'create week 1
                    WeeklyModel.InsertWeekly(New WeeklyModel With {.startdate = StartDate.AddDays((StartDate.DayOfWeek - 1) * -1),
                                                              .yearweek = myweek,
                                                              .monthly = mymonth,
                                                              .label = myweekresult.ToString})
                End If
            End If

            'Generate Weekly
            If StartDate.DayOfWeek = 1 Then
                WeeklyModel.InsertWeekly(New WeeklyModel With {.startdate = StartDate,
                                                               .yearweek = myweek,
                                                               .monthly = mymonth,
                                                               .label = myweekresult.ToString})
            End If
            If StartDate.Day = 1 Then
                MonthlyModel.InsertMonthly(New MonthlyModel With {.period = StartDate.Year.ToString & Format(StartDate.Month, "00"),
                                                                  .mydate = CDate(StartDate.Year & "-" & StartDate.Month & "-1")})
            End If


            stringbuilder1.Append(myweek & vbTab)
            stringbuilder1.Append(DateFormatyyyyMMdd(mymonth) & vbTab)
            stringbuilder1.Append(DateFormatyyyyMMdd(StartDate) & vbCrLf)
            StartDate = StartDate.AddDays(1)
            If StartDate > EndDate Then
                Exit While
            End If
            Application.DoEvents()
        End While
        Dim sqlstr As String = "copy sspdaily(weekperiod,monthperiod,dailydate) from stdin;"
        If stringbuilder1.ToString <> "" Then
            Dim myreturn As Boolean

            Dim errMessage As String = dbtools1.copy(sqlstr, stringbuilder1.ToString, myreturn)
        End If
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DataTable.DefaultView.RowFilter = "dailydate >= #" & DateTimePicker1.Value.Year & "-1-1# and dailydate <= #" & DateTimePicker1.Value.Year & "-12-31#"
    End Sub
End Class
