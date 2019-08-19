Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports SSP.PublicClass
Imports System.Threading
Imports System.Text

Public Enum SSPCompMonthlyReport
    Vendor
    Factory
End Enum


Public Class FormSSPComparisonMonthly
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Public Property department As Department
    Public Property TableName As String
    Public Property ViewName As String

    Private WithEvents bgworker As New BackgroundWorker
    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    'Dim myHashtable As New Hashtable
    Dim myArrayList As New ArrayList
    Dim period1 As Date
    Dim period2 As Date
    Dim status As Boolean = False
    Dim Filename As String
    Dim SelectedDir As String
    Dim DS As DataSet
    Dim MonthlyBS1 As BindingSource
    Dim monthlyBS2 As BindingSource
    Dim VendorBS As BindingSource
    Dim ReportType As SSPCompMonthlyReport
    'Dim dbAdapter1 = New DBAdapter
    Dim sb As New StringBuilder

    Public Sub New(ByVal ReportType As SSPCompMonthlyReport)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.department = SSP.Department.Components
        Me.ViewName = "sopallcomp"
        Me.ReportType = ReportType
    End Sub
    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        'case when v.vendorname isnull then (foo.vendorcode::text) else v.vendorname::text end as vendorname

        'sb.Append("select id,to_char(sm.monthly,'YYYY Mon') as monthname,yearweek::text as weeklyname,weekly,sm.monthly from sspmonthlytable  sm left join sspweekly sw on sw.sspweeklyid = sm.weekly order by sm.monthly desc;")
        sb.Append("with w as (select first_value(sspweeklyid) over (partition by monthly order by yearweek asc) as id,monthly from sspweekly )," &
                  " d as (select distinct * from w ) select id,to_char(d.monthly,'YYYY Mon') as monthname,yearweek::text as weeklyname,d.monthly from d left join sspweekly sw on sw.sspweeklyid = d.id order by d.monthly desc;")


        'sb.Append("select yearweek,monthly,sspweeklyid from sspweekly order by yearweek desc;")
        'sb.Append("select 0 as vendorcode, 'SELECT ALL' as vendorname union all" &
        '          " Select 1 as vendorcode,'ALL VENDOR' as vendorname  union All" &
        '          " (select foo.vendorcode as vendorcode, v.vendorname as vendorname from (select distinct ssp.vendorcode from sspcomp as ssp) as foo" &
        '          " left join vendor v on v.vendorcode = foo.vendorcode order by v.vendorname);")
        sb.Append("select 0 as vendorcode, 'SELECT ALL' as vendorname union all" &
                  " Select 1 as vendorcode,'ALL VENDOR' as vendorname  union All" &
                  " (select foo.vendorcode as vendorcode,case when v.vendorname isnull then (foo.vendorcode::text) else v.vendorname::text end as vendorname from (select distinct ssp.vendorcode from sspcomp as ssp) as foo" &
                  " left join vendor v on v.vendorcode = foo.vendorcode order by v.vendorname);")
        sb.Append("select 0 as factory, 'SELECT ALL' as factoryname union all" &
                 " Select 1 as factory,'ALL Factory' as factoryname  union All" &
                 " (select distinct m.marketid as factory,m.market factoryname from sspcomp s" &
                 " left join sspmarket m on m.marketid = s.marketid" &
                 " where m.market <> '#N/A' order by m.market );")
        'sb.Append("select id,to_char(sm.monthly,'YYYY Mon') as monthname,yearweek::text as weeklyname,weekly,sm.monthly from sspmonthlytable  sm left join sspweekly sw on sw.sspweeklyid = sm.weekly order by sm.monthly desc limit 1;")
        sb.Append("with w as (select first_value(sspweeklyid) over (partition by monthly order by yearweek asc) as id,monthly from sspweekly )," &
                  " d as (select distinct * from w ) select id,to_char(d.monthly,'YYYY Mon') as monthname,yearweek::text as weeklyname,d.monthly from d left join sspweekly sw on sw.sspweeklyid = d.id order by d.monthly desc;")

        If DBAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Monthly"
                DS.Tables(1).TableName = "Vendor"
                DS.Tables(2).TableName = "Factory"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            MonthlyBS1 = New BindingSource
                            monthlyBS2 = New BindingSource
                            VendorBS = New BindingSource

                            MonthlyBS1.DataSource = DS.Tables(0)
                            monthlyBS2.DataSource = DS.Tables(3)

                            ComboBox1.DataSource = MonthlyBS1
                            ComboBox1.DisplayMember = "monthname"
                            ComboBox1.ValueMember = "id"

                            ComboBox2.DataSource = monthlyBS2
                            ComboBox2.DisplayMember = "monthname"
                            ComboBox2.ValueMember = "id"

                            If ReportType = SSPCompMonthlyReport.Vendor Then
                                CheckedListBox1.DataSource = DS.Tables(1).DefaultView
                                CheckedListBox1.ValueMember = DS.Tables(1).Columns(0).ToString
                                CheckedListBox1.DisplayMember = DS.Tables(1).Columns(1).ToString
                            Else
                                CheckedListBox1.DataSource = DS.Tables(2).DefaultView
                                CheckedListBox1.ValueMember = DS.Tables(2).Columns(0).ToString
                                CheckedListBox1.DisplayMember = DS.Tables(2).Columns(1).ToString
                            End If
                            

                            'CheckedListBox.DataSource = DataSet.Tables(0).DefaultView
                            'CheckedListBox.ValueMember = DataSet.Tables(0).Columns(0).ToString
                            'CheckedListBox.DisplayMember = DataSet.Tables(0).Columns(0).ToString

                            'CBBS = New BindingSource
                            'Dim pk(0) As DataColumn
                            'pk(0) = DS.Tables(0).Columns("id")
                            'DS.Tables(0).PrimaryKey = pk
                            'DS.Tables(0).Columns("id").AutoIncrement = True
                            'DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            'DS.Tables(0).Columns("id").AutoIncrementStep = -1

                            'Dim monthlycol As DataColumn = DS.Tables(0).Columns("monthly")
                            'Dim weeklycol As DataColumn = DS.Tables(0).Columns("weekly")


                            'DS.Tables(0).Constraints.Add(New UniqueConstraint(monthlycol))
                            'DS.Tables(0).Constraints.Add(New UniqueConstraint(weeklycol))


                            'MonthlyBS.DataSource = DS.Tables(0)
                            'CBBS.DataSource = DS.Tables(1)
                            'DataGridView1.AutoGenerateColumns = False
                            'DataGridView1.DataSource = MonthlyBS
                            'DataGridView1.RowTemplate.Height = 22

                            ''TextBox1.DataBindings.Clear()

                            'ComboBox1.DataSource = DS.Tables(1)
                            'ComboBox1.DisplayMember = "yearweek"
                            'ComboBox1.ValueMember = "sspweeklyid"
                            'ComboBox1.DataBindings.Add(New Binding("SelectedValue", MonthlyBS, "weekly", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            'If IsNothing(MonthlyBS.Current) Then
                            '    ComboBox1.SelectedIndex = -1
                            'End If

                            'TextBox1.DataBindings.Add(New Binding("Text", MonthlyBS, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))


                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Report1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ''init Combobox
        'Application.DoEvents()
        ''tablename = ssp
        'TableName = "sspcomp"
        ''Dim sqlstr As String = "select period from " & TableName & " group by period order by period desc;"
        'Dim sqlstr As String = "select sm.id,sm.monthly,sw.yearweek from sspmonthlytable sm" &
        '                       " left join sspweekly sw on sw.sspweeklyid = sm.weekly;"
        'dbtools1.FillCombobox(ComboBox1, sqlstr)
        'dbtools1.FillCombobox(ComboBox2, sqlstr)
        'sqlstr = "select 0 as vendorcode, 'SELECT ALL' as vendorname union all " & _
        '         " Select 1 as vendorcode,'ALL VENDOR' as vendorname  union All" & _
        '         " (select foo.vendorcode as vendorcode, v.vendorname as vendorname from (" & _
        '         " select distinct ssp.vendorcode from " & TableName & " as ssp) as foo" & _
        '         " left join vendor v on v.vendorcode = foo.vendorcode " & _
        '         " order by v.vendorname)"
        'dbtools1.FillCheckedListBoxDataSource(CheckedListBox1, sqlstr)

        loaddata()

    End Sub


    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Dim clb = DirectCast(sender, CheckedListBox)
        Select Case clb.SelectedIndex
            Case 0
                Dim chkstate As CheckState
                chkstate = clb.GetItemCheckState(0)
                For i = 0 To clb.Items.Count - 1
                    clb.SetItemChecked(i, chkstate)
                Next
            Case Else
                clb.SetItemChecked(0, 0)
                Dim mycountlist As Integer = countlist(clb)
                If clb.Items.Count = mycountlist + 1 Then
                    clb.SetItemChecked(0, True)
                End If
        End Select
    End Sub

    Private Function countlist(ByVal clb As CheckedListBox) As Integer
        Dim count As Integer = 0
        For i = 0 To clb.Items.Count - 1
            If clb.GetItemCheckState(i) Then
                count += 1
            End If
        Next
        Return count
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'read destination
        'check combobox values

        If ComboBox1.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox1.Select()
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox2.Select()
            Exit Sub
        End If
        Dim drv = ComboBox1.SelectedItem
        Dim drv2 = ComboBox2.SelectedItem
        period1 = drv.item("monthly")
        period2 = drv2.item("monthly")

        myArrayList.Clear()
        For i = 1 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemCheckState(i) Then
                Dim dr = DirectCast(CheckedListBox1.Items(i), DataRowView)
                myArrayList.Add(New cblList With {.id = dr.Item(0),
                                                  .name = dr.Item(1)})
            End If
        Next
        If myArrayList.Count = 0 Then
            MessageBox.Show("Please select from the list.")
            Exit Sub
        End If
        Button2.Enabled = False
        CheckedListBox1.GetItemCheckState(0)
        If Not bgworker.IsBusy Then
            Using directorybrowser As New FolderBrowserDialog
                directorybrowser.Description = "Which directory do you want to use?"
                If (directorybrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                    Try
                        SelectedDir = directorybrowser.SelectedPath
                        bgworker.WorkerReportsProgress = True
                        bgworker.WorkerSupportsCancellation = True
                        bgworker.RunWorkerAsync()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            End Using
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If



    End Sub


    Private Sub bgworker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgworker.DoWork

        For i = 0 To myArrayList.Count - 1
            Dim errmsg As String = String.Empty
            Filename = SelectedDir & "\" & "SSPComparisonMonthly-" & DirectCast(myArrayList.Item(i), cblList).name & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
            bgworker.ReportProgress(1, DirectCast(myArrayList.Item(i), cblList).name)
            bgworker.ReportProgress(2, i + 1 & " of " & myArrayList.Count)
            If Me.department = SSP.Department.FinishGoods Then
                status = GenerateExcel(i, Filename, errmsg)
            Else
                status = GenerateExcelComponents(i, Filename, errmsg)
            End If

        Next
        bgworker.ReportProgress(3, "Done")
        bgworker.ReportProgress(4, "Enabled button")

    End Sub

    Private Function GenerateExcel(ByVal list As Integer, ByVal Filename As String, ByVal errmsg As String) As Boolean
        Dim result As Boolean = False
        Dim stopwatch As New Stopwatch
        stopwatch.Start()
        Dim mylabel As String = String.Empty
        'Open Excel
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oRange As Excel.Range = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        'Need these variable to kill excel
        'Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        'Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            bgworker.ReportProgress(3, "Create Object Excel....")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            Application.DoEvents()
            'oXl.Visible = True
            'get process pid
            'aprocesses = Process.GetProcesses
            'For i = 0 To aprocesses.GetUpperBound(0)
            '    If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
            '        aprocess = aprocesses(i)
            '        Exit For
            '    End If
            '    Application.DoEvents()
            'Next
            oXl.Visible = False
            oXl.DisplayAlerts = False
            bgworker.ReportProgress(3, "Opening Template...")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            'Go to worksheetData
            oSheet = oWb.Worksheets(2)

            'Get Data passing sqlstr,worksheet for data
            oWb.Worksheets(2).select()

            'Dim firstVar As Integer = Math.Max(CInt(period1), CInt(period2))
            'Dim LastVar As Integer = Math.Min(CInt(period1), CInt(period2))

            Dim check = DirectCast(myArrayList.Item(list), cblList)
            Dim vendorfilter As String = String.Empty
            If check.id <> 1 Then
                vendorfilter = " and vendorcode = " & check.id
            End If
            'viewname = sopall
            'Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                       " where period = " & firstVar & vendorfilter
            ' StringBuilder1.Append(sqlstr)
            ' StringBuilder1.Append(" Union all ")

            '' sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * -1 as ""OrderUnConfirmed"", orderconfirmed * -1 as ""OrderConfirmed"", forecast * -1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            ''                        " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "
            'sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * 1 as ""OrderUnConfirmed"", orderconfirmed * 1 as ""OrderConfirmed"", forecast * 1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                      " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "

            'StringBuilder1.Append(sqlstr)
            bgworker.ReportProgress(3, "DB Query...")
            ExcelStuff.FillDataSource(oWb, 3, StringBuilder1.ToString, dbtools1)
            oSheet = oWb.Worksheets(3)
            oSheet.Columns("L:O").NumberFormat = "0_);[Red](0)"
            bgworker.ReportProgress(3, "Generating PivotTable...")
            'set DbRange
            oWb.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET('" & oSheet.Name & "'!R1C1,0,0,COUNTA('" & oSheet.Name & "'!C1),COUNTA('" & oSheet.Name & "'!R1))")
            oSheet.Name = "Data"
            'Go To Worksheet(1)
            oSheet = oWb.Worksheets(1)
            oWb.Worksheets(1).select()

            oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)

            oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight3"
            oSheet.PivotTables("PivotTable1").ShowTableStyleRowStripes = True

            oSheet.PivotTables("PivotTable1").columngrand = False
            oSheet.PivotTables("PivotTable1").rowgrand = True
            oSheet.PivotTables("PivotTable1").ingriddropzones = True
            oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            'Calculated Field
            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)

            'add PageField
            oSheet.PivotTables("PivotTable1").PivotFields("Market").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            If check.id <> 1 Then
                oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").CurrentPage = check.name
            End If
            'add Rowfields
            oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

            'add columnfield
            oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            'For i = LastVar + 1 To firstVar - 1
            '    Try
            '        oSheet.PivotTables("PivotTable1").PivotFields("Period").PivotItems(i.ToString).Visible = False
            '    Catch ex As Exception
            '    End Try

            'Next

            'add datafield
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), "Sum of Forecast", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").PivotFields("Sum of Forecast").NumberFormat = "#_);[Red](#)"

            'sort column period
            oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            oSheet.Cells.Font.Size = 9
            oSheet.Cells.EntireColumn.AutoFit()
            oSheet.Name = "PivotCompare-Forecast"


            oSheet = oWb.Worksheets(2)
            oWb.Worksheets(2).select()


            'oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            oWb.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(TableDestination:=oSheet.Name & "!R6C1", TableName:="PivotTable1", DefaultVersion:=Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight3"
            oSheet.PivotTables("PivotTable1").ShowTableStyleRowStripes = True

            oSheet.PivotTables("PivotTable1").columngrand = False
            oSheet.PivotTables("PivotTable1").rowgrand = False
            oSheet.PivotTables("PivotTable1").ingriddropzones = True
            oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            ''Calculated Field
            'oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)

            'add PageField
            oSheet.PivotTables("PivotTable1").PivotFields("Market").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            If check.id <> 1 Then
                oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").CurrentPage = check.name
            End If
            'add Rowfields
            oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

            'add columnfield
            oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            'For i = LastVar + 1 To firstVar - 1
            '    Try
            '        oSheet.PivotTables("PivotTable1").PivotFields("Period").PivotItems(i.ToString).Visible = False
            '    Catch ex As Exception
            '    End Try

            'Next

            'add datafield
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Requirement"), "Sum of Requirement", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").PivotFields("Sum of Requirement").NumberFormat = "#_);[Red](#)"

            'sort column period
            oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            oSheet.Cells.Font.Size = 9
            oSheet.Cells.EntireColumn.AutoFit()
            oSheet.Name = "PivotCompare-TotalRequirement"

            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)

            stopwatch.Stop()
            mylabel = "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds
            bgworker.ReportProgress(3, "Saving File..." & mylabel)
            oWb.Worksheets(1).select()
            oWb.SaveAs(Filename)

            result = True
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            bgworker.ReportProgress(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
            releaseComObject(oRange)
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'If Not aprocess Is Nothing Then
                '    aprocess.Kill()
                'End If
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try

        If result And myArrayList.Count = 1 Then
            If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Filename)
            End If
        End If


        Return True
    End Function

    Private Sub bgworker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgworker.ProgressChanged
        Select Case e.ProgressPercentage
            Case 1
                TextBox1.Text = e.UserState
            Case 2
                TextBox2.Text = e.UserState
            Case 3
                TextBox3.Text = e.UserState
            Case 4
                Button2.Enabled = True
        End Select
    End Sub

    Private Sub bgworker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgworker.RunWorkerCompleted

    End Sub

    Private Function GenerateExcelComponents(ByVal list As Integer, ByVal Filename As String, ByVal errmsg As String) As Boolean
        Dim result As Boolean = False
        Dim stopwatch As New Stopwatch
        stopwatch.Start()
        Dim mylabel As String = String.Empty
        'Open Excel
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oRange As Excel.Range = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        'Need these variable to kill excel
        'Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        'Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            bgworker.ReportProgress(3, "Create Object Excel....")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            Application.DoEvents()

            oXl.Visible = False
            oXl.DisplayAlerts = False
            bgworker.ReportProgress(3, "Opening Template...")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\SSPComparisonMonthlyTemplate01.xltx")

            'Go to worksheetData
            oSheet = oWb.Worksheets(2)

            'Get Data passing sqlstr,worksheet for data
            oWb.Worksheets(2).select()

            'Dim firstVar As Integer = Math.Max(CInt(period1), CInt(period2))
            'Dim LastVar As Integer = Math.Min(CInt(period1), CInt(period2))

            Dim check = DirectCast(myArrayList.Item(list), cblList)
            Dim vendorfilter As String = String.Empty
            If check.id <> 1 Then
                If ReportType = SSPCompMonthlyReport.Vendor Then
                    vendorfilter = " and vendorcode = " & check.id
                Else
                    vendorfilter = String.Format(" and market = '{0}'", check.name.Replace("""", """"""))
                End If

            End If
            'viewname = sopall
            'Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                        " where period = " & firstVar & vendorfilter
            'StringBuilder1.Append(sqlstr)
            'StringBuilder1.Append(" Union all ")

            ''sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * -1 as ""OrderUnConfirmed"", orderconfirmed * -1 as ""OrderConfirmed"", forecast * -1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            ''                        " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "

            'sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * 1 as ""OrderUnConfirmed"", orderconfirmed * 1 as ""OrderConfirmed"", forecast * 1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                       " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "

            'Dim sqlstr = String.Format("select * from sspcomparisonmonthly where period >= '{0:yyyyMM}' and period <= '{1:yyyyMM}' {2} ", period1, period2, vendorfilter)
            Dim sqlstr = String.Format("select * from sp_sspcomparisonmonthly({0:yyyyMM},{1:yyyyMM},'{2}') ", period1, period2, vendorfilter)

            StringBuilder1.Append(sqlstr)
            bgworker.ReportProgress(3, "DB Query...")
            ExcelStuff.FillDataSource(oWb, 3, StringBuilder1.ToString, dbtools1)
            bgworker.ReportProgress(3, "Generating PivotTable....")
            oSheet = oWb.Worksheets(3)
            'oSheet.Columns("L:O").NumberFormat = "0_);[Red](0)"
            'bgworker.ReportProgress(3, "Generating PivotTable...")
            ''set DbRange
            oWb.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET('" & oSheet.Name & "'!R1C1,0,0,COUNTA('" & oSheet.Name & "'!C1),COUNTA('" & oSheet.Name & "'!R1))")
            oSheet.Name = "Data"
            'Go To Worksheet(1)

            If ReportType = SSPCompMonthlyReport.Vendor Then
                oSheet = oWb.Worksheets(1)
                oWb.Worksheets(1).select()
                'oXl.Visible = True
                oSheet.PivotTables("PivotTable1").PivotCache.refresh()
                With oSheet.PivotTables("PivotTable1").PivotFields("% of difference")
                    .Calculation = Excel.XlPivotFieldCalculation.xlPercentDifferenceFrom
                    .basefield = "period"
                    .BaseItem = String.Format("{0:yyyyMM}", period2)
                    .numberformat = "0%"
                End With
                oSheet.PivotTables("PivotTable1").DisplayErrorString = True
                oWb.Worksheets(2).delete()
            Else
                oSheet = oWb.Worksheets(2)
                oWb.Worksheets(2).select()
                oSheet.PivotTables("PivotTable2").PivotCache.refresh()
                With oSheet.PivotTables("PivotTable2").PivotFields("% of difference")
                    '.BaseItem = String.Format("{0:yyyyMM}", period2)
                    .Calculation = Excel.XlPivotFieldCalculation.xlPercentDifferenceFrom
                    .basefield = "period"
                    .BaseItem = String.Format("{0:yyyyMM}", period2)
                End With
                oSheet.PivotTables("PivotTable2").DisplayErrorString = True
                oWb.Worksheets(1).delete()
            End If
            

           

            'oSheet.PivotTables("PivotTable2").PivotFields("market").Caption = "Factory"

            If ReportType = SSPCompMonthlyReport.Vendor Then

            Else

            End If

            'oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)

            'oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight3"
            'oSheet.PivotTables("PivotTable1").ShowTableStyleRowStripes = True

            'oSheet.PivotTables("PivotTable1").columngrand = False
            'oSheet.PivotTables("PivotTable1").rowgrand = False
            'oSheet.PivotTables("PivotTable1").ingriddropzones = True
            'oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            ''Calculated Field
            'oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)

            ''add PageField
            'oSheet.PivotTables("PivotTable1").PivotFields("Market").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'If check.id <> 1 Then
            '    oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").CurrentPage = check.name
            'End If
            ''add Rowfields
            'oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            ''remove subtotal
            'oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            ''oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            ''oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

            ''add columnfield
            'oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            ''For i = LastVar + 1 To firstVar - 1
            ''    Try
            ''        oSheet.PivotTables("PivotTable1").PivotFields("Period").PivotItems(i.ToString).Visible = False
            ''    Catch ex As Exception
            ''    End Try

            ''Next

            ''add datafield
            'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), "Sum of Forecast", Excel.XlConsolidationFunction.xlSum)
            'oSheet.PivotTables("PivotTable1").PivotFields("Sum of Forecast").NumberFormat = "#_);[Red](#)"

            ''sort column period
            'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            'oSheet.Cells.Font.Size = 9
            'oSheet.Cells.EntireColumn.AutoFit()
            'oSheet.Name = "PivotCompare-Forecast"


            'oSheet = oWb.Worksheets(2)
            'oWb.Worksheets(2).select()


            ''oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            'oWb.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(TableDestination:=oSheet.Name & "!R6C1", TableName:="PivotTable1", DefaultVersion:=Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            'oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight3"
            'oSheet.PivotTables("PivotTable1").ShowTableStyleRowStripes = True

            'oSheet.PivotTables("PivotTable1").columngrand = False
            'oSheet.PivotTables("PivotTable1").rowgrand = False
            'oSheet.PivotTables("PivotTable1").ingriddropzones = True
            'oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            ' ''Calculated Field
            ''oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)

            ''add PageField
            'oSheet.PivotTables("PivotTable1").PivotFields("Market").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'If check.id <> 1 Then
            '    oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").CurrentPage = check.name
            'End If
            ''add Rowfields
            'oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            ''oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            ''remove subtotal
            'oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            ''oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            ''oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

            ''add columnfield
            'oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            ''For i = LastVar + 1 To firstVar - 1
            ''    Try
            ''        oSheet.PivotTables("PivotTable1").PivotFields("Period").PivotItems(i.ToString).Visible = False
            ''    Catch ex As Exception
            ''    End Try

            ''Next

            ''add datafield
            'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Requirement"), "Sum of Requirement", Excel.XlConsolidationFunction.xlSum)
            'oSheet.PivotTables("PivotTable1").PivotFields("Sum of Requirement").NumberFormat = "#_);[Red](#)"

            ''sort column period
            'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            'oSheet.Cells.Font.Size = 9
            'oSheet.Cells.EntireColumn.AutoFit()
            'oSheet.Name = "PivotCompare-TotalRequirement"

            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)

            stopwatch.Stop()
            mylabel = "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds
            bgworker.ReportProgress(3, "Saving File..." & mylabel)
            oWb.Worksheets(1).select()
            oWb.SaveAs(Filename)

            result = True
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            bgworker.ReportProgress(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
            releaseComObject(oRange)
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'If Not aprocess Is Nothing Then
                '    aprocess.Kill()
                'End If
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try

        If result And myArrayList.Count = 1 Then
            If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Filename)
            End If
        End If


        Return True
    End Function


End Class