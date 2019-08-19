Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MPSReport2
    Public Class MyForm
        Public combobox1 As String
    End Class

    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Private WithEvents backgroundworker1 As New BackgroundWorker
    Dim status As Boolean = False
    Dim FileName As String = String.Empty
    Dim Myform1 As MyForm
    Dim StartingDate As Date
    Dim weekList As String()

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Check Selected CheckedListbox
        If ComboBox1.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox1.Select()
            Exit Sub
        End If
        Dim myreport = New classes.DBClass.Reports With {.YearWeek = ComboBox1.Text}
        StartingDate = myreport.getStartingDate

        Button1.Enabled = False


        If Not (backgroundworker1.IsBusy) Then

            'Dim FileName As String = String.Empty
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                FileName = DirectoryBrowser.SelectedPath & "\" & "Exception Report" & ComboBox1.Text & "-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
                Myform1 = New MyForm With {.combobox1 = ComboBox1.SelectedItem.ToString}
                Label1.Text = ""
                Try
                    backgroundworker1.WorkerReportsProgress = True
                    backgroundworker1.WorkerSupportsCancellation = True
                    backgroundworker1.RunWorkerAsync()
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
       
    End Sub

    Private Sub MPSReport2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sqlstr As String = "select period from ssp group by period order by period desc;"
        dbtools1.FillCombobox(ComboBox1, sqlstr)
        ReDim weekList(ComboBox1.Items.Count - 1)
        For i = 0 To weekList.Length - 1
            weekList(i) = ComboBox1.Items(i)
        Next
    End Sub
    Private Sub backgroundworker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker1.DoWork
        Dim errMsg As String = String.Empty

        status = GenerateExcel(FileName, errMsg)

        If status Then
            backgroundworker1.ReportProgress(1, "Done. " & FileName)
        Else
            backgroundworker1.ReportProgress(1, "Error::" & errMsg)
        End If

    End Sub

    Private Sub backgroundworker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundworker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 1
                TextBox1.Text = e.UserState
            Case 2
                Label1.Text = e.UserState
            Case 3
                'TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub backgroundworker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker1.RunWorkerCompleted
        FormMenu.setBubbleMessage("Export To Excel", "Done")
        If status Then
            'If CheckBox1.Checked Then
            '    Me.Close()
            'End If
        End If
        If status Then
            If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(FileName)
            End If
        End If
        Button1.Enabled = True
    End Sub

    Private Function GenerateExcel(ByRef FileName As String, ByRef errormsg As String) As Boolean

        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim dataset1 As New DataSet

        Dim StopWatch As New Stopwatch
        StopWatch.Start()

        'Open Excel
        Application.DoEvents()

        Cursor.Current = Cursors.WaitCursor
        Dim source As String = FileName
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim Sqlstr As String = String.Empty

        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            Application.DoEvents()
            oXl.Visible = True
            'get process pid
            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
                Application.DoEvents()
            Next
            oXl.Visible = False
            oXl.DisplayAlerts = False
            backgroundworker1.ReportProgress(1, "Opening Template...")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")
            Dim iSheetDAta As Integer = 2
            'Loop for chart
            'Go to worksheetData
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Worksheets(iSheetDAta).select()
            backgroundworker1.ReportProgress(1, "DB Query...")
            Call QueryDataAll(oWb, iSheetDAta)


            oWb.Worksheets(iSheetDAta).select()
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Names.Add(Name:="DBRangeAll", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "DBAll"

            'Generate Chart&Pivot start from worksheet 2
            iSheetDAta = 1
            backgroundworker1.ReportProgress(1, "Generating PivotTable...")
            Call GeneratePivotTable(oWb, iSheetDAta)
            oXl.ActiveWindow.FreezePanes = True

            StopWatch.Stop()
            backgroundworker1.ReportProgress(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            FileName = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            backgroundworker1.ReportProgress(1, "Saving File...")
            oXl.DisplayAlerts = False
            'oWb.Worksheets("DBAll").delete()
            oWb.SaveAs(FileName)
            oXl.DisplayAlerts = True
            result = True
        Catch ex As Exception
            errormsg = ex.Message
        Finally
            backgroundworker1.ReportProgress(1, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                If Not aprocess Is Nothing Then
                    aprocess.Kill()
                End If
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try


        'If result Then
        '    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
        '        Process.Start(FileName)
        '    End If
        'End If
        'Button1.Enabled = True
        Return result

    End Function

    Public Sub QueryDataAll(ByRef owb As Excel.Workbook, ByVal isheet As Integer)
        Dim sqlstr As String = String.Empty
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim FinalQuery As New System.Text.StringBuilder

        Dim period As Integer = CInt(StartingDate.Year.ToString & Format(StartingDate.Month, "00"))
        Dim mycriteria As String = String.Empty
        Dim myWeek = Myform1.combobox1.Substring(4, 2)

        'Check Worksheet
        For i = owb.Worksheets.Count To isheet - 1
            owb.Worksheets.Add(After:=owb.Worksheets(i))
        Next

        'GET MPS DATA

        For i = 0 To 0
            sqlstr = "(SELECT f.period, io.typeofinfo, f.vendorcode, v.vendorname, fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,null::bigint as cmmf,null::char(50) as materialdesc,null::char varying as market, d.startingdate, d.startingdate AS monthly, wl.yearweek as datalabel1," & _
                     " case when wl.label::integer > 9 and wl.label::integer < " & myWeek - i & " then wl.label " & _
                     " when wl.label::integer >= " & myWeek - i & " then '  ' || wl.label  " & _
                     " else ' ' || wl.label " & _
                     " end as datalabel2, " & _
                     " d.datavalue::numeric AS datavalue, dl.weekperiod, dl.dailydate, mwd.count," & _
                     " case when typeofinfo = 'Bottleneck' then d.datavalue::numeric / mwd.count::numeric " & _
                     " else 0 " & _
                     " end AS bottleneck ," & _
                     " case when typeofinfo = 'Supply Plan' then d.datavalue::numeric / mwd.count::numeric " & _
                     " else 0" & _
                     " end AS supplyplan," & _
                     " null::numeric as orderconfirmed,null::numeric as orderunconfirmed,null::numeric as forecast, case when typeofinfo = 'Bottleneck' then d.datavalue::numeric / mwd.count::numeric else 0 end as diff, o.officersebname" & _
                     " FROM sspftycap f" & _
                     " LEFT JOIN sspftycapdata d ON d.ftycapid = f.ftycapid" & _
                     " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = f.sopfamilyid" & _
                     " LEFT JOIN vendor v ON v.vendorcode = f.vendorcode" & _
                     " LEFT JOIN officerseb o on o.ofsebid = v.ofsebid" & _
                     " LEFT JOIN ssptypeofinfo io ON io.ssptypeofinfoid = f.typeofinfoid" & _
                     " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = f.sopfamilyid" & _
                     " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                     " LEFT JOIN sspdaily dl ON dl.monthperiod = d.startingdate" & _
                     " left join sspweekly wl on wl.yearweek = dl.weekperiod" & _
                     " left join sspmonthlywdparam mwd on mwd.monthperiod = dl.monthperiod" & _
                     " where period = " & period - i & "  and not d.ftycapdataid is null " & _
                     " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday and dl.dailydate >= " & DateFormatyyyyMMdd(StartingDate.AddDays(-(i * 4 * 7))) & " and  dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(182 - (i * 4 * 7))) & _
                     " order by upper(sopdescription))"
            stringbuilder1.Append(sqlstr)
            stringbuilder1.Append(" union all ")
        Next


        For i = 0 To 0
            'Query Orderconfirmed
            sqlstr = "(SELECT ssp.period, 'Order confirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,cr.cmmf,c.materialdesc,sm.market, ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                     " case when wl.label::integer > 9 and wl.label::integer < " & myWeek - i & " then wl.label " & _
                     " when wl.label::integer >= " & myWeek - i & " then '  ' || wl.label  " & _
                     " else ' ' || wl.label " & _
                     " end as datalabel2, " & _
                     " ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,null::numeric as bottleneck,null::numeric as supplyplan, ssp.orderconfirmed::numeric / wm.count::numeric as orderconfirmed,null::numeric as orderunconfirmed,null::numeric as forecast,-1 * (ssp.orderconfirmed::numeric / wm.count::numeric) as diff,o.officersebname" & _
                     " FROM ssp" & _
                     " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                     " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                     " Left join cmmf c on c.cmmf = cr.cmmf " & _
                     " left join sspmarket sm on sm.marketid = ssp.marketid " & _
                     " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                     " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                     " LEFT JOIN officerseb o on o.ofsebid = v.ofsebid" & _
                     " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                     " left join sspweekly wl on wl.yearweek = ssp.week" & _
                     " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                     " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                     " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                     " where ssp.orderconfirmed > 0 and period = " & Myform1.combobox1 - i & _
                     " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(182 - (i * 7))) & _
                     " order by upper(sopdescription))"
            stringbuilder1.Append(sqlstr)
            stringbuilder1.Append(" union all ")
        Next


        For i = 0 To 0


            'Query OrderUnconfirmed
            sqlstr = "(SELECT ssp.period, 'Order unconfirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,cr.cmmf,c.materialdesc,sm.market,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                     " case when wl.label::integer > 9 and wl.label::integer < " & myWeek - i & " then wl.label " & _
                     " when wl.label::integer >= " & myWeek - i & " then '  ' || wl.label  " & _
                     " else ' ' || wl.label " & _
                     " end as datalabel2, " & _
                     " ssp.orderunconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,null::numeric as bottleneck,null::numeric as supplyplan,null::numeric as orderconfirmed,ssp.orderunconfirmed::numeric / wm.count::numeric as orderunconfirmed,null::numeric as forecast,-1 * (ssp.orderunconfirmed::numeric / wm.count::numeric) as diff,o.officersebname" & _
                     " FROM ssp" & _
                     " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                     " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                     " Left join cmmf c on c.cmmf = cr.cmmf " & _
                     " left join sspmarket sm on sm.marketid = ssp.marketid " & _
                     " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                     " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                     " LEFT JOIN officerseb o on o.ofsebid = v.ofsebid" & _
                     " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                     " left join sspweekly wl on wl.yearweek = ssp.week" & _
                     " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                     " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                     " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                     " where ssp.orderunconfirmed > 0 and period = " & Myform1.combobox1 - i & _
                     " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(182 - (i * 7))) & _
                     " order by upper(sopdescription))"
            stringbuilder1.Append(sqlstr)
            stringbuilder1.Append(" union all ")
        Next

        For i = 0 To 0
            'Query ForecastEstimation
            sqlstr = "(SELECT ssp.period, 'Forecast'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription, cr.cmmf,c.materialdesc,sm.market, ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                     " case when wl.label::integer > 9 and wl.label::integer < " & myWeek - i & " then wl.label " & _
                     " when wl.label::integer >= " & myWeek - i & " then '  ' || wl.label  " & _
                     " else ' ' || wl.label " & _
                     " end as datalabel2, " & _
                     " ssp.forecast::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,null::numeric as bottleneck,null::numeric as supplyplan,null::numeric as orderconfirmed,null::numeric as orderunconfirmed,ssp.forecast::numeric / wm.count::numeric as forecast,-1 * (ssp.forecast::numeric / wm.count::numeric) as diff,o.officersebname" & _
                     " FROM ssp" & _
                     " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                     " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                     " Left join cmmf c on c.cmmf = cr.cmmf " & _
                     " left join sspmarket sm on sm.marketid = ssp.marketid " & _
                     " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                     " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                     " LEFT JOIN officerseb o on o.ofsebid = v.ofsebid" & _
                     " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                     " left join sspweekly wl on wl.yearweek = ssp.week" & _
                     " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                     " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                     " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                     " where ssp.forecast > 0 and period = " & Myform1.combobox1 - i & _
                     " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(182 - (i * 7))) & _
                     " order by upper(sopdescription))"
            stringbuilder1.Append(sqlstr)
            'If i <> 1 Then
            '    stringbuilder1.Append(" union all ")
            'End If
        Next

        FinalQuery.Append("select * from (")
        FinalQuery.Append(stringbuilder1.ToString)
        FinalQuery.Append(") as foo ")
        FinalQuery.Append("left join " & _
                          " (select vendorcode,sopfamily,monthly, case when (validnum(sum(bottleneck)) -  (validnum(sum(orderconfirmed)) + validnum(sum(orderunconfirmed)) + validnum(sum(forecast)))) < 0 then 1 else 0 end  as filter " & _
                          " from (")
        FinalQuery.Append(stringbuilder1.ToString)
        FinalQuery.Append(") AS foo1 " & _
                          " group by vendorcode,sopfamily,monthly) foo2 on foo2.vendorcode = foo.vendorcode and foo2.sopfamily = foo.sopfamily and foo2.monthly = foo.monthly")

        'ExcelStuff.FillDataSource(owb, isheet, stringbuilder1.ToString, dbtools1)
        ExcelStuff.FillDataSource(owb, isheet, FinalQuery.ToString, dbtools1)
    End Sub

    Private Sub GeneratePivotTable(ByVal oWb As Excel.Workbook, ByVal iSheet As Integer)
        Dim osheet As Excel.Worksheet
        Try

            osheet = oWb.Worksheets(iSheet)
            oWb.Worksheets(iSheet).select()

            'oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion12)
            osheet.Cells(1, 1) = "as of " & Today.ToString("MMM dd yyyy")
            osheet.PivotTables("PivotTable1").columngrand = False
            osheet.PivotTables("PivotTable1").rowgrand = False
            osheet.PivotTables("PivotTable1").AllowMultipleFilters = True
            osheet.PivotTables("PivotTable1").ingriddropzones = True
            'osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            'Compact mode
            osheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
            osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight8"

            'Calculated Field if any
            osheet.PivotTables("PivotTable1").CalculatedFields.Add("Flagbottleneck", "=IF(filter>0,1,0)", True)


            'add PageField
            osheet.PivotTables("PivotTable1").PivotFields("officersebname").orientation = Excel.XlPivotFieldOrientation.xlPageField
            osheet.PivotTables("PivotTable1").PivotFields("officersebname").currentpage = "All"
            osheet.PivotTables("PivotTable1").PivotFields("officersebname").caption = "SP"


            'add Rowfields
            osheet.PivotTables("PivotTable1").PivotFields("vendorname").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlRowField
            



            'add columnfield
            'osheet.PivotTables("PivotTable1").PivotFields("weekperiod").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            osheet.PivotTables("PivotTable1").PivotFields("monthly").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            osheet.PivotTables("PivotTable1").PivotFields("monthly").numberformat = "MMM-yy"
            osheet.PivotTables("PivotTable1").PivotFields("monthly").Caption = "Month"

            'remove subtotal
            osheet.PivotTables("PivotTable1").pivotfields("vendorname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("sopdescription").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}


            'add datafield
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("diff"), "BottleNeck < TTL Demand", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("orderconfirmed"), "Order Confirmed", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("orderunconfirmed"), "Order Unconfirmed", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("bottleneck"), " Bottleneck", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("Flagbottleneck"), "Flag for demand exceeds bottleneck", Excel.XlConsolidationFunction.xlSum)

            osheet.PivotTables("PivotTable1").PivotFields("BottleNeck < TTL Demand").NumberFormat = "#,##0_);[Red](#,##0)"
            osheet.PivotTables("PivotTable1").PivotFields(" Bottleneck").NumberFormat = "#,##0"
            osheet.PivotTables("PivotTable1").PivotFields(" Forecast").NumberFormat = "#,##0"
            osheet.PivotTables("PivotTable1").PivotFields("Order Confirmed").NumberFormat = "#,##0"
            osheet.PivotTables("PivotTable1").PivotFields("Order Unconfirmed").NumberFormat = "#,##0"


            'Change DataPivot Orientation
            osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField

            'Filter Datafield
            osheet.PivotTables("PivotTable1").PivotFields("sopdescription").PivotFilters.Add(Excel.XlPivotFilterType.xlValueIsGreaterThan, osheet.PivotTables("PivotTable1").PivotFields("Flag for demand exceeds bottleneck"), 0)
            'sort column period 

            'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            osheet.Name = "PivotTable"
            osheet.Cells.EntireColumn.AutoFit()

            osheet.Cells(8, 2).Select()

            iSheet += 1
            
        Catch ex As Exception

        End Try

    End Sub



End Class

