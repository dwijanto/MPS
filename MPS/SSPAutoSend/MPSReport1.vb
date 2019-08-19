Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports System.ComponentModel
Imports Microsoft.Office.Interop


Public Class MPSReport1

    Public Class MyForm
        Public combobox1 As String
    End Class

    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools("user", "user")
    Dim status As Boolean = False
    Dim FileName As String = String.Empty
    Dim Myform1 As New MyForm
    Dim errMsg As String = String.Empty
    Public Sub Start()

        Dim sqlstr As String = "select period from ssp group by period order by period desc limit 1 "
        Dim datatable As New DataTable
        datatable = dbtools1.getData(sqlstr)
        Myform1.combobox1 = datatable.Rows(0).Item(0).ToString
        Dim DataReader As NpgsqlDataReader = Nothing
        sqlstr = "select paramname,cvalue from sspparamdt where paramhdid = 1"
        Dim mymessage As String = String.Empty

        Using conn As New NpgsqlConnection(dbtools1.getConnectionString)
            Try
                conn.Open()
                Dim command As New NpgsqlCommand
                command.Connection = conn
                command.CommandText = sqlstr
                command.CommandType = CommandType.Text
                DataReader = command.ExecuteReader
                While DataReader.HasRows
                    DataReader.Read()
                    Select Case DataReader.Item(1)
                        Case "1"
                            status = GenerateExcel(DataReader.Item(0) & "\MPSReport1Auto.xlsx", errMsg)
                    End Select

                End While
            Catch ex As NpgsqlException

            End Try
        End Using

    End Sub


    Private Function GenerateExcel(ByRef FileName As String, ByRef errormsg As String) As Boolean

        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim dataset1 As New DataSet

        Dim StopWatch As New Stopwatch
        StopWatch.Start()

        'Open Excel

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
            'Application.DoEvents()
            oXl.Visible = True
            'get process pid
            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
                'Application.DoEvents()
            Next
            oXl.Visible = False
            oXl.DisplayAlerts = False
            'backgroundworker1.ReportProgress(1, "Opening Template...")
            oWb = oXl.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory & "templates\ExcelTemplate.xltx")
            Dim iSheetDAta As Integer = 2
            'Loop for chart
            'Go to worksheetData
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Worksheets(iSheetDAta).select()
            'backgroundworker1.ReportProgress(1, "DB Query...")
            Call QueryDataAll(oWb, iSheetDAta)


            oWb.Worksheets(iSheetDAta).select()
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Names.Add(Name:="DBRangeAll", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "DBAll"

            'Generate Chart&Pivot start from worksheet 2
            iSheetDAta = 1
            'backgroundworker1.ReportProgress(1, "Generating PivotTable...")
            GeneratePivotTable(oWb, iSheetDAta)
            StopWatch.Stop()
            'backgroundworker1.ReportProgress(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            'FileName = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            'backgroundworker1.ReportProgress(1, "Saving File...")
            oWb.SaveAs(FileName)

            result = True
        Catch ex As Exception
            errormsg = ex.Message
        Finally
            'backgroundworker1.ReportProgress(1, "Releasing Memory...")
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
        End Try
        Return result

    End Function

    Private Sub QueryDataAll(ByRef owb As Excel.Workbook, ByVal isheet As Integer)
        Dim sqlstr As String = String.Empty
        Dim stringbuilder1 As New System.Text.StringBuilder

        Dim mycriteria As String = String.Empty

        'Check Worksheet
        For i = owb.Worksheets.Count To isheet - 1
            owb.Worksheets.Add(After:=owb.Worksheets(i))
        Next

        'GET MPS DATA
        sqlstr = "(SELECT f.period, io.typeofinfo, f.vendorcode, v.vendorname, fg.sopfamilygroup, sf.sopfamily, sf.sopdescription, d.startingdate, d.startingdate AS monthly, wl.yearweek as datalabel1,wl.label as datalabel2, d.datavalue::numeric AS datavalue, dl.weekperiod, dl.dailydate, mwd.count, d.datavalue::numeric / mwd.count::numeric AS dailyvalue" & _
                 " FROM sspftycap f" & _
                 " LEFT JOIN sspftycapdata d ON d.ftycapid = f.ftycapid" & _
                 " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = f.sopfamilyid" & _
                 " LEFT JOIN vendor v ON v.vendorcode = f.vendorcode" & _
                 " LEFT JOIN ssptypeofinfo io ON io.ssptypeofinfoid = f.typeofinfoid" & _
                 " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = f.sopfamilyid" & _
                 " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                 " LEFT JOIN sspdaily dl ON dl.monthperiod = d.startingdate" & _
                 " left join sspweekly wl on wl.yearweek = dl.weekperiod" & _
                 " left join sspmonthlywdparam mwd on mwd.monthperiod = dl.monthperiod" & _
                 " where period = " & Myform1.combobox1 & " and not d.ftycapdataid is null " & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday" & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")

        'Query Orderconfirmed
        sqlstr = "(SELECT ssp.period, 'Order confirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1,wl.label as datalabel2, ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,ssp.orderconfirmed::numeric / wm.count::numeric as dailyvalue" & _
                 " FROM ssp" & _
                 " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                 " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                 " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                 " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                 " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                 " left join sspweekly wl on wl.yearweek = ssp.week" & _
                 " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                 " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                 " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                 " where ssp.orderconfirmed > 0 and period = " & Myform1.combobox1 & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday" & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")

        'Query OrderUnconfirmed
        sqlstr = "(SELECT ssp.period, 'Order unconfirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1,wl.label as datalabel2, ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,ssp.orderunconfirmed::numeric / wm.count::numeric as dailyvalue" & _
                 " FROM ssp" & _
                 " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                 " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                 " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                 " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                 " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                 " left join sspweekly wl on wl.yearweek = ssp.week" & _
                 " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                 " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                 " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                 " where ssp.orderunconfirmed > 0 and period = " & Myform1.combobox1 & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday" & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")

        'Query ForecastEstimation
        sqlstr = "(SELECT ssp.period, 'Forecast'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1,wl.label as datalabel2, ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,ssp.forecast::numeric / wm.count::numeric as dailyvalue" & _
                 " FROM ssp" & _
                 " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                 " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
                 " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
                 " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
                 " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
                 " left join sspweekly wl on wl.yearweek = ssp.week" & _
                 " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
                 " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
                 " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                 " where ssp.forecast > 0 and period = " & Myform1.combobox1 & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday" & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)

        ExcelStuff.FillDataSource(owb, isheet, stringbuilder1.ToString, dbtools1)
    End Sub

    Private Sub GeneratePivotTable(ByVal oWb As Excel.Workbook, ByVal iSheet As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(iSheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable1").columngrand = False
        osheet.PivotTables("PivotTable1").rowgrand = False
        osheet.PivotTables("PivotTable1").ingriddropzones = True
        osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").currentpage = "Monthly"


        'add Rowfields
        osheet.PivotTables("PivotTable1").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlRowField

        'remove subtotal

        'add columnfield
        osheet.PivotTables("PivotTable1").PivotFields("weekperiod").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("datalabel2").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'add datafield
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Data Value").NumberFormat = "0"
        osheet.PivotTables("PivotTable1").PivotFields("sopdescription").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").pivotfields("sopdescription").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").pivotfields("vendorname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").pivotfields("typeofinfo").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").pivotfields("weekperiod").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").pivotfields("datalabel2").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'sort column period
        'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.Cells.EntireColumn.AutoFit()

    End Sub

End Class

