Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports SSP.PublicClass

Public Class Report1
    Public Property department As Department
    Public Property TableName As String
    Public Property ViewName As String

    Private WithEvents bgworker As New BackgroundWorker
    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    'Dim myHashtable As New Hashtable
    Dim myArrayList As New ArrayList
    Dim period1 As String
    Dim period2 As String
    Dim status As Boolean = False
    Dim Filename As String
    Dim SelectedDir As String


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        
    End Sub
    Private Sub Report1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'init Combobox
        Application.DoEvents()
        'tablename = ssp
        Dim sqlstr As String = "select period from " & TableName & " group by period order by period desc;"
        dbtools1.FillCombobox(ComboBox1, sqlstr)
        dbtools1.FillCombobox(ComboBox2, sqlstr)
        sqlstr = "select 0 as vendorcode, 'SELECT ALL' as vendorname union all " & _
                 " Select 1 as vendorcode,'ALL VENDOR' as vendorname  union All" & _
                 " (select foo.vendorcode as vendorcode, v.vendorname as vendorname from (" & _
                 " select distinct ssp.vendorcode from " & TableName & " as ssp) as foo" & _
                 " left join vendor v on v.vendorcode = foo.vendorcode " & _
                 " order by v.vendorname)"
        dbtools1.FillCheckedListBoxDataSource(CheckedListBox1, sqlstr)
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

        period1 = ComboBox1.Text
        period2 = ComboBox2.Text
        Button2.Enabled = False
        myArrayList.Clear()
        For i = 1 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemCheckState(i) Then
                Dim dr = DirectCast(CheckedListBox1.Items(i), DataRowView)
                myArrayList.Add(New cblList With {.id = dr.Item(0),
                                                  .name = dr.Item(1)})
            End If
        Next
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
            Filename = SelectedDir & "\" & "SSPComparison-" & DirectCast(myArrayList.Item(i), cblList).name & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
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

            Dim firstVar As Integer = Math.Max(CInt(period1), CInt(period2))
            Dim LastVar As Integer = Math.Min(CInt(period1), CInt(period2))

            Dim check = DirectCast(myArrayList.Item(list), cblList)
            Dim vendorfilter As String = String.Empty
            If check.id <> 1 Then
                vendorfilter = " and vendorcode = " & check.id
            End If
            'viewname = sopall
            'Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                        " where period = " & firstVar & vendorfilter
            Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" ,  cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"",  week as ""Week"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", unit as ""Unit"" from " & ViewName & _
                                    " where period = " & firstVar & vendorfilter
            StringBuilder1.Append(sqlstr)
            StringBuilder1.Append(" Union all ")

            ' sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * -1 as ""OrderUnConfirmed"", orderconfirmed * -1 as ""OrderConfirmed"", forecast * -1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                        " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "
            'sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * 1 as ""OrderUnConfirmed"", orderconfirmed * 1 as ""OrderConfirmed"", forecast * 1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                      " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "
            sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", week as ""Week"", orderconfirmed * 1 as ""OrderConfirmed"", forecast * 1 as ""Forecast"", unit as ""Unit"" from " & ViewName & _
                                 " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "
            StringBuilder1.Append(sqlstr)
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
            'oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)
            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Requirement", "=OrderConfirmed +Forecast", True)

            'add PageField
            oSheet.PivotTables("PivotTable1").PivotFields("Market").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            If check.id <> 1 Then
                oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").CurrentPage = check.name
            End If
            'add Rowfields
            oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField

            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
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
            'oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField

            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
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

            Dim firstVar As Integer = Math.Max(CInt(period1), CInt(period2))
            Dim LastVar As Integer = Math.Min(CInt(period1), CInt(period2))

            Dim check = DirectCast(myArrayList.Item(list), cblList)
            Dim vendorfilter As String = String.Empty
            If check.id <> 1 Then
                vendorfilter = " and vendorcode = " & check.id
            End If
            'viewname = sopall
            Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
                                    " where period = " & firstVar & vendorfilter
            StringBuilder1.Append(sqlstr)
            StringBuilder1.Append(" Union all ")

            'sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * -1 as ""OrderUnConfirmed"", orderconfirmed * -1 as ""OrderConfirmed"", forecast * -1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
            '                        " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "

            sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed * 1 as ""OrderUnConfirmed"", orderconfirmed * 1 as ""OrderConfirmed"", forecast * 1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from " & ViewName & _
                                   " where period >= " & LastVar & " and period <= " & firstVar - 1 & vendorfilter & " order by ""Period"" desc "


            StringBuilder1.Append(sqlstr)
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
            oSheet.PivotTables("PivotTable1").rowgrand = False
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
            'oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

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
            'oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("CMMF").SubtotalName = "? Variance"
            'remove subtotal
            oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

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


End Class

Public Class cblList
    Public Property id As Integer
    Public Property name As String
End Class