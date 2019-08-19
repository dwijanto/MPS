Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports Microsoft.Office.Interop

Public Class Report1
    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'init Combobox
        Dim sqlstr As String = "select period from ssp group by period order by period desc;"
        dbtools1.FillCombobox(ComboBox1, sqlstr)
        dbtools1.FillCombobox(ComboBox2, sqlstr)
    End Sub
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

        Button2.Enabled = False
        Dim result As Boolean = False
        Dim dataset1 As New DataSet
        Dim FileName As String = String.Empty
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"

        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim stopwatch As New Stopwatch
            stopwatch.Start()

            FileName = DirectoryBrowser.SelectedPath & "\" & "SSPComparison-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
            'Open Excel
            Cursor.Current = Cursors.WaitCursor
            Dim source As String = FileName
            Dim StringBuilder1 As New System.Text.StringBuilder

            'Excel Variable
            Dim oXl As Excel.Application = Nothing
            Dim oWb As Excel.Workbook = Nothing
            Dim oSheet As Excel.Worksheet = Nothing
            Dim SheetName As String = vbEmpty


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
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")
                'Loop for chart
                'Go to worksheetData
                oSheet = oWb.Worksheets(2)

                'Get Data passing sqlstr,worksheet for data
                oWb.Worksheets(2).select()

                Dim firstVar As Integer = Math.Max(CInt(ComboBox1.Text), CInt(ComboBox2.Text))
                Dim LastVar As Integer = Math.Min(CInt(ComboBox1.Text), CInt(ComboBox2.Text))



                Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from sopall" & _
                                        " where period = " & firstVar
                StringBuilder1.Append(sqlstr)
                StringBuilder1.Append(" Union all ")

                sqlstr = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast * -1 as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from sopall" & _
                                        " where period = " & LastVar
                StringBuilder1.Append(sqlstr)

                ExcelStuff.FillDataSource(oWb, 2, StringBuilder1.ToString, dbtools1)


                'set DbRange
                oWb.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET('" & oSheet.Name & "'!R1C1,0,0,COUNTA('" & oSheet.Name & "'!C1),COUNTA('" & oSheet.Name & "'!R1))")

                'Go To Worksheet(1)
                oSheet = oWb.Worksheets(1)
                oWb.Worksheets(1).select()

                oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
                oSheet.PivotTables("PivotTable1").columngrand = False
                oSheet.PivotTables("PivotTable1").rowgrand = False
                oSheet.PivotTables("PivotTable1").ingriddropzones = True
                oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

                'add Rowfields
                oSheet.PivotTables("PivotTable1").PivotFields("Sop Family").orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable1").PivotFields("Range").orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable1").PivotFields("Period").orientation = Excel.XlPivotFieldOrientation.xlRowField

                'remove subtotal
                oSheet.PivotTables("PivotTable1").pivotfields("Sop Family").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable1").pivotfields("Range").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable1").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

                'add columnfield
                oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField

                'add datafield
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), "Sum of Forecast", Excel.XlConsolidationFunction.xlSum)

                'sort column period
                oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
                oSheet.Cells.EntireColumn.AutoFit()
                FileName = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
                'oWb.Connections("Connection").Delete()
                stopwatch.Stop()
                Label3.Text = "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds
                oWb.SaveAs(FileName)

                result = True
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
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
        End If

        If result Then
            If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(FileName)
            End If
        End If
        Button2.Enabled = True

    End Sub
   


    Private Sub Report1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class