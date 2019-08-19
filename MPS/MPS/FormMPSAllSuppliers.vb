Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class FormMPSAllSuppliers

    Public Class MyForm
        Public Property combobox1 As String
    End Class

    Dim myform1 As MyForm
    Dim Dataset1 As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Private WithEvents backgroundworker1 As New BackgroundWorker
    Dim status As Boolean = False
    Dim FileName As String = String.Empty
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Check Selected CheckedListbox
        If ComboBox1.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox1.Select()
            Exit Sub
        End If

        Button1.Enabled = False


        If Not (backgroundworker1.IsBusy) Then

            'Dim FileName As String = String.Empty
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                FileName = DirectoryBrowser.SelectedPath & "\" & "WeeklyMPSAllVendor-" & ComboBox1.Text & "-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
                Myform1 = New MyForm With {.combobox1 = ComboBox1.SelectedItem.ToString}              
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

    Private Sub FormMPSAllSuppliers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim sqlstr As String = "select period from ssp group by period order by period desc;"
        dbtools1.FillCombobox(ComboBox1, sqlstr)
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
            'Call QueryDataAll(oWb, iSheetDAta)


            oWb.Worksheets(iSheetDAta).select()
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Names.Add(Name:="DBRangeAll", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "DBAll"

            'Generate Chart&Pivot start from worksheet 2
            iSheetDAta = 1
            backgroundworker1.ReportProgress(1, "Generating PivotTable...")
            ' Call GeneratePivotTable(oWb, iSheetDAta)
            StopWatch.Stop()
            backgroundworker1.ReportProgress(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            FileName = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            backgroundworker1.ReportProgress(1, "Saving File...")
            oWb.SaveAs(FileName)

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
End Class