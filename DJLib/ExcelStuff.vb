Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools

Public Class ExcelStuff
    Public Shared Sub ExportToExcel(ByRef FileName As String, ByVal DataSet1 As DataSet)
        Dim Result As Boolean = False

        Dim StringBuilder1 As New System.Text.StringBuilder

        Dim source As String = ""
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then


            Application.DoEvents()
            Cursor.Current = Cursors.WaitCursor
            source = DirectoryBrowser.SelectedPath & "\" & FileName

            'Excel Variable
            Dim oXl As Excel.Application = Nothing
            Dim oWb As Excel.Workbook = Nothing
            Dim oSheet As Excel.Worksheet = Nothing
            Dim SheetName As String = vbEmpty
            Dim oRange As Excel.Range

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
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\RawDataTemplate.xltx")
                'Loop for chart
                oSheet = oWb.Worksheets(1)

                'save excel
                If Not ConvertDataForExcel(StringBuilder1, DataSet1.Tables(0).DefaultView) Then
                    Return
                End If
                'put other info
                Dim abc As String = StringBuilder1.ToString

                Clipboard.SetDataObject(StringBuilder1.ToString, False)

                'oRange = oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1))
                oRange = oSheet.Range("A1")
                oRange.Select()
                oSheet.Paste()
                oRange = oSheet.Range("1:1")
                'oRange.AutoFilter()
                oSheet.Cells.EntireColumn.AutoFit()
                FileName = ValidateFileName(DirectoryBrowser.SelectedPath, source)
                oWb.SaveAs(FileName)
                Result = True
                'FormMenu.setBubbleMessage("Export To Excel", "Done")
            Catch ex As Exception
                'MsgBox(ex.Message)

            Finally
                'clear excel from memory
                oXl.Quit()
                oXl.Visible = True
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
                If Result Then
                    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                        Process.Start(FileName)
                    End If
                End If
                Cursor.Current = Cursors.Default

            End Try

        End If
    End Sub

    Public Shared Sub ExportToExcel(ByRef FileName As String, ByVal DataSetCollection As Collection, ByVal TableIndex As Integer)
        Dim Result As Boolean = False

        Dim StringBuilder1 As System.Text.StringBuilder

        Dim source As String = ""
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then

            Application.DoEvents()
            Cursor.Current = Cursors.WaitCursor
            source = DirectoryBrowser.SelectedPath & "\" & FileName

            'Excel Variable
            Dim oXl As Excel.Application = Nothing
            Dim oWb As Excel.Workbook = Nothing
            Dim oSheet As Excel.Worksheet = Nothing
            Dim SheetName As String = vbEmpty
            Dim oRange As Excel.Range

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
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\RawDataTemplate.xltx")
                'Loop for chart
                Dim dataset1 As DataSet
                'prepare worksheets
                Dim Cursor As System.Windows.Forms.Cursor = Cursors.WaitCursor
                For i = 3 To DataSetCollection.Count - 1
                    oWb.Worksheets.Add(After:=oWb.Worksheets(i))
                Next
                For i = 0 To DataSetCollection.Count - 1
                    oWb.Worksheets(i + 1).Select()
                    oSheet = oWb.Worksheets(1 + i)                    
                    'save excel
                    dataset1 = New DataSet
                    dataset1 = DataSetCollection(1 + i)
                    oSheet.Name = dataset1.Tables(0).TableName
                    StringBuilder1 = New System.Text.StringBuilder

                    If Not ConvertDataForExcel(StringBuilder1, dataset1.Tables(TableIndex).DefaultView) Then
                        Return
                    End If
                    'put other info
                    Dim abc As String = StringBuilder1.ToString

                    Clipboard.SetDataObject(StringBuilder1.ToString, False)

                    'oRange = oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1))

                    oRange = oSheet.Range("A1")
                    oRange.Select()
                    oSheet.Paste()
                    oRange = oSheet.Range("1:1")
                    oSheet.Cells.EntireColumn.AutoFit()
                Next
                oWb.Worksheets(1).select()
                FileName = ValidateFileName(DirectoryBrowser.SelectedPath, source)
                oWb.SaveAs(FileName)
                Result = True
                'FormMenu.setBubbleMessage("Export To Excel", "Done")
            Catch ex As Exception
                'MsgBox(ex.Message)

            Finally
                'clear excel from memory
                oXl.Quit()
                oXl.Visible = True
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
                If Result Then
                    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                        Process.Start(FileName)
                    End If
                End If
                Cursor.Current = Cursors.Default

            End Try

        End If
    End Sub

    Public Shared Function ConvertDataForExcel(ByRef StringBuilderData As System.Text.StringBuilder, ByVal DataView As DataView) As Boolean
        Dim myReturn As Boolean = False
        Dim DataTable As DataTable = DataView.ToTable
        Try
            'Add header
            For i = 0 To DataTable.Columns.Count - 1
                StringBuilderData.Append(DataTable.Columns(i).ToString)
                StringBuilderData.Append(vbTab)
            Next

            StringBuilderData.Append(vbCrLf)
            'Add Detail
            For Each dr In DataTable.Rows

                For i As Long = 0 To dr.itemarray.length - 1
                    '
                    ' Convert the data and fill the string. Null values become blanks.
                    '
                    If dr.itemarray(i).ToString Is DBNull.Value Then
                        StringBuilderData.Append("")
                    Else
                        StringBuilderData.Append(dr.itemarray(i).ToString)
                    End If
                    StringBuilderData.Append(vbTab)
                Next
                '
                ' Add a line feed to the end of each row.
                '
                StringBuilderData.Append(vbCrLf)
                Application.DoEvents()
            Next
            myReturn = True
        Catch ex As Exception
            ' Display an error message.
        End Try
        Return myReturn
    End Function

    Public Shared Sub ExportToExcelAskDirectory(ByRef FileName As String, ByVal DataSet1 As DataSet)
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            If ExportToExcelFullPath(DirectoryBrowser.SelectedPath & "\" & FileName, DataSet1) Then
                If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(FileName)
                End If
            End If
        End If
    End Sub

    Public Shared Sub ExportToExcelAskDirectory(ByRef FileName As String, ByVal Sqlstr As String, ByVal dbTools As Dbtools, Optional ByVal Location As String = "A1", Optional ByVal Template As String = "\templates\RawDataTemplate.xltx", Optional ByVal CreationDate As Boolean = False, Optional ByVal SheetNum As Integer = 1)
        'ask export location
        'MessageBox.Show("ExportToExcelAskDirectory")
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim myFilename As String = DirectoryBrowser.SelectedPath & "\" & FileName
            'If ExportToExcelFullPath(myFilename, Sqlstr, dbTools, "A4", "\templates\TATemplate", True) Then
            'MessageBox.Show(String.Format("ExportToExcelAskDirectory {0}", FileName))
            If ExportToExcelFullPath(myFilename, Sqlstr, dbTools, Location, Template, CreationDate, SheetNum) Then
                If MsgBox("File name: " & myFilename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(myFilename)
                End If
            End If
        End If
    End Sub
    Public Shared Function ExportToExcelFullPath(ByRef Filename As String, ByVal Dataset1 As DataSet) As Boolean
        Dim result As Boolean = False
        Application.DoEvents()
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim oRange As Excel.Range

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
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\RawDataTemplate.xltx")
            'Loop for chart
            oSheet = oWb.Worksheets(1)

            'save excel
            If Not ConvertDataForExcel(StringBuilder1, Dataset1.Tables(0).DefaultView) Then
                Return result
            End If
            'put other info
            Dim abc As String = StringBuilder1.ToString

            Clipboard.SetDataObject(StringBuilder1.ToString, False)

            'oRange = oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1))
            oRange = oSheet.Range("A1")
            oRange.Select()
            oSheet.Paste()
            oRange = oSheet.Range("1:1")
            'oRange.AutoFilter()
            oSheet.Cells.EntireColumn.AutoFit()
            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            oWb.SaveAs(Filename)
            result = True
            'FormMenu.setBubbleMessage("Export To Excel", "Done")
        Catch ex As Exception
            'MsgBox(ex.Message)

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
        Return result
    End Function


    Public Shared Function ExportToExcelFullPath(ByRef Filename As String, ByVal Sqlstr As String, ByVal dbTools As Dbtools, Optional ByVal Location As String = "A1", Optional ByVal Template As String = "\templates\RawDataTemplate.xltx", Optional ByVal CreationDate As Boolean = False, Optional ByVal SheetNum As Integer = 1) As Boolean
        'MessageBox.Show(String.Format("ExportToExcelFullPath "))
        Dim result As Boolean = False
        Application.DoEvents()
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        'Dim oRange As Excel.Range

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
            oWb = oXl.Workbooks.Open(Application.StartupPath & Template)
            'Loop for chart
            oSheet = oWb.Worksheets(SheetNum)
            If CreationDate Then
                oSheet.Cells(1, 1) = "Updated: " & Format(DateTime.Now, "dd MMM yyyy")
            End If
            'MessageBox.Show(String.Format("FillDataSource"))
            FillDataSource(oWb, SheetNum, Sqlstr, dbTools, Location)

            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            oWb.SaveAs(Filename)
            result = True

        Catch ex As Exception
            'MsgBox(ex.Message)

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
        Return result
    End Function

    Public Shared Sub FillDataSource(ByRef owb As Excel.Workbook, ByVal SheetNum As Integer, ByVal sqlstr As String, ByVal dbtools As Dbtools, Optional ByVal Location As String = "A1")

        owb.Worksheets(SheetNum).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(SheetNum)
        Dim oRange As Excel.Range
        ' Dim oExCon As String = "ODBC;DRIVER={PostgreSQL ANSI};" & My.Settings.connectionstring1 & ";UID=" & dbtools.Userid & ";PWD=" & dbtools.password & ";SSLmode=disable;ReadOnly=1;" & _
        '"Protocol=7.4;FakeOidIndex=0;ShowOidColumn=0;RowVersioning=0;ShowSystemTables=0;ConnSettings=;Fetch=100;Socket=4096;UnknownSizes=0;" & _
        '"MaxVarcharSize=255;MaxLongVarcharSize=8190;Debug=0;CommLog=0;Optimizer=1;Ksqo=1;UseDeclareFetch=0;TextAsLongVarchar=1;" & _
        '"UnknownsAsLongVarchar=0;BoolsAsChar=1;Parse=0;CancelAsFreeStmt=0;ExtraSysTablePrefixes=dd_;LFConversion=1;UpdatableCursors=1;" & _
        '"DisallowPremature=0;TrueIsMinus1=0;BI=0;ByteaAsLongVarBinary=0;UseServerSidePrepare=0;LowerCaseIdentifier=0;XaOpt=1"

        Dim oExCon As String = My.Settings.oExCon.tostring
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbtools.Userid & ";PWD=" & dbtools.Password)
        'MessageBox.Show(String.Format("FillDataSource DJLib {0}", oExCon))
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), osheet.Range(Location))
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With
        If owb.Connections.Count > 0 Then
            owb.Connections(owb.Connections.Count).Delete()
        End If

        'oRange = osheet.Range("1:1")
        oRange = osheet.Range(Location)
        oRange.Select()
        osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()

    End Sub
    Public Shared Function ValidateSheetName(ByRef SheetName As String)
        Dim ListChar As String
        Dim TmpList As Object
        Dim i As Integer
        ListChar = "\,?,[,],*,/"
        TmpList = Split(ListChar, ",")
        For i = 0 To UBound(TmpList)
            SheetName = Replace(SheetName, TmpList(i), " ")
        Next
        ValidateSheetName = SheetName
    End Function
End Class
