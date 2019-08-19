Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports Npgsql
Imports System
Imports Microsoft.Office.Interop
Imports System.ComponentModel
Public Class FormWeeklyMPS
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim StartingDate As Date
    Dim FirstVar As Integer
    Private WithEvents backgroundworker1 As New BackgroundWorker
    Dim listSelected() As String = Nothing
    Dim Status As Boolean = False
    Dim Filename As String = String.Empty
    Dim MyVendorCode As String
    Dim Monthlychart1 As Boolean
    Dim Weeklychart1 As Boolean
    Dim SSPPivot As Boolean
    Dim IncludeSupplyPlan As Boolean
    Dim IncludeBottleNeck As Boolean
    Dim IncludeSupplyPlanIF As Boolean
    Dim IncludeBottleNeckIF As Boolean
    Dim IncludeSupplyPlanGRP As Boolean
    Dim IncludeBottleNeckGRP As Boolean

    Dim mylabel As String = String.Empty
    Dim my18Week As Date
    Dim allfamiliesMonthHT As Hashtable
    Dim allfamiliesWeekHT As Hashtable
    Dim SelectedWeek As Integer
    Private Sub FormWeeklyMPS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'get list of combobox
        Dim sqlstr As String = "select period from ssp group by period order by period desc;"
        dbtools1.FillCombobox(ComboBox1, sqlstr)
        'ListBox1.DataSource = {"All Families", "Individual Family", "Group Family"}

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker1.DoWork
        'BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Start")

        Dim errMsg As String = String.Empty
        'Dim Filename As String = String.Empty

        Status = GenerateExcel(Filename, errMsg)
        If Status Then
            backgroundworker1.ReportProgress(1, TextBox1.Text & " Done.")
        Else
            backgroundworker1.ReportProgress(1, "Error::" & errMsg)
        End If

        If Status Then
            If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Filename)
            End If
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundworker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 1
                TextBox1.Text = e.UserState
            Case 2
                'TextBox2.Text = e.UserState
            Case 3
                'TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker1.RunWorkerCompleted
        FormMenu.setBubbleMessage("Export To Excel", "Done")
        If Status Then
            'If CheckBox1.Checked Then
            '    Me.Close()
            'End If
        End If
        Button1.Enabled = True
    End Sub
    Public Function GenerateExcel(ByRef Filename As String, ByRef errmsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim dataset1 As New DataSet
        Dim stopwatch As New Stopwatch
        stopwatch.Start()


        'Open Excel
        Application.DoEvents()

        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
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
            For i = 3 To 10
                oWb.Worksheets.Add(After:=oWb.Worksheets(i))
            Next
            'Loop for chart
            'Go to worksheetData
            oSheet = oWb.Worksheets(1)

            'Get Data SSP DATa 
            oWb.Worksheets(1).select()

            'FirstVar = CInt(ComboBox1.Text)
            Dim sqlstr As String = "SELECT period as ""Period"", sopfamily as ""Sop Family"" , sopdescription as ""Description"",range as ""Range"", cmmf as ""CMMF"",  materialdesc as ""Material Description"", vendorcode as ""Vendor Code"", vendorname as ""Vendor Name"", market as ""Market"", startingdate::date as ""Starting Date"", periodofetd as ""Period of ETD"", week as ""Week"", orderunconfirmed as ""OrderUnConfirmed"", orderconfirmed as ""OrderConfirmed"", forecast as ""Forecast"", totalamount as ""Total Amount"", unit as ""Unit"", crcycode as ""CrcyCode"" from sopall" & _
                                    " where period = " & FirstVar & " and vendorcode = " & MyVendorCode

            'sqlstr = "select ssp.period,fg.sopfamilygroup,ssf.sopfamily, case when  ssf.sopdescription is null then 'OTHER FAMILIES' else ssf.sopdescription  end as sopdescription,sr.range,scr.cmmf,c.materialdesc as ""Material Description"",ssp.vendorcode,v.vendorname,sm.market,ssp.startingdate,ssp.periodofetd,ssp.week,ssp.orderunconfirmed,ssp.orderconfirmed,ssp.forecast,ssp.orderunconfirmed + ssp.orderconfirmed + ssp.forecast as total ,ssp.totalamount," & _
            '" case when (substring(ssp.week::text,5,2)::integer > 9) then ' ' || substring(ssp.week::text,5,2) else (substring(ssp.week::text,5,2)::integer)::text end as weeklable " & _
            '" from ssp " & _
            '" LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
            '" LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
            '" LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
            '" LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
            '" Left join cmmf c on c.cmmf = scr.cmmf" & _
            '" Left join sspmarket sm on sm.marketid = ssp.marketid" & _
            '" Left join vendor v on v.vendorcode = ssp.vendorcode" & _
            '" left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = scs.sopfamilyid" & _
            '" left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
            '" where period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & myCriteria & _
            '" order by upper(sopdescription)"
            Dim myWeek As String = FirstVar.ToString.Substring(4, 2)
            sqlstr = "select ssp.period,fg.sopfamilygroup,ssf.sopfamily, case when  ssf.sopdescription is null then 'OTHER FAMILIES' else ssf.sopdescription  end as sopdescription,sr.range,scr.cmmf,c.materialdesc as ""Material Description"",ssp.vendorcode,v.vendorname,sm.market,ssp.startingdate,ssp.periodofetd,ssp.week,ssp.orderunconfirmed,ssp.orderconfirmed,ssp.forecast,ssp.orderunconfirmed + ssp.orderconfirmed + ssp.forecast as total ,ssp.totalamount," & _
            " case when substring(ssp.week::text,5,2)::integer > 9 and substring(ssp.week::text,5,2)::integer < " & myWeek & " then substring(ssp.week::text,5,2) " & _
            " when substring(ssp.week::text,5,2)::integer >= " & myWeek & " then '  ' || substring(ssp.week::text,5,2)  " & _
            " else ' ' || (substring(ssp.week::text,5,2)::integer)::text " & _
            " end as weeklabel " & _
            "  from ssp " & _
            " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
            " LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
            " LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
            " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
            " Left join cmmf c on c.cmmf = scr.cmmf" & _
            " Left join sspmarket sm on sm.marketid = ssp.marketid" & _
            " Left join vendor v on v.vendorcode = ssp.vendorcode" & _
            " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = scs.sopfamilyid" & _
            " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
            " where period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & myCriteria & _
            " order by upper(sopdescription)"

            StringBuilder1.Append(sqlstr)

            Dim isheet As Integer = 1
            ExcelStuff.FillDataSource(oWb, isheet, StringBuilder1.ToString, dbtools1)
            'getDBRange1
            oWb.Names.Add(Name:="DBRangeSSP", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "SSP"

            Dim iSheetData = 2 'After SSP
            Dim IndividualFamilyTable As New DataTable
            Dim GroupTable As New DataTable

            Dim CallQueryData As Boolean = False



            'Get SheetLocation For DB
            For i = 0 To listSelected.Count - 1
                Select Case listSelected(i)
                    Case "All Families"
                        iSheetData += 1
                    Case "Individual Family"
                        'CallQueryData = True
                        IndividualFamilyTable = getindividualFamily()
                        iSheetData += IndividualFamilyTable.Rows.Count
                    Case "Group Family"
                        'CallQueryData = True
                        GroupTable = getGroup()
                        iSheetData += GroupTable.Rows.Count
                End Select
            Next
            'If Not CallQueryData Then
            If Monthlychart1 Or Weeklychart1 Then
                Call QueryDataAll(oWb, iSheetData)
            End If
            'End If


            oWb.Worksheets(iSheetData).select()
            oSheet = oWb.Worksheets(iSheetData)
            oWb.Names.Add(Name:="DBRangeAll", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "DBAll"

            'Generate Chart&Pivot start from worksheet 2
            isheet = 2
            For i = 0 To listSelected.Count - 1
                Select Case listSelected(i)
                    Case "All Families"
                        'Create Sheet All Families  
                        CreateWorksheetAllFamilies(oWb, isheet)
                        iSheetData += 1
                        oSheet = oWb.Worksheets(isheet)
                        SheetName = "All Families"
                        oSheet.Name = SheetName
                        isheet += 1
                    Case "Individual Family"
                        Dim mycount As Integer = 0
                        For Each DataRow As DataRow In IndividualFamilyTable.Rows
                            'Create Worksheet
                            Try
                                'TextBox1.Text = "Working on " & DataRow.Item(1).ToString & " " & mycount + 1 & " of " & IndividualFamilyTable.Rows.Count
                                backgroundworker1.ReportProgress(1, "Working on " & DataRow.Item(1).ToString & " " & mycount + 1 & " of " & IndividualFamilyTable.Rows.Count)
                                Application.DoEvents()
                                mylabel = "Working on " & DataRow.Item(1).ToString & " " & mycount + 1 & " of " & IndividualFamilyTable.Rows.Count

                                oSheet = oWb.Worksheets(isheet)
                                SheetName = DataRow.Item(1).ToString
                                If SheetName.Length = 0 Then
                                    'fill data from ssp where sopfamily is blank
                                    sqlstr = "select ssp.period,fg.sopfamilygroup,ssf.sopfamily, case when  ssf.sopdescription is null then 'OTHER FAMILIES' else ssf.sopdescription  end as sopdescription,sr.range,scr.cmmf,c.materialdesc as ""Material Description"",ssp.vendorcode,v.vendorname,sm.market,ssp.startingdate,ssp.periodofetd,ssp.week,ssp.orderunconfirmed,ssp.orderconfirmed,ssp.forecast,ssp.orderunconfirmed + ssp.orderconfirmed + ssp.forecast as total ,ssp.totalamount" & _
                                             " from ssp " & _
                                             " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                                             " LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
                                             " LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
                                             " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
                                             " Left join cmmf c on c.cmmf = scr.cmmf" & _
                                             " Left join sspmarket sm on sm.marketid = ssp.marketid" & _
                                             " Left join vendor v on v.vendorcode = ssp.vendorcode" & _
                                             " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = scs.sopfamilyid" & _
                                             " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                                             " where period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & myCriteria & _
                                             " and sopdescription is null order by upper(sopdescription)"
                                    ExcelStuff.FillDataSource(oWb, isheet, sqlstr, dbtools1)
                                Else
                                    CreateWorksheetIndividual(oWb, isheet, DataRow, mycount)
                                End If

                                'oSheet = oWb.Worksheets(isheet)

                                If SheetName.Length > 30 Then
                                    SheetName = SheetName.Substring(1, 30)
                                ElseIf SheetName.Length = 0 Then
                                    SheetName = "OTHER FAMILIES"
                                End If
                                oSheet.Name = ExcelStuff.ValidateSheetName(SheetName)

                                isheet += 1
                                mycount += 1
                            Catch ex As Exception
                                errmsg = ex.Message
                            End Try
                            Application.DoEvents()
                        Next

                    Case "Group Family"
                        Dim mycount As Integer = 0
                        For Each DataRow As DataRow In GroupTable.Rows
                            Try
                                'TextBox1.Text = "Working on " & DataRow.Item(0).ToString & " " & mycount + 1 & " of " & GroupTable.Rows.Count
                                backgroundworker1.ReportProgress(1, "Working on " & DataRow.Item(0).ToString & " " & mycount + 1 & " of " & GroupTable.Rows.Count)
                                mylabel = "Working on " & DataRow.Item(0).ToString & " " & mycount + 1 & " of " & GroupTable.Rows.Count
                                CreateWorksheetWorksheetGroup(oWb, isheet, DataRow, mycount)
                                oSheet = oWb.Worksheets(isheet)
                                SheetName = DataRow.Item(0).ToString
                                If SheetName.Length > 30 Then
                                    SheetName = SheetName.Substring(1, 30)
                                ElseIf SheetName.Length = 0 Then
                                    SheetName = "(blank)"
                                End If
                                oSheet.Name = ExcelStuff.ValidateSheetName(SheetName)

                                isheet += 1
                                mycount += 1
                            Catch ex As Exception
                                errmsg = ex.Message
                            End Try
                        Next

                End Select
            Next
            oXl.DisplayAlerts = False
            'oWb.Worksheets("DBAll").delete()
            oXl.DisplayAlerts = True
            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            stopwatch.Stop()
            mylabel = mylabel & " Elapsed Time:" & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString
            backgroundworker1.ReportProgress(1, mylabel)
            oWb.SaveAs(Filename)

            result = True
        Catch ex As Exception
            errmsg = ex.Message
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


        'If result Then
        '    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
        '        Process.Start(FileName)
        '    End If
        'End If
        'Button1.Enabled = True
        Return result
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Check Selected CheckedListbox
        Dim period As Integer = CInt(StartingDate.Year.ToString & Format(StartingDate.Month, "00"))

        If ComboBox1.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox1.Select()
            Exit Sub
        End If

        getSelectedItems(CheckedListBox1, "", ",", FirstItemAll:=False, ListSelected:=listSelected)

        If IsNothing(listSelected) Then
            MsgBox("Please select from CheckedListbox!")
            CheckedListBox1.Select()
            Exit Sub
        End If
        Button1.Enabled = False
        SelectedWeek = ComboBox1.Text.Substring(4, 2)

        If Not (backgroundworker1.IsBusy) Then


            'Dim FileName As String = String.Empty
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                Filename = DirectoryBrowser.SelectedPath & "\" & "MPS-DV-" & ComboBox1.Text & "-" & ComboBox2.SelectedValue.ToString & ".xlsx"
                FirstVar = CInt(ComboBox1.Text)
                MyVendorCode = ComboBox2.SelectedValue.ToString
                Monthlychart1 = MonthlyChart.Checked
                Weeklychart1 = WeeklyChart.Checked
                SSPPivot = SSPPivotTable.Checked
                IncludeBottleNeck = SeriesBottleNeck.Checked
                IncludeSupplyPlan = SeriesSupplyPlan.Checked
                IncludeBottleNeckIF = SeriesBottleNeckIF.Checked
                IncludeSupplyPlanIF = SeriesSupplyPlanIF.Checked
                IncludeBottleNeckGRP = SeriesBottleNeckGRP.Checked
                IncludeSupplyPlanGRP = SeriesSupplyPlanGRP.Checked

                Try
                    backgroundworker1.WorkerReportsProgress = True
                    backgroundworker1.WorkerSupportsCancellation = True
                    backgroundworker1.RunWorkerAsync()
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
            End If
            Button1.Enabled = True
        Else
            MsgBox("Please wait until the current process is finished")
        End If
        'Play AVI
        'MsgBox("Play Avi")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sqlstr As String = String.Empty
        If ComboBox1.Text = "" Then
            'CheckedListBox1.DataSource = Nothing
        Else
            'get month
            StartingDate = getStartingDate(ComboBox1.Text)
            sqlstr = "select v.vendorcode,v.vendorname from (select ssp.vendorcode from ssp " & _
                                   " where(SSP.period = " & ComboBox1.Text & ")" & _
                                   " group by ssp.vendorcode) as foo " & _
                                   " left join vendor v on v.vendorcode = foo.vendorcode order by vendorname"

            Try
                Cursor.Current = Cursors.WaitCursor
                dbtools1.FillComboboxDataSource(ComboBox2, sqlstr)

            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedValue.ToString <> "System.Data.DataRowView" Then
            'Group check
            Dim sqlstr As String = " select ssf.sspsopfamilyid,ssf.sopdescription from ssp" & _
                     " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                     " LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
                     " LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
                     " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
                     " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = scs.sopfamilyid" & _
                     " right join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                     " where(SSP.vendorcode = " & ComboBox2.SelectedValue.ToString & ")" & _
                     " group by ssf.sspsopfamilyid,ssf.sopdescription order by upper(sopdescription)"

            CheckedListBox1.Items.Clear()
            CheckedListBox1.Items.Add("All Families")
            CheckedListBox1.Items.Add("Individual Family")
            If getNumberRecord(sqlstr) > 0 Then
                CheckedListBox1.Items.Add("Group Family")
            End If


            Try
                Cursor.Current = Cursors.WaitCursor
                'dbtools1.FillCheckedListBoxDataSource(CheckedListBox2, sqlstr)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default

        End If
    End Sub

    Private Function getSelectedItems(ByVal CLB As CheckedListBox, ByVal Fieldname As String, ByVal JoinText As String, Optional ByVal DataTypeString As Boolean = True, Optional ByVal FirstItemAll As Boolean = True, Optional ByRef ListSelected() As String = Nothing) As String
        Dim mylist(0) As String

        Dim myindex As Integer = 0
        Dim Istart As Integer
        Dim myreturn As String = String.Empty
        'check for 'ALL'
        If FirstItemAll Then
            If CLB.GetItemChecked(0) = True Then
                Return myreturn
            Else
                Istart = 1
            End If
        Else
            Istart = 0
        End If

        For i = Istart To CLB.Items.Count - 1
            If CLB.GetItemChecked(i) = True Then
                ReDim Preserve mylist(myindex)
                ReDim Preserve ListSelected(myindex)
                CLB.SetSelected(i, True)
                ListSelected(myindex) = CLB.SelectedItem.ToString
                If DataTypeString Then
                    mylist(myindex) = Fieldname & "='" & CLB.SelectedValue & "'"
                Else
                    mylist(myindex) = Fieldname & "=" & CLB.SelectedValue
                End If
                myindex += 1
            End If
        Next
        myreturn = Join(mylist, JoinText)
        If myreturn <> "" Then
            myreturn = "(" & myreturn & ")"
        End If
        Return myreturn
    End Function

#Region "Hide"

    'Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
    '    If ComboBox2.DisplayMember <> "" Then

    '        Dim sqlstr As String = "select 0 as sspsopfamilyid, 'All' as sopdescription union all (select ssf.sspsopfamilyid,ssf.sopdescription from ssp" & _
    '                 " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
    '                 " LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
    '                 " LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
    '                 " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
    '                 " where(SSP.vendorcode = " & ComboBox2.SelectedValue.ToString & ")" & _
    '                 " group by ssf.sspsopfamilyid,ssf.sopdescription order by upper(sopdescription))"
    '        Try
    '            Cursor.Current = Cursors.WaitCursor
    '            dbtools1.FillCheckedListBoxDataSource(CheckedListBox2, sqlstr)
    '        Catch ex As Exception
    '        End Try
    '        Cursor.Current = Cursors.Default

    '    End If
    'End Sub

    'Private Sub CheckedListBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If ComboBox1.Text <> "" And ComboBox1.Text <> "System.Data.DataRowView" Then
    '        Select Case sender.selectedindex
    '            Case 0
    '                Dim chkstate As CheckState
    '                chkstate = sender.GetItemCheckState(0)
    '                For i = 0 To sender.items.count - 1
    '                    sender.SetItemChecked(i, chkstate)
    '                Next
    '            Case Else
    '                sender.SetItemChecked(0, 0)
    '                If sender.Items.Count = Countlist(sender) + 1 Then
    '                    sender.SetItemChecked(0, True)
    '                End If
    '        End Select
    '    End If
    'End Sub

    'Public Function Countlist(ByVal sender As Object) As Integer
    '    Dim count As Integer = 0
    '    For i = 0 To sender.Items.Count - 1
    '        If sender.GetItemCheckState(i) Then
    '            count += 1
    '        End If
    '    Next
    '    Return count
    'End Function

    'Private Function getSelectedItems(ByVal CLB As CheckedListBox, ByVal Fieldname As String, ByVal JoinText As String, Optional ByVal DataTypeString As Boolean = True, Optional ByVal FirstItemAll As Boolean = True) As String
    '    Dim mylist(0) As String
    '    Dim myindex As Integer = 0
    '    Dim Istart As Integer
    '    Dim myreturn As String = String.Empty
    '    'check for 'ALL'
    '    If FirstItemAll Then
    '        If CLB.GetItemChecked(0) = True Then
    '            Return myreturn
    '        Else
    '            Istart = 1
    '        End If
    '    Else
    '        Istart = 0
    '    End If

    '    For i = Istart To CheckedListBox2.Items.Count - 1
    '        If CLB.GetItemChecked(i) = True Then
    '            ReDim Preserve mylist(myindex)
    '            CLB.SetSelected(i, True)
    '            If DataTypeString Then
    '                mylist(myindex) = Fieldname & "='" & CLB.SelectedValue & "'"
    '            Else
    '                mylist(myindex) = Fieldname & "=" & CLB.SelectedValue
    '            End If
    '            myindex += 1
    '        End If
    '    Next
    '    myreturn = Join(mylist, JoinText)
    '    If myreturn <> "" Then
    '        myreturn = "(" & myreturn & ")"
    '    End If
    '    Return myreturn
    'End Function
#End Region
    Private Function getStartingDate(ByVal YearWeek As Integer) As Date
        Dim sqlstr As String = "select startdate from weektomonth where yearweek = " & YearWeek
        Dim mydate As Date
        Using conn As New NpgsqlConnection(dbtools1.getConnectionString)
            conn.Open()
            Dim command As New NpgsqlCommand(sqlstr, conn)

            mydate = command.ExecuteScalar
        End Using
        Return (mydate)
    End Function
    Private Function getNumberRecord(ByVal sqlstr As String) As Integer
        Dim myret As Integer
        Try


            Using conn As New NpgsqlConnection(dbtools1.getConnectionString)
                conn.Open()
                Dim command As New NpgsqlCommand(sqlstr, conn)
                myret = command.ExecuteScalar
            End Using

        Catch ex As Exception

        End Try
        Return myret
    End Function
#Region "Hide2"


    Private Sub temp()
        'isheet = 1
        'oSheet = oWb.Worksheets(isheet)
        'oWb.Worksheets(isheet).select()
        ''set DbRange
        'oWb.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET('" & oSheet.Name & "'!R1C1,0,0,COUNTA('" & oSheet.Name & "'!C1),COUNTA('" & oSheet.Name & "'!R1))")

        ''Go To Worksheet()
        'isheet -= 1
        'oSheet = oWb.Worksheets(isheet)
        'oWb.Worksheets(isheet).select()


        'oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        'oSheet.PivotTables("PivotTable1").columngrand = False
        'oSheet.PivotTables("PivotTable1").rowgrand = False
        'oSheet.PivotTables("PivotTable1").ingriddropzones = True
        'oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

        ''Calculated Field
        ''oSheet.PivotTables("PivotTable1").CalculatedFields.Add(" Total", "=OrderUnConfirmed+OrderConfirmed +Forecast", True)


        ''add Rowfields
        'oSheet.PivotTables("PivotTable1").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField


        ''remove subtotal
        'oSheet.PivotTables("PivotTable1").pivotfields("sopdescription").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").PivotFields("Material Description").LayoutBlankLine = True
        'oSheet.PivotTables("PivotTable1").PivotFields("CMMF").LayoutBlankLine = True
        'oSheet.PivotTables("PivotTable1").PivotFields("sopdescription").LayoutBlankLine = True
        ''add columnfield
        'oSheet.PivotTables("PivotTable1").PivotFields("Week").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        ''add datafield
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("OrderUnConfirmed"), " OrderUnConfirmed", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("OrderConfirmed"), " OrderConfirmed", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Total"), "            Total", Excel.XlConsolidationFunction.xlSum)


        ''sort column period
        ''oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        'oSheet.Cells.EntireColumn.AutoFit()




    End Sub
#End Region


    Public Sub QueryDataAll(ByRef owb As Excel.Workbook, ByVal isheet As Integer)
        Dim sqlstr As String = String.Empty
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim period As Integer = CInt(StartingDate.Year.ToString & Format(StartingDate.Month, "00"))
        Dim mycriteria As String = String.Empty
        Dim myWeek As String = FirstVar.ToString.Substring(4, 2)
        'Check Worksheet
        For i = owb.Worksheets.Count To isheet - 1
            owb.Worksheets.Add(After:=owb.Worksheets(i))
        Next

        'GET MPS DATA
        sqlstr = "(SELECT f.period, io.typeofinfo, f.vendorcode, v.vendorname, fg.sopfamilygroup, sf.sopfamily, sf.sopdescription, d.startingdate, d.startingdate AS monthly, wl.yearweek as datalabel1, " & _
                 " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
                 " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
                 " else ' ' || wl.label " & _
                 " end as datalabel2, " & _
                 "  d.datavalue::numeric AS datavalue, dl.weekperiod, dl.dailydate, mwd.count, d.datavalue::numeric / mwd.count::numeric AS dailyvalue,1 as uom" & _
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
                 " where period = " & period & " and not d.ftycapdataid is null  and v.vendorcode = " & MyVendorCode & mycriteria & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not dl.isholiday and dl.dailydate >= " & DateFormatyyyyMMdd(StartingDate) & " and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")

        'Query Orderconfirmed
        'sqlstr = "(SELECT ssp.period, 'Order confirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
        '         " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
        '         " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
        '         " else ' ' || wl.label " & _
        '         " end as datalabel2, " & _
        '         " ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.orderconfirmed::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end )/ wm.count::numeric  as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
        '         " FROM ssp" & _
        '         " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
        '         " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
        '         " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
        '         " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
        '         " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
        '         " left join sspweekly wl on wl.yearweek = ssp.week" & _
        '         " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
        '         " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
        '         " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
        '         " left join cmmf c on c.cmmf = cr.cmmf" & _
        '         " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        '         " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
        '         " order by upper(sopdescription))"

        sqlstr = "(SELECT ssp.period, 'Order confirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                 " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
                 " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
                 " else ' ' || wl.label " & _
                 " end as datalabel2, " & _
                 " ssp.orderconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.orderconfirmed::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end )/ wm.count::numeric  as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
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
                 " left join cmmf c on c.cmmf = cr.cmmf" & _
                 " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
                 " order by upper(sopdescription))"

        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")
        '" where ssp.orderconfirmed > 0 and period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        'Query OrderUnconfirmed
        'sqlstr = "(SELECT ssp.period, 'Order unconfirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
        '         " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
        '         " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
        '         " else ' ' || wl.label " & _
        '         " end as datalabel2, " & _
        '         " ssp.orderunconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.orderunconfirmed::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end ) / wm.count::numeric as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
        '         " FROM ssp" & _
        '         " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
        '         " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
        '         " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
        '         " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
        '         " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
        '         " left join sspweekly wl on wl.yearweek = ssp.week" & _
        '         " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
        '         " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
        '         " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
        '         " left join cmmf c on c.cmmf = cr.cmmf" & _
        '         " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        '         " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
        '         " order by upper(sopdescription))"

        sqlstr = "(SELECT ssp.period, 'Order unconfirmed'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                 " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
                 " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
                 " else ' ' || wl.label " & _
                 " end as datalabel2, " & _
                 " ssp.orderunconfirmed::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.orderunconfirmed::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end ) / wm.count::numeric as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
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
                 " left join cmmf c on c.cmmf = cr.cmmf" & _
                 " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday ) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        stringbuilder1.Append(" union all ")
        '" where ssp.orderunconfirmed > 0 and period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        'Query ForecastEstimation
        'sqlstr = "(SELECT ssp.period, 'Forecast'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
        '         " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
        '         " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
        '         " else ' ' || wl.label " & _
        '         " end as datalabel2, " & _
        '         " ssp.forecast::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.forecast::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end ) / wm.count::numeric as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
        '         " FROM ssp" & _
        '         " LEFT JOIN sspcmmfrange cr ON cr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
        '         " LEFT JOIN sspcmmfsop cs ON cs.cmmf = cr.cmmf" & _
        '         " LEFT JOIN sspsopfamilies sf ON sf.sspsopfamilyid = cs.sopfamilyid" & _
        '         " LEFT JOIN vendor v ON v.vendorcode = ssp.vendorcode" & _
        '         " LEFT JOIN sspdaily dl ON dl.weekperiod = ssp.week" & _
        '         " left join sspweekly wl on wl.yearweek = ssp.week" & _
        '         " left join sspweeklywdparam wm on wm.weekperiod = ssp.week" & _
        '         " left join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = cs.sopfamilyid" & _
        '         " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
        '         " left join cmmf c on c.cmmf = cr.cmmf" & _
        '         " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        '         " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday and wl.crossmonth) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
        '         " order by upper(sopdescription))"
        sqlstr = "(SELECT ssp.period, 'Forecast'::text AS typeofinfo, ssp.vendorcode, v.vendorname,fg.sopfamilygroup, sf.sopfamily, sf.sopdescription,  ssp.startingdate, dl.monthperiod AS monthly, ssp.week AS datalabel1," & _
                 " case when wl.label::integer > 9 and wl.label::integer < " & myWeek & " then wl.label " & _
                 " when wl.label::integer >= " & myWeek & " then '  ' || wl.label  " & _
                 " else ' ' || wl.label " & _
                 " end as datalabel2, " & _
                 " ssp.forecast::integer AS datavalue ,dl.weekperiod,dl.dailydate,wm.count,(ssp.forecast::numeric * case when c.uom isnull then 1::numeric else c.uom::numeric end ) / wm.count::numeric as dailyvalue,case when c.uom isnull then 1 else c.uom end as uom" & _
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
                 " left join cmmf c on c.cmmf = cr.cmmf" & _
                 " where  period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
                 " and date_part('dow'::text, dl.dailydate) <> 0 AND not (dl.isholiday ) and dl.dailydate < " & DateFormatyyyyMMdd(StartingDate.AddDays(126)) & _
                 " order by upper(sopdescription))"
        stringbuilder1.Append(sqlstr)
        '" where ssp.forecast > 0 and period = " & FirstVar & " and v.vendorcode = " & MyVendorCode & mycriteria & _
        ExcelStuff.FillDataSource(owb, isheet, stringbuilder1.ToString, dbtools1)
    End Sub


    Private Function getindividualFamily() As DataTable
        Dim myRet As New DataTable
        Dim sqlstr As String = "select ssf.sspsopfamilyid,ssf.sopdescription from ssp" & _
                      " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                          " LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid" & _
                          " LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf" & _
                          " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid" & _
                          " where(SSP.vendorcode = " & MyVendorCode & ")" & _
                          " group by ssf.sspsopfamilyid,ssf.sopdescription order by upper(sopdescription)"
        myRet = dbtools1.getData(sqlstr)

        Return myRet
    End Function

    Private Function getGroup() As DataTable
        Dim myRet As DataTable
        Dim sqlstr As String = "select fg.sopfamilygroup from ssp" & _
                 " LEFT JOIN sspcmmfrange scr ON scr.sspcmmfrangeid = ssp.sspcmmfrangeid LEFT JOIN ssprange sr ON sr.rangeid = scr.rangeid LEFT JOIN sspcmmfsop scs ON scs.cmmf = scr.cmmf " & _
                 " LEFT JOIN sspsopfamilies ssf ON ssf.sspsopfamilyid = scs.sopfamilyid " & _
                 " right join sspsopfamilygrouptx fgtx on fgtx.sspsopfamilyid = scs.sopfamilyid" & _
                 " left join sspsopfamilygroup fg on fg.sspsopfamilygroupid = fgtx.sspsopfamilygroupid" & _
                 " where(SSP.vendorcode = " & MyVendorCode & ")" & _
                 " group by fg.sopfamilygroup order by upper(fg.sopfamilygroup)"
        myRet = dbtools1.getData(sqlstr)
        Return myRet
    End Function

#Region "All Families"
    Private Sub CreateWorksheetAllFamilies(ByVal oWb As Excel.Workbook, ByVal isheet As Integer)
        Dim lastposition As Integer = 0
        Dim mylabel As String = "Working on "
        If Monthlychart1 Then
            'Create PivotTable Monthly
            backgroundworker1.ReportProgress(1, mylabel & "All Families Monthly Chart ")
            'TextBox1.Text = mylabel & "All Families Monthly Chart "
            Application.DoEvents()
            createPivotTableMonthly(oWb, isheet, lastposition)


        End If
        If Weeklychart1 Then
            'Create PivotTable Weekly
            'TextBox1.Text = mylabel & "All Families Weekly Chart "
            backgroundworker1.ReportProgress(1, mylabel & "All Families Weekly Chart ")
            Application.DoEvents()
            createPivotTableWeekly(oWb, isheet, lastposition)
            'Create ChartWeekly
            If IncludeBottleNeck Then
                'include series
            End If
            If IncludeSupplyPlan Then
                'include series
            End If
        End If

        If SSPPivot Then
            'TextBox1.Text = mylabel & "All Families SSP PivotTable"
            backgroundworker1.ReportProgress(1, mylabel & "All Families SSP PivotTable")
            Application.DoEvents()
            createPivotSSP(oWb, isheet, lastposition)
            'Create PivotTable SSP
        End If
    End Sub

    Private Sub createPivotTableMonthly(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef LastPosition As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C52", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable1").columngrand = False
        osheet.PivotTables("PivotTable1").rowgrand = False
        osheet.PivotTables("PivotTable1").ingriddropzones = True
        osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable1").NullString = "0"

        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").currentpage = "Monthly"


        'add Rowfields
        osheet.PivotTables("PivotTable1").PivotFields("monthly").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("monthly").numberformat = "MMM-yyy"

        'remove subtotal

        'add columnfield
        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Data Value").NumberFormat = "0"
        'sort column period
        'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")

        'allfamiliesMonthHT = New Hashtable
        'Dim i As Integer = 0
        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()

        Call createMonthlyChart(oWb, osheet, LastPosition)




    End Sub

    Private Sub createPivotTableWeekly(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef LastPosition As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C60", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable2").columngrand = False
        osheet.PivotTables("PivotTable2").rowgrand = False
        osheet.PivotTables("PivotTable2").ingriddropzones = True
        osheet.PivotTables("PivotTable2").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable2").NullString = "0"
        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").currentpage = "Weekly"

        'add Rowfields
        'osheet.PivotTables("PivotTable2").PivotFields("datalabel1").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").orientation = Excel.XlPivotFieldOrientation.xlRowField

        'remove subtotal
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}



        'add columnfield
        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)

        'sort column period
        'oSheet.PivotTables("PivotTable2").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.Cells.EntireColumn.AutoFit()
        'oWb.Names.Add(Name:="Label" & osheet.Name, RefersToR1C1:="=OFFSET(" & osheet.Name & "!R8C61,0,0,COUNTA(" & osheet.Name & "!C61)-1,1")
        osheet.PivotTables("PivotTable2").PivotFields(" Data Value").NumberFormat = "0"

        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()

        createWeeklyChart(oWb, osheet, LastPosition)


    End Sub

    Private Sub createPivotSSP(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal LastPosition As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()
        Dim myPosition() As Integer = {6, 30, 50}
        Dim myindex As Integer = 0
        Select Case LastPosition
            Case 0
                myindex = 0
            Case 310
                myindex = 1
            Case Else
                myindex = 2
        End Select
        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeSSP").CreatePivotTable(osheet.Name & "!R" & myPosition(myindex) & "C1", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable3").columngrand = True
        osheet.PivotTables("PivotTable3").rowgrand = False
        osheet.PivotTables("PivotTable3").ingriddropzones = True
        osheet.PivotTables("PivotTable3").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        'osheet.PivotTables("PivotTable3").NullString = "0"
        'Calculated Field if any

        'add PageField


        'add Rowfields
        osheet.PivotTables("PivotTable3").PivotFields("cmmf").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField


        'remove subtotal
        osheet.PivotTables("PivotTable3").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").LayoutBlankLine = True
        osheet.PivotTables("PivotTable3").PivotFields("CMMF").LayoutBlankLine = True
        'add columnfield
        osheet.PivotTables("PivotTable3").PivotFields("Weeklabel").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable3").PivotFields("Weeklabel").caption = " Week"
        ''add datafield
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderConfirmed"), " OrderConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderUnConfirmed"), " OrderUnConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Total"), " Requirement", Excel.XlConsolidationFunction.xlSum)

        'sort column period
        'oSheet.PivotTables("PivotTable3").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.PivotTables("PivotTable3").TableStyle2 = "PivotStyleLight3"
        osheet.PivotTables("PivotTable3").ShowTableStyleRowStripes = True
        osheet.Cells.EntireColumn.AutoFit()
    End Sub


#End Region
#Region "Individual Family"


    Private Sub CreateWorksheetIndividual(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim lastposition As Integer = 0
        'Dim mylabel As String = "" 'TextBox1.Text
        If Monthlychart1 Then
            'Create PivotTable Monthly
            'TextBox1.Text = mylabel & " Monthly Chart "
            backgroundworker1.ReportProgress(1, mylabel & " Monthly Chart ")
            Application.DoEvents()
            createPivotTableMonthlyIndividual(oWb, isheet, lastposition, DataRow, mycount)


        End If
        If Weeklychart1 Then
            'Create PivotTable Weekly
            'TextBox1.Text = mylabel & " Weekly Chart "
            backgroundworker1.ReportProgress(1, mylabel & " Weekly Chart ")
            Application.DoEvents()
            createPivotTableWeeklyIndividual(oWb, isheet, lastposition, DataRow, mycount)
            'Create ChartWeekly

        End If

        If SSPPivot Then
            'TextBox1.Text = mylabel & " SSP PivotTable "
            backgroundworker1.ReportProgress(1, mylabel & " SSP PivotTable ")
            Application.DoEvents()
            createPivotSSPIndividual(oWb, isheet, lastposition, DataRow, mycount)
            'Create PivotTable SSP
        End If

    End Sub

    Private Sub createPivotTableMonthlyIndividual(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef lastposition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C52", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable1").columngrand = False
        osheet.PivotTables("PivotTable1").rowgrand = False
        osheet.PivotTables("PivotTable1").ingriddropzones = True
        osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable1").NullString = "0"
        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").currentpage = "Monthly"
        osheet.PivotTables("PivotTable1").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("sopdescription").currentpage = osheet.PivotTables("PivotTable1").PivotFields("sopdescription").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)


        'add Rowfields
        osheet.PivotTables("PivotTable1").PivotFields("monthly").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("monthly").numberformat = "MMM-yyy"

        'remove subtotal

        'add columnfield
        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)

        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()

        'sort column period
        'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.PivotTables("PivotTable1").PivotFields(" Data Value").NumberFormat = "0"
        osheet.Cells.EntireColumn.AutoFit()

        Dim myChart As Excel.Chart
        Dim myRange As Excel.Range = osheet.Range(osheet.Range("AZ6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked
        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try

        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try

        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        lastposition += 10
        myChart.ChartArea.Top = lastposition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = DataRow.Item(1).ToString & " Monthly"
        lastposition += 300


        Try
            myChart.SeriesCollection(1)
        Catch ex As Exception

        End Try

        'Create ChartMonthly
        If Not (IncludeBottleNeckIF) Then
            'include series
            Try
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Bottleneck").Visible = False
            Catch ex As Exception
            End Try

        End If
        If Not (IncludeSupplyPlanIF) Then
            'include series
            Try
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Supply Plan").Visible = False
            Catch ex As Exception
            End Try

        End If

    End Sub

    Private Sub createPivotTableWeeklyIndividual(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef lastposition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C60", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable2").columngrand = False
        osheet.PivotTables("PivotTable2").rowgrand = False
        osheet.PivotTables("PivotTable2").ingriddropzones = True
        osheet.PivotTables("PivotTable2").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable2").NullString = "0"
        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").currentpage = "Weekly"
        osheet.PivotTables("PivotTable2").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable2").PivotFields("sopdescription").currentpage = osheet.PivotTables("PivotTable2").PivotFields("sopdescription").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)

        'add Rowfields
        'osheet.PivotTables("PivotTable2").PivotFields("datalabel1").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").orientation = Excel.XlPivotFieldOrientation.xlRowField

        'remove subtotal
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'add columnfield
        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)


        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()


        'sort column period
        'oSheet.PivotTables("PivotTable2").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.Cells.EntireColumn.AutoFit()
        osheet.PivotTables("PivotTable2").PivotFields(" Data Value").NumberFormat = "0"
        Dim myChart As Excel.Chart
        Dim myRange As Excel.Range = osheet.Range(osheet.Range("BH6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked
        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try

        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try

        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        lastposition += 10
        myChart.ChartArea.Top = lastposition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = DataRow.Item(1).ToString & " Weekly"
        lastposition += 300
    End Sub

    Private Sub createPivotSSPIndividual(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef LastPosition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()
        Dim myPosition() As Integer = {6, 30, 50}
        Dim myindex As Integer = 0
        Select Case LastPosition
            Case 0
                myindex = 0
            Case 310
                myindex = 1
            Case Else
                myindex = 2
        End Select
        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeSSP").CreatePivotTable(osheet.Name & "!R" & myPosition(myindex) & "C1", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable3").columngrand = True
        osheet.PivotTables("PivotTable3").rowgrand = False
        osheet.PivotTables("PivotTable3").ingriddropzones = True
        osheet.PivotTables("PivotTable3").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        'osheet.PivotTables("PivotTable3").NullString = "0"
        'Calculated Field if any

        'add PageField


        'add Rowfields
        osheet.PivotTables("PivotTable3").PivotFields("cmmf").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField


        'remove subtotal
        osheet.PivotTables("PivotTable3").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").LayoutBlankLine = True
        osheet.PivotTables("PivotTable3").PivotFields("CMMF").LayoutBlankLine = True
        'add columnfield
        osheet.PivotTables("PivotTable3").PivotFields("weeklabel").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable3").PivotFields("weeklabel").Caption = " Week"
        osheet.PivotTables("PivotTable3").PivotFields("sopdescription").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable3").PivotFields("sopdescription").currentpage = osheet.PivotTables("PivotTable3").PivotFields("sopdescription").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)

        ''add datafield
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderConfirmed"), " OrderConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderUnConfirmed"), " OrderUnConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Total"), " Requirement", Excel.XlConsolidationFunction.xlSum)

        'sort column period
        'oSheet.PivotTables("PivotTable3").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.PivotTables("PivotTable3").TableStyle2 = "PivotStyleLight3"
        osheet.PivotTables("PivotTable3").ShowTableStyleRowStripes = True
        osheet.Cells.EntireColumn.AutoFit()

    End Sub
#End Region

#Region "Group Family"
    Private Sub CreateWorksheetWorksheetGroup(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim lastposition As Integer = 0
        'Dim mylabel As String = "" 'TextBox1.Text
        If Monthlychart1 Then
            'Create PivotTable Monthly
            'TextBox1.Text = mylabel & " Monthly Chart "
            backgroundworker1.ReportProgress(1, mylabel & " Monthly Chart ")
            Application.DoEvents()
            createPivotTableMonthlyGroup(oWb, isheet, lastposition, DataRow, mycount)


        End If
        If Weeklychart1 Then
            'Create PivotTable Weekly
            'TextBox1.Text = mylabel & " Weekly Chart "
            backgroundworker1.ReportProgress(1, mylabel & " Weekly Chart ")
            Application.DoEvents()
            createPivotTableWeeklyGroup(oWb, isheet, lastposition, DataRow, mycount)
            'Create ChartWeekly

        End If

        If SSPPivot Then
            'TextBox1.Text = mylabel & " SSP PivotTable "
            backgroundworker1.ReportProgress(1, mylabel & " SSP PivotTable ")
            Application.DoEvents()
            createPivotSSPGroup(oWb, isheet, lastposition, DataRow, mycount)
            'Create PivotTable SSP
        End If
    End Sub

    Private Sub createPivotTableMonthlyGroup(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef lastposition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C52", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable1").columngrand = False
        osheet.PivotTables("PivotTable1").rowgrand = False
        osheet.PivotTables("PivotTable1").ingriddropzones = True
        osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable1").NullString = "0"
        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable1").PivotFields("periodtype").currentpage = "Monthly"
        osheet.PivotTables("PivotTable1").PivotFields("sopfamilygroup").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("sopfamilygroup").currentpage = osheet.PivotTables("PivotTable1").PivotFields("sopfamilygroup").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)


        'add Rowfields
        osheet.PivotTables("PivotTable1").PivotFields("monthly").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("monthly").numberformat = "MMM-yyy"

        'remove subtotal

        'add columnfield
        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)

        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()

        'sort column period
        'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.PivotTables("PivotTable1").PivotFields(" Data Value").NumberFormat = "0"
        osheet.Cells.EntireColumn.AutoFit()

        Dim myChart As Excel.Chart
        Dim myRange As Excel.Range = osheet.Range(osheet.Range("AZ6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked
        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try

        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try

        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        lastposition += 10
        myChart.ChartArea.Top = lastposition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = DataRow.Item(0).ToString & " Monthly"
        lastposition += 300

        'Create ChartMonthly
        If Not (IncludeBottleNeckGRP) Then
            'include series
            Try
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Bottleneck").Visible = False
            Catch ex As Exception
            End Try

        End If
        If Not (IncludeSupplyPlanGRP) Then
            'include series
            Try
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Supply Plan").Visible = False
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub createPivotTableWeeklyGroup(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef lastposition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()

        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C60", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable2").columngrand = False
        osheet.PivotTables("PivotTable2").rowgrand = False
        osheet.PivotTables("PivotTable2").ingriddropzones = True
        osheet.PivotTables("PivotTable2").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        osheet.PivotTables("PivotTable2").NullString = "0"
        'Calculated Field if any

        'add PageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable2").PivotFields("periodtype").currentpage = "Weekly"
        osheet.PivotTables("PivotTable2").PivotFields("sopfamilygroup").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable2").PivotFields("sopfamilygroup").currentpage = osheet.PivotTables("PivotTable2").PivotFields("sopfamilygroup").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)

        'add Rowfields
        'osheet.PivotTables("PivotTable2").PivotFields("datalabel1").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").orientation = Excel.XlPivotFieldOrientation.xlRowField

        'remove subtotal
        osheet.PivotTables("PivotTable2").PivotFields("datalabel2").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'add columnfield
        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'add datafield
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("dailyvalue"), " Data Value", Excel.XlConsolidationFunction.xlSum)



        osheet.PivotTables("PivotTable2").PivotFields(" Data Value").NumberFormat = "0"


        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order confirmed").Caption = " Order confirmed"
        Catch ex As Exception
        End Try
        Try
            osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").PivotItems("Order unconfirmed").Caption = " Order unconfirmed"
        Catch ex As Exception
        End Try

        osheet.PivotTables("PivotTable2").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlDescending, "typeofinfo")
        osheet.Cells.EntireColumn.AutoFit()
        'sort column period
        'oSheet.PivotTables("PivotTable2").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.Cells.EntireColumn.AutoFit()

        Dim myChart As Excel.Chart
        Dim myRange As Excel.Range = osheet.Range(osheet.Range("BH6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked
        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try

        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try

        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        lastposition += 10
        myChart.ChartArea.Top = lastposition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = DataRow.Item(0).ToString & " Weekly"
        lastposition += 300
    End Sub

    Private Sub createPivotSSPGroup(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByRef lastposition As Integer, ByVal DataRow As DataRow, ByVal mycount As Integer)
        Dim osheet As Excel.Worksheet
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets(isheet).select()
        Dim myPosition() As Integer = {6, 30, 50}
        Dim myindex As Integer = 0
        Select Case lastposition
            Case 0
                myindex = 0
            Case 310
                myindex = 1
            Case Else
                myindex = 2
        End Select
        oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeSSP").CreatePivotTable(osheet.Name & "!R" & myPosition(myindex) & "C1", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable3").columngrand = True
        osheet.PivotTables("PivotTable3").rowgrand = False
        osheet.PivotTables("PivotTable3").ingriddropzones = True
        osheet.PivotTables("PivotTable3").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)
        'osheet.PivotTables("PivotTable3").NullString = "0"
        'Calculated Field if any

        'add PageField


        'add Rowfields
        osheet.PivotTables("PivotTable3").PivotFields("cmmf").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").orientation = Excel.XlPivotFieldOrientation.xlRowField


        'remove subtotal
        osheet.PivotTables("PivotTable3").pivotfields("Material Description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").pivotfields("CMMF").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable3").PivotFields("Material Description").LayoutBlankLine = True
        osheet.PivotTables("PivotTable3").PivotFields("CMMF").LayoutBlankLine = True
        'add columnfield
        osheet.PivotTables("PivotTable3").PivotFields("weeklabel").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable3").PivotFields("weeklabel").Caption = " Week"
        osheet.PivotTables("PivotTable3").PivotFields("sopfamilygroup").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable3").PivotFields("sopfamilygroup").currentpage = osheet.PivotTables("PivotTable3").PivotFields("sopfamilygroup").pivotitems(mycount + 1).name 'IIf(DataRow.Item(1).ToString = "", "(blank)", DataRow.Item(1).ToString)

        ''add datafield
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderConfirmed"), " OrderConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("OrderUnConfirmed"), " OrderUnConfirmed", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("Total"), " Requirement", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").TableStyle2 = "PivotStyleLight3"
        osheet.PivotTables("PivotTable3").ShowTableStyleRowStripes = True
        'sort column period
        'oSheet.PivotTables("PivotTable3").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
        osheet.Cells.EntireColumn.AutoFit()
    End Sub
#End Region

    Private Sub createMonthlyChart(ByVal oWb As Excel.Workbook, ByRef osheet As Excel.Worksheet, ByRef lastposition As Integer)
        Dim myChart As Excel.Chart

        Dim myRange As Excel.Range = osheet.Range(osheet.Range("AZ6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked

        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        'osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").AutoSort(Excel.XlSortOrder.xlAscending, "typeofinfo")

        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try

        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try
        lastposition += 10
        myChart.ChartArea.Top = lastposition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = "ALL Families Monthly"
        lastposition += 300

        'Create ChartMonthly
        If Not (IncludeBottleNeck) Then
            'include series
            Try
                'osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems(allfamiliesMonthHT.Item("Bottleneck")).Visible = False
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Bottleneck").Visible = False
            Catch ex As Exception
            End Try

        End If
        If Not (IncludeSupplyPlan) Then
            'include series
            Try
                osheet.PivotTables("PivotTable1").PivotFields("typeofinfo").PivotItems("Supply Plan").Visible = False
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub createWeeklyChart(ByRef oWb As Excel.Workbook, ByRef osheet As Excel.Worksheet, ByRef LastPosition As Integer)
        Dim myChart As Excel.Chart
        Dim myRange As Excel.Range = osheet.Range(osheet.Range("BH6").CurrentRegion.Address)
        myChart = osheet.Shapes.AddChart.Chart
        myChart.SetSourceData(myRange)
        oWb.ShowPivotChartActiveFields = False
        myChart.ChartType = Excel.XlChartType.xlColumnStacked

        Try
            myChart.SeriesCollection(" Order confirmed").interior.colorindex = 14
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection(" Order unconfirmed").interior.colorindex = 18
        Catch ex As Exception
        End Try
        Try
            myChart.SeriesCollection("Forecast").interior.colorindex = 17
        Catch ex As Exception
        End Try


        Try
            myChart.SeriesCollection("Supply Plan").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Supply Plan").border.weight = Excel.XlBorderWeight.xlThin
            myChart.SeriesCollection("Supply Plan").border.colorindex = 23
        Catch ex As Exception

        End Try
        Try
            myChart.SeriesCollection("Bottleneck").charttype = Excel.XlChartType.xlLine
            myChart.SeriesCollection("Bottleneck").border.colorindex = 3
            myChart.SeriesCollection("Bottleneck").border.weight = Excel.XlBorderWeight.xlThin
        Catch ex As Exception

        End Try
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataTableWithLegendKeys)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
        myChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone)
        LastPosition += 10
        myChart.ChartArea.Top = LastPosition
        myChart.ChartArea.Left = 10
        myChart.ChartArea.Width = 1000
        myChart.ChartArea.Height = 300
        myChart.ChartTitle.Text = "ALL Families Weekly"
        LastPosition += 300
    End Sub


    Private Sub SeriesBottleNeck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SeriesBottleNeck.CheckedChanged

    End Sub
End Class