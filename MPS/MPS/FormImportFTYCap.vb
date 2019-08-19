Imports DJLib
Imports DJLib.Dbtools
Imports System.ComponentModel
Imports System.IO
Imports Npgsql
Imports System.Text
Public Class FormImportFTYCap
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String
    Dim Status As Boolean = False
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim myconverter As New utf8towin1252
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim Dataset1 As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                Try
                    BackgroundWorker1.WorkerReportsProgress = True
                    BackgroundWorker1.WorkerSupportsCancellation = True
                    BackgroundWorker1.RunWorkerAsync()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Start")

        Dim errMsg As String = String.Empty
        Status = ImportData(FileName, errMsg)
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
        Else
            BackgroundWorker1.ReportProgress(2, "Error::" & errMsg)
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 2
                TextBox2.Text = e.UserState
            Case 3
                TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        FormMenu.setBubbleMessage("Import Factory Capacity", "Done")
        If Status Then
            If CheckBox1.Checked Then
                Me.Close()
            End If
        End If
    End Sub

    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim list As New List(Of String)
        Dim stringBuilder1 As New System.Text.StringBuilder
        Dim stringBuilder2 As New System.Text.StringBuilder
        Dim myRecord() As String
        Dim FTYCapId As Long = 0
        Dim RangeId As Long = 0
        Dim SSPCMMFRangeid As Long = 0
        Dim period As String = String.Empty
        Dim myId As Long = 0
        Dim sqlstr As String = String.Empty
        Try
            BackgroundWorker1.ReportProgress(2, "Preparing Data...")
            Dataset1 = New DataSet

            sqlstr = "select sspsopfamilyid,sopfamily from sspsopfamilies;" & _
                     "select ssptypeofinfoid,typeofinfo from ssptypeofinfo;" & _
                     "select periodtypeid,periodtype from sspperiodtype;" & _
                     "select yearweek,mymonth,myyear,startdate,ivalue - case when holidays isnull then 0 else holidays end as workingdays from weektomonth left join paramhd on paramhd.paramname = 'workingdays' order by yearweek;" & _
                     "select mydate,workingdays from sspmonthlywd;"

            dbtools1.getDataSet(sqlstr, Dataset1)
            'Dim keys(0) As DataColumn
            'keys(0) = Dataset1.Tables(0).Columns(1)
            'Dataset1.Tables(0).PrimaryKey = keys

            Dim key2(0) As DataColumn
            key2(0) = Dataset1.Tables(1).Columns(1)
            Dataset1.Tables(1).PrimaryKey = key2
            Dim idstring As String = String.Empty

            Dim Key3(0) As DataColumn
            Key3(0) = Dataset1.Tables(2).Columns(1)
            Dataset1.Tables(2).PrimaryKey = Key3

            Dim Key4(0) As DataColumn
            Key4(0) = Dataset1.Tables(3).Columns(0)
            Dataset1.Tables(3).PrimaryKey = Key4

            Dim Key5(1) As DataColumn
            Key5(0) = Dataset1.Tables(4).Columns(0)
            Dataset1.Tables(4).PrimaryKey = Key5

            Dataset1.Tables(0).TableName = "SSPSOPFamily"
            Dataset1.Tables(1).TableName = "SSPTypeOfInfoId"
            Dataset1.Tables(2).TableName = "SSPPeriodType"
            Dataset1.Tables(3).TableName = "SSPWeeklyWD"
            Dataset1.Tables(4).TableName = "SSPMonthlyWD"

            BackgroundWorker1.ReportProgress(3, "Open File...")
            Try
                Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)
                    Dim line As String = myStream.ReadLine
                    Do While (Not line Is Nothing)
                        list.Add(line)
                        line = myStream.ReadLine
                    Loop
                End Using
            Catch ex As Exception
                errMessage = ex.Message
                Return False
            End Try

            Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
            BackgroundWorker1.ReportProgress(2, "Scanning Data...")
            Dim myhastable As New Hashtable
            Dim cmmfvendorHT As New Hashtable
            Dim mySB As New StringBuilder
            Dim myMonth As String = String.Empty

            myRecord = list(1).Split(vbTab)
            period = myRecord(0)
            'Remove existing period
            BackgroundWorker1.ReportProgress(3, "Clear Data for Period : " & period)
            sqlstr = "Delete from sspftycap where period = " & period
            Dim mymessage As String = String.Empty

            If Not dbtools1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                errMessage = mymessage
                Return False
            End If

            For i = 1 To list.Count - 1
                Try
                    BackgroundWorker1.ReportProgress(3, "Processing line " & i & " of " & list.Count - 1)
                    myRecord = list(i).Split(vbTab)
                    period = myRecord(0)
                    'If i = 1329 Then
                    '    Debug.WriteLine("debug Mode")
                    'End If
                    FTYCapId = getFTYCapId(myRecord, errMessage)
                    If FTYCapId = -1 Then
                        Return False
                    End If

                    'read data monthly
                    For j = 6 To UBound(myRecord) - 1
                        'check for blank value
                        If Not myRecord(j) = "" Then
                            mySB.Append(FTYCapId & vbTab)
                            mySB.Append(1 & vbTab) 'Periodtype 1 for monthly 2 for weekly
                            mySB.Append(getdate(j - 6, period) & vbTab) 'calculation starting from 0 as current month
                            If j = 6 Then
                                myMonth = "M"
                            Else
                                myMonth = "M+" & Format(j - 6, "00")
                            End If
                            mySB.Append(myMonth & vbTab)
                            mySB.Append(myRecord(j) & vbCrLf)


                        End If
                    Next



                    'sspcmmfvendor
                    'If myRecord(4) <> "" Then
                    '    Try
                    '        cmmfvendorHT.Add(myRecord(0).ToString & myRecord(4).ToString, myRecord(0).ToString & "," & myRecord(4).ToString)
                    '        Dim pkey(1) As Object
                    '        pkey(0) = myRecord(0).ToString
                    '        pkey(1) = myRecord(4).ToString
                    '        Dim DataRow1 As DataRow = Dataset1.Tables(2).Rows.Find(pkey) 'Table sspcmmfvendor
                    '        If DataRow1 Is Nothing Then
                    '            stringBuilder2.Append(myRecord(0) & vbTab)
                    '            stringBuilder2.Append(myRecord(4) & vbCrLf)
                    '        End If
                    '    Catch ex As Exception

                    '    End Try
                    'End If

                    ''cmmfsop Update Or Add
                    'Try
                    '    myhastable.Add(myRecord(0).ToString, myRecord(0))
                    '    Dim pkey(1) As Object
                    '    pkey(0) = myRecord(0)
                    '    pkey(1) = FTYCapId
                    '    Dim DataRow1 As DataRow = Dataset1.Tables(1).Rows.Find(pkey)
                    '    If DataRow1 Is Nothing Then
                    '        stringBuilder1.Append(myRecord(0) & vbTab)
                    '        stringBuilder1.Append(FTYCapId & vbCrLf)
                    '    End If
                    'Catch ex As Exception
                    'End Try

                    Application.DoEvents()


                Catch ex As Exception
                    errMessage = ex.Message
                    Return False
                Finally
                End Try
            Next

            'copy Fty Cap Data

            sqlstr = "copy sspftycapdata(ftycapid,periodtypeid,startingdate,datalabel,datavalue) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (Fty Cap Data)")
            If mySB.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, mySB.ToString, myreturn)
                If Not myreturn Then
                    Return False
                Else
                    BackgroundWorker1.ReportProgress(2, "Copy To Db Done.")
                End If

            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myreturn = True
            End If
            BackgroundWorker1.ReportProgress(3, "")
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function
    Private Sub FormImportSSPSOPFamily_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If BackgroundWorker1.IsBusy Then
            If MessageBox.Show("Are you sure to exit? Your process will continue running in background ", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
            Exit Sub
        End If
    End Sub

    Private Function getFTYCapId(ByVal myRecord() As String, Optional ByRef myMessage As String = "") As Long
        Dim Myreturn As Long = -1
        Dim cmd As NpgsqlCommand = New NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            'Dim sqlstr As String = "select sspftycap( from sspsopfamilies where sopfamily = :sopfamily"
            'cmd = New NpgsqlCommand(sqlstr, conn)
            'cmd.Parameters.Add("sopfamily", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(1)
            'Myreturn = cmd.ExecuteScalar
            'If Myreturn = 0 Then
            Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
            cmd1.CommandText = "insert into sspftycap(period,typeofinfoid,vendorcode,sopfamilyid) values(:period,:typeofinfoid,:vendorcode,:sopfamilyid);select currval('sspftycap_ftycapid_seq');"
            cmd1.Connection = conn
            cmd1.Parameters.Add("period", NpgsqlTypes.NpgsqlDbType.Integer).Value = myRecord(0)
            cmd1.Parameters.Add("typeofinfoid", NpgsqlTypes.NpgsqlDbType.Integer).Value = getTypeofindoid(myRecord(5))
            cmd1.Parameters.Add("vendorcode", NpgsqlTypes.NpgsqlDbType.Bigint).Value = myRecord(1)
            cmd1.Parameters.Add("sopfamilyid", NpgsqlTypes.NpgsqlDbType.Integer).Value = getsopfamilyid(myRecord(3))
            Myreturn = cmd1.ExecuteScalar
            'End If

        Catch ex As Exception
            myMessage = ex.Message
        Finally
            conn.Close()
        End Try
        Return Myreturn
    End Function

    Private Function getdate(ByVal index As Integer, ByVal period As String) As String
        Dim myresult As String = String.Empty

        Dim myyear = CInt(period.Substring(0, 4))
        Dim mymonth = CInt(period.Substring(4, 2))
        mymonth = mymonth + index
        If mymonth > 12 Then
            mymonth = mymonth - 12
            myyear = myyear + 1
        End If
        myresult = "'" & myyear & "-" & mymonth & "-1'"
        Return myresult

    End Function

    Private Function getsopfamilyid(ByVal myRecord As String) As Integer
        Dim myreturn As Integer = 0
        'Dim pkey(0) As Object
        'pkey(0) = myRecord

        'Dim DataRow1 As DataRow = Dataset1.Tables(0).Rows.Find(pkey) 'Table sspcmmfvendor
        'If Not DataRow1 Is Nothing Then
        '    myreturn = DataRow1.Item(0)
        'Else
        'check first if not avail then create
        Dim sqlstr As String = "Select sspsopfamilyid from sspsopfamilies where sopfamily = " & escapeString(myRecord)
        dbtools1.ExecuteScalar(sqlstr, myreturn)
        If myreturn = 0 Then
            sqlstr = "insert into sspsopfamilies(sopfamily) values('" & myRecord & "');select currval('sspsopfamilies_sspsopfamilyid_seq');"
            dbtools1.ExecuteScalar(sqlstr, myreturn)
        End If
        'End If

        Return myreturn
    End Function

    Private Function getTypeofindoid(ByVal myRecord As String) As Integer
        Dim myreturn As Integer = 0
        Dim pkey(0) As Object
        pkey(0) = myRecord

        Dim DataRow1 As DataRow = Dataset1.Tables(1).Rows.Find(pkey) 'Table sspcmmfvendor
        If Not DataRow1 Is Nothing Then
            myreturn = DataRow1.Item(0)        
        End If

        Return myreturn
    End Function

End Class