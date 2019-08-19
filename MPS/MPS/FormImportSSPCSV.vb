Imports DJLib
Imports DJLib.Dbtools
Imports System.ComponentModel
Imports System.IO
Imports Npgsql
Imports System.Text

Public Class FormImportSSPCSV
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String
    Dim Status As Boolean = False
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim myconverter As New utf8towin1252

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
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
            BackgroundWorker1.ReportProgress(3, "Error::" & errMsg)
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
        FormMenu.setBubbleMessage("Import SSP CSV", "Done")
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
        Dim myRecord() As String
        Dim MarketId As Long = 0
        Dim RangeId As Long = 0
        Dim SSPCMMFRangeid As Long = 0
        Dim period As Long = 0


        BackgroundWorker1.ReportProgress(2, "Preparing Data...")
        Dim DataSet1 As New DataSet
        BackgroundWorker1.ReportProgress(3, "Convert TextFile to Unicode..")
        BackgroundWorker1.ReportProgress(3, "Connect to Db...")
        Dim sqlstr As String = "select marketid,market from sspmarket;select rangeid,range from ssprange;select sspcmmfrangeid,cmmf,rangeid from sspcmmfrange;"
        dbtools1.getDataSet(sqlstr, DataSet1)

        Try
            DataSet1.Tables(0).TableName = "SSPMarket"
            DataSet1.Tables(1).TableName = "SSPRange"
            DataSet1.Tables(2).TableName = "SSPCMMFRange"
            BackgroundWorker1.ReportProgress(3, "Open File...")

            'Dim ascii As Encoding = Encoding.ASCII
            'Dim utf8 As Encoding = Encoding.UTF8
            'Using myStream As StreamReader = File.OpenText(FileName)
            Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)
                Dim line As String = myStream.ReadLine

                Do While (Not line Is Nothing)
                    'Dim utf8bytes As Byte() = utf8.GetBytes(line)
                    'Dim asciibytes As Byte() = Encoding.Convert(utf8, ascii, utf8bytes)
                    'line = Encoding.ASCII.GetString(asciibytes)
                    list.Add(line)
                    line = myStream.ReadLine
                Loop
            End Using
            myRecord = list(1).Split(";")
            period = myRecord(0)
            BackgroundWorker1.ReportProgress(2, "Scanning Data...")
            For i = 1 To list.Count - 1
                Try
                    BackgroundWorker1.ReportProgress(3, "Processing line " & i & " of " & list.Count - 1)
                    myRecord = list(i).Split(";")
                    MarketId = getMarketId(myRecord(6))
                    RangeId = getRangeId(myRecord(3))
                    SSPCMMFRangeid = getSSPCMMFRange(myRecord(4), RangeId)
                    stringBuilder1.Append(myRecord(0) & vbTab) 'period
                    stringBuilder1.Append(myRecord(1) & vbTab) 'vendorcode
                    stringBuilder1.Append(CInt(SSPCMMFRangeid) & vbTab) 'sspcmmfrangeid
                    stringBuilder1.Append(CInt(MarketId) & vbTab) 'marketid
                    stringBuilder1.Append(myRecord(7) & vbTab) 'periodofetd
                    stringBuilder1.Append(DateFormatDDMMYYYY(myRecord(8)) & vbTab) 'startingdate
                    stringBuilder1.Append(CInt(myRecord(9)) & vbTab) 'orderunconfirmed
                    stringBuilder1.Append(CInt(myRecord(10)) & vbTab) 'orderconfirmed
                    stringBuilder1.Append(CInt(myRecord(11)) & vbTab) 'forecast
                    stringBuilder1.Append(myRecord(13) & vbTab) 'unit
                    stringBuilder1.Append(CInt(myRecord(14)) & vbTab) 'week
                    stringBuilder1.Append(CDec(myRecord(15)) & vbTab) 'totalamount
                    stringBuilder1.Append(myRecord(16) & vbCrLf) 'crcycode
                    Application.DoEvents()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Next
            sqlstr = "Delete from ssp where period = " & period & ";select setval('ssp_sspid_seq',(select sspid from ssp order by sspid desc limit 1)+1,false);copy ssp(period,vendorcode,sspcmmfrangeid,marketid,periodofetd,startingdate,orderunconfirmed,orderconfirmed,forecast,unit,week,totalamount,crcycode) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db")

            errMessage = dbtools1.copy(sqlstr, stringBuilder1.ToString, myreturn)
            BackgroundWorker1.ReportProgress(2, "Copy To Db Done.")
            BackgroundWorker1.ReportProgress(3, "")
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function
    Private Function getMarketId(ByVal market As String) As Long
        Dim Myreturn As Long = 0
        Dim cmd As NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select marketid from sspmarket where market = :market"
            cmd = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("market", NpgsqlTypes.NpgsqlDbType.Varchar).Value = market
            Myreturn = cmd.ExecuteScalar
            If Myreturn = 0 Then
                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
                cmd1.CommandText = "insert into sspmarket(market) values(:market);select currval('sspmarket_marketid_seq');"
                cmd1.Connection = conn
                cmd1.Parameters.Add("market", NpgsqlTypes.NpgsqlDbType.Varchar).Value = market
                Myreturn = cmd1.ExecuteScalar
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
        Return Myreturn
    End Function
    Private Function getRangeId(ByVal Range As String) As Long
        Dim Myreturn As Long = 0
        Dim cmd As NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select rangeid from ssprange where range = :range"
            cmd = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("range", NpgsqlTypes.NpgsqlDbType.Varchar).Value = Range
            Myreturn = cmd.ExecuteScalar
            If Myreturn = 0 Then
                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
                cmd1.CommandText = "insert into ssprange(range) values(:range);select currval('ssprange_rangeid_seq');"
                cmd1.Connection = conn
                cmd1.Parameters.Add("range", NpgsqlTypes.NpgsqlDbType.Varchar).Value = Range
                Myreturn = cmd1.ExecuteScalar
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
        Return Myreturn

    End Function

    Private Function getSSPCMMFRange(ByVal cmmf As String, ByVal rangeid As String) As Long
        Dim Myreturn As Long = 0
        Dim cmd As NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select sspcmmfrangeid from sspcmmfrange where rangeid = :rangeid and cmmf = :cmmf"
            cmd = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("rangeid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = rangeid
            cmd.Parameters.Add("cmmf", NpgsqlTypes.NpgsqlDbType.Bigint).Value = cmmf
            Myreturn = cmd.ExecuteScalar
            If Myreturn = 0 Then
                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
                cmd1.CommandText = "insert into sspcmmfrange(cmmf,rangeid) values(:cmmf,:rangeid);select currval('sspcmmfrange_sspcmmfrangeid_seq');"
                cmd1.Connection = conn
                cmd1.Parameters.Add("rangeid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = rangeid
                cmd1.Parameters.Add("cmmf", NpgsqlTypes.NpgsqlDbType.Bigint).Value = cmmf
                Myreturn = cmd1.ExecuteScalar
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
        Return Myreturn
    End Function

    Private Sub FormImportSSPCSV_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If BackgroundWorker1.IsBusy Then
            If MessageBox.Show("Are you sure to exit? Your process will continue running in background ", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
            Exit Sub
        End If
    End Sub
End Class