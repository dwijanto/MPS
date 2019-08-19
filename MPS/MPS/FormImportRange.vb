Imports DJLib
Imports DJLib.Dbtools
Imports System.ComponentModel
Imports System.IO
Imports Npgsql
Imports System.Text

Public Class FormImportRange
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
        FormMenu.setBubbleMessage("Import Range", "Done")
        If Status Then
            If CheckBox1.Checked Then
                Me.Close()
            End If
        End If
    End Sub
    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim list As New List(Of String)
        Dim myRecord() As String
        Dim sqlstr As String = String.Empty
        Dim stopwatch As New Stopwatch
        stopwatch.Start()
        Try
            BackgroundWorker1.ReportProgress(2, "Preparing Data...")
            Dataset1 = New DataSet

            sqlstr = "select range,rangedesc from range;"
                     
            dbtools1.getDataSet(sqlstr, Dataset1)
            Dim keys(0) As DataColumn
            keys(0) = Dataset1.Tables(0).Columns(0)
            Dataset1.Tables(0).PrimaryKey = keys


            Dataset1.Tables(0).TableName = "Range"

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

            Dim myhashtable As New Hashtable 'used for duplicate checking
            Dim mySB As New StringBuilder
            Dim myMonth As String = String.Empty

            myRecord = list(1).Split(vbTab)

            'Remove existing period

            For i = 3 To list.Count - 1
                Try
                    BackgroundWorker1.ReportProgress(3, "Processing line " & i & " of " & list.Count - 1)
                    myRecord = list(i).Split(vbTab)

                    If myRecord(1) <> "" Then
                        Try
                            myhashtable.Add(myRecord(1).ToString, myRecord(1).ToString & "," & myRecord(2).ToString)
                            Dim pkey(0) As Object
                            pkey(0) = myRecord(1).ToString
                            Dim DataRow1 As DataRow = Dataset1.Tables(0).Rows.Find(pkey) 'Table range
                            If DataRow1 Is Nothing Then
                                mySB.Append(myRecord(1) & vbTab)
                                mySB.Append(myRecord(2) & vbCrLf)
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    Application.DoEvents()
                Catch ex As Exception
                    errMessage = ex.Message
                    Return False
                Finally
                End Try
            Next

            'copy Fty Cap Data

            sqlstr = "copy range(range,rangedesc) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (Fty Cap Data)")
            If mySB.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, mySB.ToString, myreturn)
                BackgroundWorker1.ReportProgress(2, "Copy To Db Done.")
            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myreturn = True
            End If
            stopwatch.Stop()
            Dim et As Double = stopwatch.Elapsed.Minutes
            BackgroundWorker1.ReportProgress(3, Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & (stopwatch.ElapsedMilliseconds / 1000).ToString)
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function
End Class