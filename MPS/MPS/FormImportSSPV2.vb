Imports SSP.PublicClass
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class FormImportSSPV2

    'Public Property sspftycap As String
    'Public Property sspftycapdata As String
    'Public Property sspftycapseqname As String
    'Public Property sspftycapdataseqname As String
    Public Property sspTable As String
    Public Property Department As Department

    Dim myThread As New Thread(AddressOf DoWork)
    Dim ProgressReportDelegate1 As New ProgressReportDelegate(AddressOf ProgressReport)
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim FileName As String = String.Empty


    Public Sub New(ByVal mytype As Department, ByVal SSPTable As String)
        InitializeComponent()
        Me.sspTable = SSPTable
        Department = mytype
        Me.Text = "Import SSP " & mytype.ToString
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (myThread.IsAlive) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Text files (*.csv)|*.csv|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim Result As Boolean = False
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, TextBox2.Text & "Open File..")

        Result = ImportData(FileName, errMsg)
        If Result Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} {3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, "Done"))
            ProgressReport(5, "Set to continuous mode again")
        Else
            errSB.Append(errMsg & vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()

    End Sub

    Sub ProgressReport(ByVal sender As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Me.Invoke(ProgressReportDelegate1, New Object() {sender, message})
        Else
            Select Case sender
                Case 1
                    TextBox1.Text = message
                Case 2
                    'For Progress
                    ToolStripStatusLabel1.Text = message
                Case 3
                    'For Error message
                    TextBox3.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                    ToolStripStatusLabel1.Text = "Preparing Data .." & myvalue(0) & "/" & myvalue(1)
            End Select
        End If
    End Sub

    Private Sub EndWork()
        ProgressReport(2, "Done")
    End Sub


    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim list As New List(Of String())
        Dim myRecord() As String
        Dim sqlstr As String = String.Empty
        Dim message As String = String.Empty

        Dim marketseq As Long
        Dim rangeseq As Long
        Dim marketid As Long
        Dim rangeid As String
        Dim cmmfrangeSeq As Long
        Dim cmmfrangeid As Long

        Dim period As String = String.Empty

        Dim marketSB As New StringBuilder
        Dim rangeSB As New StringBuilder
        Dim cmmfrangeSB As New StringBuilder
        Dim sspSB As New StringBuilder

        Try
            ProgressReport(2, "Preparing Data...")
            ProgressReport(6, "Set Progressbar to Marque")
            Dim Dataset1 = New DataSet


            sqlstr = "select marketid,market from sspmarket;" &
                     "select marketid from sspmarket order by marketid desc limit 1;" &
                     "select max(rangeid) as myrangeid,upper(range) as myrange from ssprange" &
                     " group by myrange order by myrangeid;;" &
                     "select rangeid from ssprange order by rangeid desc limit 1;" &
                     "select sspcmmfrangeid,cmmf,rangeid from sspcmmfrange;" &
                     "select sspcmmfrangeid from sspcmmfrange order by sspcmmfrangeid desc limit 1;"



            If Not dbtools1.getDataSet(sqlstr, Dataset1, message) Then
                ProgressReport(3, message)
                Return myreturn
            End If

            Dataset1.Tables(0).TableName = "SSPMarket"
            Dim pk0(0) As DataColumn
            pk0(0) = Dataset1.Tables(0).Columns(1)
            Dataset1.Tables(0).PrimaryKey = pk0

            Dataset1.Tables(1).TableName = "SSPRange"
            Dim pk2(0) As DataColumn
            pk2(0) = Dataset1.Tables(2).Columns(1)
            Dataset1.Tables(2).PrimaryKey = pk2

            Dim pk4(1) As DataColumn
            pk4(0) = Dataset1.Tables(4).Columns(1)
            pk4(1) = Dataset1.Tables(4).Columns(2)
            Dataset1.Tables(4).PrimaryKey = pk4

            marketseq = Dataset1.Tables(1).Rows(0).Item(0).ToString
            rangeseq = Dataset1.Tables(3).Rows(0).Item(0).ToString
            cmmfrangeSeq = Dataset1.Tables(5).Rows(0).Item(0).ToString

            ProgressReport(3, "Read File...")

            Using objTFParser = New FileIO.TextFieldParser(FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(";")
                    .HasFieldsEnclosedInQuotes = True
                    Do Until .EndOfData
                        myRecord = .ReadFields
                        list.Add(myRecord)
                    Loop
                End With
            End Using


            ProgressReport(5, "Set To continuous")
            For i = 1 To list.Count - 1
                ProgressReport(7, i & "," & list.Count)
                'If i = 10165 Then
                '    Debug.Print("debugmode")
                'End If
                'get marketid
                Dim idx0(0) As Object
                idx0(0) = list(i)(6)
                Dim result = Dataset1.Tables(0).Rows.Find(idx0)
                If IsNothing(result) Then
                    'create
                    marketseq += 1
                    marketid = marketseq
                    Dim dr As DataRow = Dataset1.Tables(0).NewRow
                    dr.Item(0) = marketid
                    dr.Item(1) = list(i)(6)
                    Dataset1.Tables(0).Rows.Add(dr)
                    marketSB.Append(list(i)(6) & vbCrLf)
                    'table sspmarket : market
                Else
                    marketid = result.Item("marketid")
                End If


                'getRangeId
                If list(i)(3) = "" Then
                    list(i)(3) = "N/A"
                End If

                Dim idx2(0) As Object
                idx2(0) = list(i)(3)
                result = Dataset1.Tables(2).Rows.Find(idx2)
                If IsNothing(result) Then
                    rangeseq += 1
                    rangeid = rangeseq
                    Dim dr As DataRow = Dataset1.Tables(2).NewRow
                    dr.Item(0) = rangeid
                    dr.Item(1) = list(i)(3)
                    Dataset1.Tables(2).Rows.Add(dr)
                    rangeSB.Append(list(i)(3) & vbCrLf)
                    'table ssprange : range
                Else
                    rangeid = result.Item(0).ToString
                End If

                'If Not IsDate(HelperClass1.DateFormatDDMMYYYY(list(i)(8)).Replace("'", "")) Then
                '    Err.Raise(1, , String.Format("Invalid date format dd-mm-yyyy {0} line {1}", list(i)(8), i))
                'End If


                'getcmmfrangeid
                Dim idx4(1) As Object
                If list(i)(4) <> "" Then


                    idx4(0) = list(i)(4)
                    idx4(1) = rangeid
                    result = Dataset1.Tables(4).Rows.Find(idx4)
                    If IsNothing(result) Then
                        cmmfrangeSeq += 1
                        cmmfrangeid = cmmfrangeSeq
                        Dim dr As DataRow = Dataset1.Tables(4).NewRow
                        dr.Item(0) = cmmfrangeid
                        dr.Item(1) = list(i)(4)
                        dr.Item(2) = rangeid
                        Dataset1.Tables(4).Rows.Add(dr)
                        cmmfrangeSB.Append(list(i)(4) & vbTab & rangeid & vbCrLf)
                        'table sspcmmfrange : rangeid, cmmf
                    Else
                        cmmfrangeid = result.Item(0)
                    End If


                    'ssp data
                    If list(i)(9) = "" Then
                        list(i)(9) = 0
                    End If
                    If list(i)(10) = "" Then
                        list(i)(10) = 0
                    End If
                    If list(i)(11) = "" Then
                        list(i)(11) = 0
                    End If
                    If list(i)(15) = "" Then
                        list(i)(15) = 0
                    End If
                    If Department = SSP.Department.FinishGoods Then
                        sspSB.Append(list(i)(0) & vbTab &
                                 list(i)(1) & vbTab &
                                 cmmfrangeid & vbTab &
                                 marketid & vbTab &
                                 list(i)(7) & vbTab &
                                 HelperClass1.DateFormatDDMMYYYY(list(i)(8)) & vbTab &
                                 CInt(list(i)(9)) & vbTab &
                                 CInt(list(i)(10)) & vbTab &
                                 CInt(list(i)(11)) & vbTab &
                                 list(i)(13) & vbTab &
                                 list(i)(14) & vbTab &
                                 CDec(list(i)(15)) & vbTab &
                                 list(i)(16) & vbCrLf)
                    Else
                        sspSB.Append(list(i)(0) & vbTab &
                       list(i)(1) & vbTab &
                        cmmfrangeid & vbTab &
                        marketid & vbTab &
                        list(i)(7) & vbTab &
                        "Null" & vbTab &
                        CInt(list(i)(9)) & vbTab &
                        CInt(list(i)(10)) & vbTab &
                        CInt(list(i)(11)) & vbTab &
                        list(i)(13) & vbTab &
                        list(i)(14) & vbTab &
                        CDec(list(i)(15)) & vbTab &
                        list(i)(16) & vbCrLf)
                    End If
                    


                   
                    'table ssp: period,vendorcode,cmmfrangeid,marketid,periodof etd,startingdate,orderunconfirmed,order confirmed,forecast,unit,week,totalamount,currency
                End If
            Next

            If marketSB.Length > 0 Then
                'table sspmarket : market
                sqlstr = "copy sspmarket(market) from stdin;"
                ProgressReport(2, "Copy Market")
                errMessage = dbtools1.copy(sqlstr, marketSB.ToString, myreturn)
            End If
            If rangeSB.Length > 0 Then
                'table ssprange : range
                sqlstr = "copy ssprange(range) from stdin with null as 'Null';"
                ProgressReport(2, "Copy Range")
                errMessage = dbtools1.copy(sqlstr, rangeSB.ToString, myreturn)

            End If
            If cmmfrangeSB.Length > 0 Then
                'table sspcmmfrange : rangeid, cmmf
                sqlstr = "copy sspcmmfrange(cmmf,rangeid) from stdin;"
                ProgressReport(2, "Copy CMMFRange")

                errMessage = dbtools1.copy(sqlstr, cmmfrangeSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If

            End If

            If sspSB.Length > 0 Then
                sqlstr = "Delete from " & sspTable & " where period = " & list(1)(0) & ";select setval('" & sspTable & "_sspid_seq',(select sspid from " & sspTable & " order by sspid desc limit 1)+1,false);" & _
                     "copy " & sspTable & "(period,vendorcode,sspcmmfrangeid,marketid,periodofetd,startingdate,orderunconfirmed,orderconfirmed,forecast,unit,week,totalamount,crcycode) from stdin with null as 'Null';"
                ProgressReport(2, "Copy " & sspTable)

                errMessage = dbtools1.copy(sqlstr, sspSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            End If


        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function


End Class