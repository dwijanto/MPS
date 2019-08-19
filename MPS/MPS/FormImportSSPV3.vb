Imports SSP.PublicClass
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class FormImportSSPV3
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
                Dim message = String.Copy(FileName)
                TextRenderer.MeasureText(Message, Font, New Drawing.Size(TextBox1.Width, 0), TextFormatFlags.PathEllipsis Or TextFormatFlags.ModifyString)
                TextBox1.Text = message
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
                    TextRenderer.MeasureText(message, Font, New Drawing.Size(TextBox1.Width, 0), TextFormatFlags.PathEllipsis Or TextFormatFlags.ModifyString)
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
        ' Dim message As String = String.Empty

        Dim marketseq As Long
        Dim rangeseq As Long
        Dim marketid As Long
        'Dim sopid As Long
        Dim mpsId As Long
        Dim rangeid As String
        Dim cmmfrangeSeq As Long
        Dim cmmfrangeid As Long
        ' Dim cmmfsopseq As Long
        ' Dim cmmfsopid As Long
        Dim cmmfmpsseq As Long
        Dim cmmfmpsid As Long
        'Dim sopseq As Long
        Dim mpsseq As Long


        Dim period As String = String.Empty

        Dim marketSB As New StringBuilder
        Dim rangeSB As New StringBuilder
        Dim cmmfrangeSB As New StringBuilder
        Dim sspSB As New StringBuilder
        Dim sopSB As New StringBuilder
        Dim mpsSB As New StringBuilder
        Dim cmmfsopSB As New StringBuilder
        Dim cmmfmpsSB As New StringBuilder
        Dim updCMMFSOPSB As New StringBuilder
        Dim updCMMFMPSSB As New StringBuilder
        Try
            ProgressReport(2, "Preparing Data...")
            ProgressReport(6, "Set Progressbar to Marque")
            Dim Dataset1 = New DataSet

            'cmmfsopfamily -> sspsopfamilies

            sqlstr = "select marketid,market from sspmarket;" &
                     "select marketid from sspmarket order by marketid desc limit 1;" &
                     "select max(rangeid) as myrangeid,upper(range) as myrange from ssprange group by myrange order by myrangeid;" &
                     "select rangeid from ssprange order by rangeid desc limit 1;" &
                     "select sspcmmfrangeid,cmmf,rangeid from sspcmmfrange;" &
                     "select sspcmmfrangeid from sspcmmfrange order by sspcmmfrangeid desc limit 1;" &
                     "select id,mpsfamily,mpsdesc from mpsfamily;" &
                     "select id from mpsfamily order by id desc limit 1;" &                     
                     "select id,cmmf,mpsfamilyid from sspcmmfmps;" &
                     "select id from sspcmmfmps order by id desc limit 1;"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
                ProgressReport(3, errMessage)
                Return myreturn
            End If

            Dataset1.Tables(0).TableName = "SSPMarket"
            Dim pk0(0) As DataColumn
            pk0(0) = Dataset1.Tables(0).Columns(1)
            Dataset1.Tables(0).PrimaryKey = pk0

            Dataset1.Tables(2).TableName = "SSPRange"
            Dim pk2(0) As DataColumn
            pk2(0) = Dataset1.Tables(2).Columns(1)
            Dataset1.Tables(2).PrimaryKey = pk2

            Dataset1.Tables(4).TableName = "SSPCMMFRange"
            Dim pk4(1) As DataColumn
            pk4(0) = Dataset1.Tables(4).Columns(1)
            pk4(1) = Dataset1.Tables(4).Columns(2)
            Dataset1.Tables(4).PrimaryKey = pk4

            'Dataset1.Tables(6).TableName = "SOPFamily"
            'Dim pk6(0) As DataColumn
            'pk6(0) = Dataset1.Tables(6).Columns(1)            
            'Dataset1.Tables(6).PrimaryKey = pk6

            Dataset1.Tables(6).TableName = "MPSFamily"
            Dim pk6(0) As DataColumn
            pk6(0) = Dataset1.Tables(6).Columns(1)
            Dataset1.Tables(6).PrimaryKey = pk6

            'Dataset1.Tables(10).TableName = "SSPCMMFSOP"
            'Dim pk10(1) As DataColumn
            'pk10(0) = Dataset1.Tables(10).Columns(1)
            'pk10(1) = Dataset1.Tables(10).Columns(2)

            'Dataset1.Tables(10).PrimaryKey = pk10


            Dataset1.Tables(8).TableName = "SSPCMMFMPS"
            Dim pk8(0) As DataColumn
            pk8(0) = Dataset1.Tables(8).Columns(1)
            Dataset1.Tables(8).PrimaryKey = pk8

            marketseq = Dataset1.Tables(1).Rows(0).Item(0).ToString
            rangeseq = Dataset1.Tables(3).Rows(0).Item(0).ToString
            cmmfrangeSeq = Dataset1.Tables(5).Rows(0).Item(0).ToString

            If Dataset1.Tables(7).Rows.Count > 0 Then
                mpsseq = Dataset1.Tables(7).Rows(0).Item(0)
            End If

            If Dataset1.Tables(9).Rows.Count > 0 Then
                cmmfmpsseq = Dataset1.Tables(9).Rows(0).Item(0)
            End If

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


                'get marketid
                Dim idx0(0) As Object
                idx0(0) = list(i)(9)
                Dim result = Dataset1.Tables(0).Rows.Find(idx0)
                If IsNothing(result) Then
                    'create
                    marketseq += 1
                    marketid = marketseq
                    Dim dr As DataRow = Dataset1.Tables(0).NewRow
                    dr.Item(0) = marketid
                    dr.Item(1) = list(i)(9)
                    Dataset1.Tables(0).Rows.Add(dr)
                    marketSB.Append(list(i)(9) & vbCrLf)
                    'table sspmarket : market
                Else
                    marketid = result.Item("marketid")
                End If


                'getRangeId
                'If list(i)(3) = "" Then
                '    list(i)(3) = "N/A"
                'End If

                Dim idx2(0) As Object
                idx2(0) = "N/A" 'list(i)(3)
                result = Dataset1.Tables(2).Rows.Find(idx2)
                If IsNothing(result) Then
                    rangeseq += 1
                    rangeid = rangeseq
                    Dim dr As DataRow = Dataset1.Tables(2).NewRow
                    dr.Item(0) = rangeid
                    dr.Item(1) = "N/A" 'list(i)(3)
                    Dataset1.Tables(2).Rows.Add(dr)
                    'rangeSB.Append(list(i)(3) & vbCrLf)
                    rangeSB.Append("N/A" & vbCrLf)
                    'table ssprange : range
                Else
                    rangeid = result.Item(0).ToString
                End If


                'find sopfamily
                'Dim idx6(0) As Object
                'idx6(0) = list(i)(2)
                'result = Dataset1.Tables(6).Rows.Find(idx6)
                'If IsNothing(result) Then
                '    'create
                '    sopseq += 1
                '    sopid = sopseq
                '    Dim dr As DataRow = Dataset1.Tables(6).NewRow
                '    dr.Item(0) = sopid
                '    dr.Item(1) = list(i)(2)
                '    dr.Item(2) = list(i)(3)
                '    Dataset1.Tables(6).Rows.Add(dr)
                '    sopSB.Append(list(i)(2) & vbTab & list(i)(3) & vbCrLf)                    
                'Else
                '    sopid = result.Item("sopfamilyid")
                'End If

                'find mpsfamily
                Dim idx6(0) As Object
                idx6(0) = list(i)(4)
                result = Dataset1.Tables(6).Rows.Find(idx6)
                If IsNothing(result) Then
                    'create
                    mpsseq += 1
                    mpsId = mpsseq
                    Dim dr As DataRow = Dataset1.Tables(6).NewRow
                    dr.Item(0) = mpsId
                    dr.Item(1) = list(i)(4)
                    dr.Item(2) = list(i)(5)
                    Dataset1.Tables(6).Rows.Add(dr)
                    mpsSB.Append(list(i)(4) & vbTab & list(i)(5) & vbCrLf)
                Else
                    mpsId = result.Item("id")
                End If

                'getcmmfrangeid

                If list(i)(6) <> "" Then
                    Dim idx4(1) As Object
                    'cmmfrangeid
                    idx4(0) = list(i)(6)
                    idx4(1) = rangeid
                    result = Dataset1.Tables(4).Rows.Find(idx4)
                    If IsNothing(result) Then
                        cmmfrangeSeq += 1
                        cmmfrangeid = cmmfrangeSeq
                        Dim dr As DataRow = Dataset1.Tables(4).NewRow
                        dr.Item(0) = cmmfrangeid
                        dr.Item(1) = list(i)(6)
                        dr.Item(2) = rangeid
                        Dataset1.Tables(4).Rows.Add(dr)
                        cmmfrangeSB.Append(list(i)(6) & vbTab & rangeid & vbCrLf)
                        'table sspcmmfrange : rangeid, cmmf
                    Else
                        cmmfrangeid = result.Item(0)
                    End If

                    'cmmfsopfamily : 1 CMMF = 1 SOP Family
                    'Dim idx10(0) As Object
                    'idx10(0) = list(i)(6)                    
                    'result = Dataset1.Tables(10).Rows.Find(idx10)
                    'If IsNothing(result) Then
                    '    cmmfsopseq += 1
                    '    cmmfsopid = cmmfsopseq
                    '    Dim dr As DataRow = Dataset1.Tables(10).NewRow
                    '    dr.Item(0) = cmmfsopid
                    '    dr.Item(1) = list(i)(6)
                    '    dr.Item(2) = sopid
                    '    Dataset1.Tables(10).Rows.Add(dr)
                    '    cmmfsopSB.Append(list(i)(6) & vbTab & sopid & vbCrLf)
                    '    'table sspcmmfrange : rangeid, cmmf
                    'Else
                    '    cmmfsopid = result.Item(0)
                    '    If result.Item(2) <> sopid Then

                    '    End If
                    '    'Update sopid if not equal
                    'End If

                    'cmmfmpsfamily
                    Dim idx8(0) As Object
                    idx8(0) = list(i)(6)
                    result = Dataset1.Tables(8).Rows.Find(idx8)
                    If IsNothing(result) Then
                        cmmfmpsseq += 1
                        cmmfmpsid = cmmfmpsseq
                        Dim dr As DataRow = Dataset1.Tables(8).NewRow
                        dr.Item(0) = cmmfmpsid
                        dr.Item(1) = list(i)(6)
                        dr.Item(2) = mpsId
                        Dataset1.Tables(8).Rows.Add(dr)
                        cmmfmpsSB.Append(list(i)(6) & vbTab & mpsId & vbCrLf)
                        'table sspcmmfrange : rangeid, cmmf
                    Else
                        cmmfmpsid = result.Item(0)
                        If result.Item(2) <> mpsId Then
                            If updCMMFMPSSB.Length > 0 Then
                                updCMMFMPSSB.Append(",")
                            End If
                            updCMMFMPSSB.Append(String.Format("[{0}::character varying,'{1}']", cmmfmpsid, mpsId))
                        End If
                        'Update mpsid if not equal
                    End If



                    'ssp data
                    If Department = SSP.Department.FinishGoods Then
                        If list(i)(10) = "" Then
                            list(i)(10) = 0
                        End If
                        If list(i)(11) = "" Then
                            list(i)(11) = 0
                        End If
                        If list(i)(12) = "" Then
                            list(i)(12) = 0
                        End If
                        '
                        'period,vendorcode,sspcmmfrangeid,marketid,periodofetd,startingdate,orderunconfirmed,orderconfirmed,forecast,unit,week,totalamount,crcycode
                        'PI2 Supplier Code;Supplier Name;SOP Family;SOP Description;MPS Family;MPS Description;Product;Product Description;Calendar Year/Week;Market;PO	Forecast;Total;Year_Week;Year-Month;Supplier Code;current week;starting date
                        Dim mydate = CDate(list(i)(17))
                        sspSB.Append(list(i)(16) & vbTab &
                                 validstr(list(i)(15)) & vbTab &
                                 cmmfrangeid & vbTab &
                                 marketid & vbTab &
                                 "Null" & vbTab &
                                 String.Format("'{0:yyyy}-{0:MM}-{0:dd}'", CDate(list(i)(17))) & vbTab &
                                 0 & vbTab &
                                 CInt(list(i)(10)) & vbTab &
                                 CInt(list(i)(11)) & vbTab &
                                "Null" & vbTab &
                                 list(i)(13) & vbTab &
                                0 & vbTab &
                                "Null" & vbTab &
                                list(i)(0) & vbCrLf)
                    Else
                        'Need To check first
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
                If Not myreturn Then
                    Return myreturn
                End If
            End If
            If rangeSB.Length > 0 Then
                'table ssprange : range
                sqlstr = "copy ssprange(range) from stdin with null as 'Null';"
                ProgressReport(2, "Copy Range")
                errMessage = dbtools1.copy(sqlstr, rangeSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If

            End If
            'If sopSB.Length > 0 Then
            '    'table ssprange : range
            '    sqlstr = "copy sopfamily(sopfamily,familydesc) from stdin with null as 'Null';"
            '    ProgressReport(2, "Copy SOPFamily")
            '    errMessage = dbtools1.copy(sqlstr, sopSB.ToString, myreturn)
            '    If Not myreturn Then
            '        Return myreturn
            '    End If

            'End If
            If mpsSB.Length > 0 Then
                'table ssprange : range
                sqlstr = "copy mpsfamily(mpsfamily,mpsdesc) from stdin with null as 'Null';"
                ProgressReport(2, "Copy MPSFamily")
                errMessage = dbtools1.copy(sqlstr, mpsSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
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
            'If cmmfsopSB.Length > 0 Then
            '    'table sspcmmfrange : rangeid, cmmf
            '    sqlstr = "copy sspcmmfsop(cmmf,sopfamilyid) from stdin;"
            '    ProgressReport(2, "Copy CMMFSOP")

            '    errMessage = dbtools1.copy(sqlstr, cmmfsopSB.ToString, myreturn)
            '    If Not myreturn Then
            '        Return myreturn
            '    End If

            'End If
            If cmmfmpsSB.Length > 0 Then
                'table sspcmmfrange : rangeid, cmmf
                sqlstr = "copy sspcmmfmps(cmmf,mpsfamilyid) from stdin;"
                ProgressReport(2, "Copy CMMFMPS")

                errMessage = dbtools1.copy(sqlstr, cmmfmpsSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If

            End If

            If updCMMFMPSSB.Length > 0 Then
                ProgressReport(2, "Update sspsopfamilies")
                sqlstr = "update sspcmmfmps set mpsfamilyid = foo.mpsfamilyid::bigint from (select * from array_to_set2(Array[" & updCMMFMPSSB.ToString & "]) as tb (id character varying,mpsfamilyid character varying))foo where sspcmmfmps.id = foo.id::bigint;"
                Dim ra As Long

                If Not dbtools1.ExecuteNonQuery(sqlstr, ra, errMessage) Then
                    Return False
                End If
            End If

            If sspSB.Length > 0 Then
                sqlstr = "Delete from " & sspTable & " where period = " & list(1)(16) & ";select setval('" & sspTable & "_sspid_seq',(select sspid from " & sspTable & " order by sspid desc limit 1)+1,false);" & _
                     "copy " & sspTable & "(period,vendorcode,sspcmmfrangeid,marketid,periodofetd,startingdate,orderunconfirmed,orderconfirmed,forecast,unit,week,totalamount,crcycode,pi2suppliercode) from stdin with null as 'Null';"
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

    Private Sub TextBox1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.SizeChanged
        If FileName <> "" Then
            Dim message = String.Copy(FileName) 'OpenFileDialog1.FileName
            TextBox1.Text = ""
            TextRenderer.MeasureText(Message, Font, New Drawing.Size(TextBox1.Width, 0), TextFormatFlags.PathEllipsis Or TextFormatFlags.ModifyString)
            TextBox1.Text = message
        End If
    End Sub

    Private Function validstr(ByVal p1 As String) As String
        If p1 = "" Then
            p1 = "Null"
        End If
        Return p1
    End Function

End Class