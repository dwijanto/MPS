Imports DJLib
Imports DJLib.Dbtools
Imports System.Threading
Imports SSP.PublicClass
Public Class FormMenu

    Private DynamicMenu1 As DynamicMenu
    Private MenuStrip1 As MenuStrip
    Private StatustStrip1 As StatusStrip
    Private ToolStripStatusLable1 As ToolStripStatusLabel
    Private WithEvents ToolStripButton1 As ToolStripButton
    'private dbtools1 As New Dbtools(myUserid, myPassword)

    Private DataTable1 As New DataTable

    Public Sub New()

        'StatustStrip1 = New System.Windows.Forms.StatusStrip
        'ToolStripStatusLable1 = New System.Windows.Forms.ToolStripStatusLabel
        'ToolStripButton1 = New System.Windows.Forms.ToolStripButton

        ' This call is required by the designer.
        InitializeComponent()
        StatustStrip1 = New System.Windows.Forms.StatusStrip
        ToolStripStatusLable1 = New System.Windows.Forms.ToolStripStatusLabel
        ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        dbtools1.Userid = myUserid
        dbtools1.Password = myPassword

        DBAdapter1 = New DBAdapter
        ' Add any initialization after the InitializeComponent() call.
        Me.Text = "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & dbtools1.Userid
        Me.Location = New Point(300, 10)
    End Sub

    Private Sub FormMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim errMessage As String = vbNull
        ' If Not dbtools1.getData("Select isactive,programid,parentid,myorder,description,programname,icon,iconindex from program  where isactive order by parentid,myorder", DataTable1, errMessage) Then
        Dim sqlstr As String = String.Empty
        'If ConnectionStringCollections.Item("HOST") = "hon14nt" Then
        'sqlstr = "Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,_groupname from sspprogram " & _
        '                    " left join _groupuser gu on gu._membername = '" & myUserid & "'" & _
        '                    " where isactive and  members ~ '\m" & myGroup & "\y' order by parentid,myorder"
        sqlstr = "Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,_groupname from sspprogram " & _
                            " left join _groupuser gu on gu._membername = '" & myUserid & "'" & _
                            " where isactive and  members ~ ('\m'" & "|| gu._groupname ||" & "'\y') order by parentid,myorder"
        'Else
        'sqlstr = "Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,_groupname from sspprogram " & _
        '                    " left join _groupuser gu on gu._membername = '" & myUserid & "'" & _
        '                    " where isactive and  members ~ '\\m" & myGroup & "\\y' order by parentid,myorder"
        'End If
        If Not dbtools1.getData(sqlstr, DataTable1, errMessage) Then
            MsgBox(errMessage)
        Else
            If DataTable1.Rows.Count > 0 Then
                DynamicMenu1 = New DynamicMenu(Me, DataTable1, ImageList1)
                'DataGridView1.DataSource = DataTable1
                DynamicMenu1.LoadMenu(MenuStrip1)
                Me.MainMenuStrip = MenuStrip1
                Me.Controls.Add(MenuStrip1)

            Else
                MessageBox.Show("You don't have any access. Please contact admin!", "No menu found for this user")
                Me.Close()
            End If

        End If
    End Sub

    Private Sub MenuItemOnClick_mSSPMonthlyTable(ByVal sender As Object, ByVal e As System.EventArgs)
        'FormImportSSPCSV.Show()
        Dim myform As New FormMonthlyTable()
        myform.Show()
    End Sub

    Private Sub MenuItemOnClick_mSSPCSV(ByVal sender As Object, ByVal e As System.EventArgs)
        'FormImportSSPCSV.Show()
        'Dim myform As New FormImportSSPV2(Department.FinishGoods, "ssp")
        Dim myform As New FormImportSSPV3(Department.FinishGoods, "ssp")
        myform.Show()
    End Sub

    Private Sub MenuItemOnClick_mSSPCSVNew(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform As New FormImportSSPV2(Department.Components, "sspcomp")
        myform.Show()
    End Sub
    Private Sub MenuItemOnClick_mSSPSOPTXT(ByVal sender As Object, ByVal e As System.EventArgs)
        'FormImportSOPFamily.Show()
        Dim myform As New FormImportSOPFamilies2(Department.FinishGoods)
        myform.Show()
    End Sub
    Private Sub MenuItemOnClick_mSSPSOPTXT2(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform As New FormImportSOPFamilies2(Department.Components)
        myform.Show()
    End Sub
    Private Sub MenuItemOnClick_mFCTCAP(ByVal sender As Object, ByVal e As System.EventArgs)
        'FormImportFTYCap.Show()
        Dim myform As New FormImportFTYCapV2(Department.FinishGoods, "sspftycap", "sspftycap_ftycapid_seq", "sspftycapdata", "sspftycapdata_ftycapdataid_seq")
        myform.Show()
    End Sub
    Private Sub MenuItemOnClick_mFCTCAPNew(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform As New FormImportFTYCapV2(Department.Components, "sspftycapcomp", "sspftycapcomp_ftycapid_seq", "sspftycapdatacomp", "sspftycapdatacomp_ftycapdataid_seq")
        myform.Show()
    End Sub
    Private Sub MenuItemOnClick_mImportRange(ByVal sender As Object, ByVal e As System.EventArgs)
        FormImportRange.Show()
    End Sub

    Private Sub MenuItemOnClick_mSSPComparisonCompMonthly(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform = New FormSSPComparisonMonthly(SSPCompMonthlyReport.Vendor)
        'FormSSPComparisonMonthly.Show()
        myform.Show()
    End Sub

    Private Sub MenuItemOnClick_mSSPComparisonCompMonthlyFactory(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform = New FormSSPComparisonMonthly(SSPCompMonthlyReport.Factory)
        'FormSSPComparisonMonthly.Show()
        myform.Show()
    End Sub

    Private Sub MenuItemOnClick_mReport1(ByVal sender As Object, ByVal e As System.EventArgs)
        'Report1.Show()
        'Dim ReportSSPComparison As New Report1 With {.department = Department.FinishGoods,
        '                                             .TableName = "ssp",
        '                                             .ViewName = "sopallnew"}
        Dim ReportSSPComparison As New SSPComparisonFG
        ReportSSPComparison.Show()

    End Sub
    Private Sub MenuItemOnClick_mReport1Comp(ByVal sender As Object, ByVal e As System.EventArgs)
        'Report1.Show()
        Dim ReportSSPComparison As New Report1 With {.department = Department.Components,
                                                     .TableName = "sspcomp",
                                                     .ViewName = "sopcompall"}
        ReportSSPComparison.Show()

    End Sub
    Private Sub MenuItemOnClick_mWMPS(ByVal sender As Object, ByVal e As System.EventArgs)
        FormWeeklyMPS.Show()
    End Sub

    Private Sub MenuItemOnClick_mWorkingDays(ByVal sender As Object, ByVal e As System.EventArgs)
        DialogWorkingDays.ShowDialog()
    End Sub
    Private Sub MenuItemOnClick_mGroupFamily(ByVal sender As Object, ByVal e As System.EventArgs)
        FormGroupFamily.ShowDialog()
    End Sub


    Private Sub MenuItemOnClick_mMPSReport1(ByVal sender As Object, ByVal e As System.EventArgs)
        MPSReport1.Show()
    End Sub
    Private Sub MenuItemOnClick_mMPSReport2(ByVal sender As Object, ByVal e As System.EventArgs)
        MPSReport2.Show()
    End Sub

    Private Sub MenuItemOnClick_mImpVendorSSM(ByVal sender As Object, ByVal e As System.EventArgs)
        FormUpdateVendorSSM.Show()
    End Sub
    Private Sub MenuItemOnClick_mExit(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer = MsgBox("Are you sure?", MsgBoxStyle.OkCancel)
        If i = 1 Then
            For i = 1 To (My.Application.OpenForms.Count - 1)
                My.Application.OpenForms.Item(1).Close()
            Next
            fadeout(Me)
            Me.Close()
        End If
    End Sub

    Private Sub MenuItemOnClick_mImpMPSFamily(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myform As New FormImportMPSFamily
        myform.Show()
    End Sub
    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.Visible = True
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.ShowBalloonTip(200)
        ShowballonWindow(200)
    End Sub
    Private Sub ShowballonWindow(ByVal timeout As Integer)
        If timeout <= 0 Then
            Exit Sub
        End If
        Dim timeoutcount As Integer = 0
        While (timeoutcount < timeout)
            Thread.Sleep(1)
            timeoutcount += 1
        End While
        NotifyIcon1.Visible = False
    End Sub

End Class
Public Enum Department
    FinishGoods
    Components
End Enum