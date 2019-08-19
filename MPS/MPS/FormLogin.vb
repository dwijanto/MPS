Imports DJLib
Imports DJLib.Dbtools

Imports SSP.PublicClass

Public Class FormLogin
    Private trycount As Byte

    Private Sub ButtonLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLogin.Click
        Dim groupname As String = Nothing
        Dim errorMessage As String = Nothing
        myUserid = userId.Text
        myPassword = password.Text

        'Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        dbtools1 = New DJLib.Dbtools(myUserid, myPassword)

        If DbTools1.canLogin(errorMessage) Then
            myGroup = DbTools1.Group
            Dim connectionString As String = DbTools1.getConnectionString.ToString
            Dim connectionstrings() As String = connectionString.Split(";")
            For i = 0 To (connectionstrings.Length - 1)
                Dim mystrs() As String = connectionstrings(i).Split("=")
                ConnectionStringCollections.Add(mystrs(1), mystrs(0))
            Next i
            Dim myhost As String = ConnectionStringCollections.Item("HOST")

            DBAdapter1 = New DBAdapter
            Try
                loglogin(myUserid)
            Catch ex As Exception
            End Try
            FormMenu.Show()

            'If Not SecureConnectionString.ProtectConnectionString(True, errorMessage) Then
            '    MsgBox(errorMessage)
            'End If
        Else
            MsgBox(errorMessage)
            trycount = trycount + 1
            If trycount < 3 Then
                Exit Sub
            Else
                MsgBox("Please contact admin!")
                Call ButtonExit_Click(Me, e)
            End If
        End If

        Me.Close()
    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "SSP"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub
    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        fadeout(Me)
        Application.Exit()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub password_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles password.KeyDown
        If e.KeyCode = 13 Then
            ButtonLogin_Click(sender, e)
        End If
    End Sub

    Private Sub userId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles userId.KeyDown
        If e.KeyCode = 13 Then
            ButtonLogin_Click(sender, e)
        End If
    End Sub

    Private Sub userId_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles userId.Validating, password.Validating
        If userId.Text = "" Then
            ErrorProvider1.SetError(userId, "Userid cannot be blank!")
        Else
            ErrorProvider1.SetError(userId, "")
        End If

    End Sub

    Private Sub password_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles password.Validating
        If password.Text = "" Then
            ErrorProvider1.SetError(password, "Password cannot be blank!")
        Else
            ErrorProvider1.SetError(password, "")
        End If
    End Sub

    Private Sub FormLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fadein(Me)
    End Sub

    Private Sub password_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles password.TextChanged

    End Sub
    
End Class
