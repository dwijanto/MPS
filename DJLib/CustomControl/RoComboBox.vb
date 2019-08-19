Imports System.Windows.Forms

Public Class RoComboBox
    Inherits ComboBox

    Private mReadOnly As Boolean = False
    Public Property [ReadOnly]() As Boolean
        Get
            Return mReadOnly
        End Get
        Set(ByVal value As Boolean)
            mReadOnly = value
            If value Then
                Me.BackColor = Drawing.Color.FromArgb(255, 236, 233, 216) 'control
            Else
                Me.BackColor = Drawing.Color.FromArgb(255, 255, 255, 255) 'windows
            End If
        End Set
    End Property

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        If mReadOnly Then
            e.Handled = True
        End If
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If mReadOnly Then
            MyBase.OnKeyPress(e)
        End If
        MyBase.OnKeyPress(e)
    End Sub

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If ((m.Msg <> &H201 And m.Msg <> &H203) Or Not (mReadOnly)) Then
            MyBase.WndProc(m)
        End If
    End Sub



End Class
