Imports System.Windows.Forms

Public Class DialogWorkingDays

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogWorkingDays_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'check data unsaved data
        Weektomonth1.AllowValidate = False
        Dim DataRow As DataRow = Nothing
        If Weektomonth1.CheckRowStateChanged(DataRow) Then
            Select MsgBox("There is unsave data in a row." & vbCrLf & "Do you want to store to the database?", vbYesNoCancel, "Unsaved Data")
                Case DialogResult.Yes
                    Weektomonth1.updateRow(DataRow)
                Case DialogResult.Cancel
                    e.Cancel = False
            End Select
        End If
    End Sub



End Class
