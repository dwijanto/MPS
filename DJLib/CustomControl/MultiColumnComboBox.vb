Imports System.Windows.Forms
Imports System.Drawing

Public Class MultiColumnComboBox
    Inherits ComboBox


    Public Property FieldSeparator As Char
    Public Property SelectedIdValue As Object


    Public Sub New()
        Me.DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
        'm_chFieldSeparator = ","c
        _FieldSeparator = IIf(FieldSeparator = "", ",", FieldSeparator)
    End Sub

    Protected Overrides Sub OnDrawItem(ByVal e As System.Windows.Forms.DrawItemEventArgs)
        Dim iXPos As Integer
        Dim iYPos As Integer
        Dim sText As String
        Dim objSizeF As SizeF
        Dim sTextPartArray() As String
        Dim iTextPartIndex As Integer
        Dim sTextPart As String
        Dim objBrush As Brush
        Dim iMaxLength As Integer

        If e.Index >= 0 Then
            e.DrawBackground()
            sText = Me.Items(e.Index).ToString
            sTextPartArray = sText.Split(_FieldSeparator)
            objBrush = New SolidBrush(e.ForeColor)
            iXPos = e.Bounds.X
            For iTextPartIndex = 0 To sTextPartArray.Length - 1
                sTextPart = sTextPartArray(iTextPartIndex)
                objSizeF = e.Graphics.MeasureString(sTextPart, e.Font)
                If iTextPartIndex > 0 Then
                    iMaxLength = getMaxLength(e.Graphics, e.Font, iTextPartIndex - 1, _FieldSeparator)
                Else
                    iMaxLength = 0
                End If
                iXPos += iMaxLength
                iYPos = e.Bounds.Y
                e.Graphics.DrawString(sTextPart, e.Font, objBrush, iXPos, iYPos)
            Next
            objBrush.Dispose()
            Select Case e.State
                Case DrawItemState.NoFocusRect
                Case Else
                    e.DrawFocusRectangle()
            End Select
        Else
            MyBase.OnDrawItem(e)
        End If
    End Sub

    Private Function getMaxLength(ByVal objgraphics As Graphics, ByVal objfont As Drawing.Font, ByVal iTextPartIndex As Integer, ByVal chSeparator As Char) As Integer
        Dim iResult As Integer
        Dim iIndex As Integer
        Dim sTextPartArray() As String
        Dim sTextPart As String
        Dim sText As String
        Dim objSizeF As SizeF
        Dim iWidth As Integer
        iResult = 10

        For iIndex = 0 To Me.Items.Count - 1
            sText = Me.Items(iIndex).ToString
            sTextPartArray = sText.Split(chSeparator)
            If iTextPartIndex > sTextPartArray.Length - 1 Then
            Else
                sTextPart = sTextPartArray(iTextPartIndex)
                objSizeF = objgraphics.MeasureString(sTextPart, objfont)
                iWidth = CType(objSizeF.Width, Integer)
                If iWidth > iResult Then
                    iResult = iWidth
                End If
            End If
        Next
        Return iResult
    End Function
End Class
