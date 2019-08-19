Imports System.Windows.Forms
Imports System.Drawing

Public Class DataGridViewRolloverCellColumn
    Inherits DataGridViewColumn

    Public Sub New()
        Me.CellTemplate = New DataGridViewRolloverCell

    End Sub
End Class
Public Class DataGridViewRolloverCell
    Inherits DataGridViewTextBoxCell

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)

        Dim cursorPosition As Point = Me.DataGridView.PointToClient(Cursor.Position)
        If cellBounds.Contains(cursorPosition) Then
            Dim newRect As New Rectangle(cellBounds.X + 1, cellBounds.Y + 1,
                                         cellBounds.Width - 4, cellBounds.Height - 4)
            graphics.DrawRectangle(Pens.Red, newRect)
        End If
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal rowIndex As Integer)
        'MyBase.OnMouseEnter(rowIndex)
        Me.DataGridView.InvalidateCell(Me)
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal rowIndex As Integer)
        'MyBase.OnMouseLeave(rowIndex)
        Me.DataGridView.InvalidateCell(Me)
    End Sub
End Class