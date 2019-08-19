Imports System.Windows.Forms
Imports System.Drawing
Public Class DataGridViewBarGraphColumn
    Inherits DataGridViewColumn

    Public MaxValue As Long
    Private needsRecalc As Boolean = True

    Public Sub calcMaxValue()
        If needsRecalc Then
            Dim colIndex As Integer = Me.DisplayIndex
            For i = 0 To Me.DataGridView.Rows.Count - 1
                Dim row As DataGridViewRow = Me.DataGridView.Rows(i)
                MaxValue = Math.Max(MaxValue, CLng(IIf(IsDBNull(row.Cells(colIndex).Value), 0, row.Cells(colIndex).Value)))
            Next
            needsRecalc = False
        End If
    End Sub

    Public Sub New()
        Me.CellTemplate = New DataGridViewBarGraphCell
        Me.ReadOnly = True
    End Sub
End Class

Public Class DataGridViewBarGraphCell
    Inherits DataGridViewTextBoxCell
    Const HORIZONTALOFSET As Integer = 1
    Const SPACER As Integer = 4
    Const VERTOFFSET As Integer = 4
    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics,
                                  ByVal clipBounds As System.Drawing.Rectangle,
                                  ByVal cellBounds As System.Drawing.Rectangle,
                                  ByVal rowIndex As Integer,
                                  ByVal cellState As System.Windows.Forms.DataGridViewElementStates,
                                  ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String,
                                  ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle,
                                  ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle,
                                  ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        'MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, "", errorText, cellStyle, advancedBorderStyle, paintParts)
        Dim cellValue As Decimal
        If IsDBNull(value) Then
            cellValue = 0
        Else
            cellValue = CDec(value)
        End If
        If cellValue = 0 Then
            cellValue = 1
        End If

        Dim parent As DataGridViewBarGraphColumn = CType(Me.OwningColumn, DataGridViewBarGraphColumn)
        parent.calcMaxValue()
        Dim maxValue As Long = parent.MaxValue
        Dim fnt As Font = parent.InheritedStyle.Font
        Dim maxValueSize As SizeF = graphics.MeasureString(maxValue.ToString(), fnt)
        Dim availableWidth As Single = cellBounds.Width - maxValueSize.Width - SPACER - (HORIZONTALOFSET * 2)
        cellValue = Convert.ToDecimal(Convert.ToDouble(cellValue) / maxValue) * availableWidth
        Dim newRect As New RectangleF(cellBounds.X + HORIZONTALOFSET,
                                      cellBounds.Y + VERTOFFSET,
                                      cellValue, cellBounds.Height - (VERTOFFSET * 2))
        graphics.FillRectangle(Brushes.Blue, newRect)

        Dim cellText As String = formattedValue.ToString
        Dim textSize As SizeF = graphics.MeasureString(cellText, fnt)
        Dim textStart As PointF = New PointF(HORIZONTALOFSET + cellValue + SPACER, (cellBounds.Height - textSize.Height) / 2)
        Dim textColor As Color = parent.InheritedStyle.ForeColor
        If (cellState And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
            textColor = parent.InheritedStyle.SelectionForeColor
        End If

        Using Brush As New SolidBrush(textColor)
            graphics.DrawString(cellText, fnt, Brush, cellBounds.X + textStart.X, cellBounds.Y + textStart.Y)
        End Using
    End Sub
End Class
