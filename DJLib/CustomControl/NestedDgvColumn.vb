Imports System.Windows.Forms
Imports System.Drawing

Public Class NestedDgvColumn
    Inherits DataGridViewColumn

    Public Sub New()
        MyBase.New(New NestedDgvCell)
    End Sub

    Public Overrides Property CellTemplate As System.Windows.Forms.DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridViewCell)
            If (value IsNot Nothing) AndAlso Not value.GetType().IsAssignableFrom(GetType(NestedDgvCell)) Then
                Throw New InvalidCastException("Must be NestedDgvCell")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property
End Class

Public Class NestedDgvCell
    Inherits DataGridViewCell

    Private dgv As New DataGridView

    Private Sub setupDGVToDraw()
        dgv.BackgroundColor = SystemColors.Control
        dgv.Size = New Size(400, 100)
        dgv.AllowUserToAddRows = False
        dgv.RowHeadersVisible = False
        dgv.ColumnHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.BorderStyle = Windows.Forms.BorderStyle.None
        dgv.AutoGenerateColumns = False

        If TypeOf (Value) Is DataTable Then
            Dim table As DataTable = DirectCast(Value, DataTable)
            dgv.Columns.Clear()
            Dim dgvColumn As DataGridViewTextBoxColumn
            For Each column As DataColumn In table.Columns
                dgvColumn = New DataGridViewTextBoxColumn
                dgvColumn.Name = column.ColumnName
                dgvColumn.HeaderText = column.ColumnName
                dgvColumn.DefaultCellStyle.Format = "C2"
                dgvColumn.ValueType = GetType(Decimal)
                dgv.Columns.Add(dgvColumn)
            Next

            For Each DataRow As DataRow In CType(Value, DataTable).Rows
                dgv.Rows.Add(DataRow.ItemArray)
            Next
        End If
    End Sub

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)

        setupDGVToDraw()
        Dim abbreviation As New Bitmap(cellBounds.Width, cellBounds.Height)
        dgv.DrawToBitmap(abbreviation, New Rectangle(0, 0, cellBounds.Width, cellBounds.Height))
        graphics.DrawImage(abbreviation, cellBounds, New Rectangle(0, 0, abbreviation.Width, abbreviation.Height), GraphicsUnit.Pixel)
    End Sub

    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle)
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)
    End Sub

    Public Overrides ReadOnly Property EditType As System.Type
        Get
            'Return MyBase.EditType
            Return GetType(DgvEditingControl)
        End Get
    End Property

    Public Overrides Property ValueType As System.Type
        Get
            Return MyBase.ValueType
        End Get
        Set(ByVal value As System.Type)
            'MyBase.ValueType = value
        End Set
    End Property
End Class

Class DgvEditingControl
    Inherits DataGridView
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer

    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle) Implements System.Windows.Forms.IDataGridViewEditingControl.ApplyCellStyleToEditingControl
        Me.Font = dataGridViewCellStyle.Font
        Me.ForeColor = dataGridViewCellStyle.ForeColor
        Me.BackgroundColor = dataGridViewCellStyle.BackColor
    End Sub

    Public Property EditingControlDataGridView As System.Windows.Forms.DataGridView Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlDataGridView
        Get
            Return dataGridViewControl
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridView)
            dataGridViewControl = value
        End Set
    End Property

    Public Property EditingControlFormattedValue As Object Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlFormattedValue
        Get
            Return Me.RowCount
        End Get
        Set(ByVal value As Object)

        End Set
    End Property

    Public Property EditingControlRowIndex As Integer Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get
        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Public Property EditingControlValueChanged As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get
        Set(ByVal value As Boolean)
            valueIsChanged = True
        End Set
    End Property

    Public Function EditingControlWantsInputKey(ByVal keyData As System.Windows.Forms.Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlWantsInputKey
        Select Case keyData And Keys.KeyCode
            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, Keys.Enter, Keys.Escape, Keys.Tab
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public ReadOnly Property EditingPanelCursor As System.Windows.Forms.Cursor Implements System.Windows.Forms.IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.DefaultCursor
        End Get
    End Property

    Public Function GetEditingControlFormattedValue(ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object Implements System.Windows.Forms.IDataGridViewEditingControl.GetEditingControlFormattedValue
        Return Me.RowCount
    End Function

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) Implements System.Windows.Forms.IDataGridViewEditingControl.PrepareEditingControlForEdit
        'no preparation needs to be done
    End Sub

    Public ReadOnly Property RepositionEditingControlOnValueChange As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.RepositionEditingControlOnValueChange
        Get
            Return False
        End Get
    End Property
End Class