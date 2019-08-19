Imports System.Windows.Forms
Imports System.Drawing
Public Class DataGridViewComboBoxMultiColumn
    Inherits DataGridViewColumn
    Friend DropDownDataSource As DataTable

    Public Sub New(ByVal Datasource As DataTable)
        MyBase.New(New DataGridViewMultiColumnComboBoxCell())
        DropDownDataSource = Datasource
    End Sub
   

    Public Overrides Property CellTemplate As System.Windows.Forms.DataGridViewCell
        Get

            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridViewCell)
            If (value IsNot Nothing) AndAlso
                Not value.GetType().IsAssignableFrom(GetType(DataGridViewMultiColumnComboBoxCell)) Then
                Throw New InvalidCastException("Must be a Date Cell")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property
End Class


Public Class DataGridViewMultiColumnComboBoxCell
    Inherits DataGridViewTextBoxCell

    Private DisplayValue As String = Nothing


    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle)
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)
        Dim ctl As MultiColumnComboBoxEditingControl = DirectCast(DataGridView.EditingControl, MultiColumnComboBoxEditingControl)
        Dim col As DataGridViewComboBoxMultiColumn = DirectCast(Me.OwningColumn, DataGridViewComboBoxMultiColumn)
        ctl.DataSource = col.DropDownDataSource
        ctl.OwnerCell = Me
    End Sub
    Friend Sub setDisplayValue(ByVal newValue As String)
        DisplayValue = newValue
    End Sub
    Public Overrides ReadOnly Property EditType As System.Type
        Get
            'Return MyBase.EditType
            Return GetType(MultiColumnComboBoxEditingControl)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType As System.Type
        Get
            'Return MyBase.ValueType
            Return GetType(Long)
        End Get
    End Property

    Public Overrides ReadOnly Property DefaultNewRowValue As Object
        Get
            'Return MyBase.DefaultNewRowValue
            Return DBNull.Value
        End Get
    End Property

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)

        'If DisplayValue is Nothing then setDisplayValue(l
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, DisplayValue, errorText, cellStyle, advancedBorderStyle, paintParts)

    End Sub

End Class

'Public Class MultiColumnComboBoxEditingControl
'    Inherits ListView
'    Implements IDataGridViewEditingControl

'    Private datagridviewcontrol As DataGridView
'    Private valueIsChanged As Boolean = False
'    Private rowIndexNum As Integer

'    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle) Implements System.Windows.Forms.IDataGridViewEditingControl.ApplyCellStyleToEditingControl
'        Me.Font = dataGridViewCellStyle.Font
'        Me.ForeColor = dataGridViewCellStyle.ForeColor
'        Me.BackColor = dataGridViewCellStyle.BackColor
'    End Sub

'    Public Property EditingControlDataGridView As System.Windows.Forms.DataGridView Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlDataGridView
'        Get
'            Return datagridviewcontrol
'        End Get
'        Set(ByVal value As System.Windows.Forms.DataGridView)
'            datagridviewcontrol = value
'        End Set
'    End Property

'    Public Property EditingControlFormattedValue As Object Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlFormattedValue
'        Get
'            Return ""
'        End Get
'        Set(ByVal value As Object)

'        End Set
'    End Property

'    Public Property EditingControlRowIndex As Integer Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlRowIndex
'        Get
'            Return rowIndexNum
'        End Get
'        Set(ByVal value As Integer)
'            rowIndexNum = value
'        End Set
'    End Property

'    Public Property EditingControlValueChanged As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlValueChanged
'        Get
'            Return valueIsChanged
'        End Get
'        Set(ByVal value As Boolean)
'            valueIsChanged = value
'        End Set
'    End Property

'    Public Function EditingControlWantsInputKey(ByVal keyData As System.Windows.Forms.Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlWantsInputKey
'        ' Let the DateTimePicker handle the keys listed.
'        Select Case keyData And Keys.KeyCode
'            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
'                Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

'                Return True

'            Case Else
'                Return Not dataGridViewWantsInputKey
'        End Select
'    End Function

'    Public ReadOnly Property EditingPanelCursor As System.Windows.Forms.Cursor Implements System.Windows.Forms.IDataGridViewEditingControl.EditingPanelCursor
'        Get
'            Return MyBase.Cursor
'        End Get
'    End Property

'    Public Function GetEditingControlFormattedValue(ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object Implements System.Windows.Forms.IDataGridViewEditingControl.GetEditingControlFormattedValue
'        Return Me.Items(0).Text
'    End Function

'    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) Implements System.Windows.Forms.IDataGridViewEditingControl.PrepareEditingControlForEdit

'    End Sub

'    Public ReadOnly Property RepositionEditingControlOnValueChange As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.RepositionEditingControlOnValueChange
'        Get
'            Return False
'        End Get
'    End Property
'End Class
Class MultiColumnComboBoxEditingControl
    Inherits MultiColumnComboBox
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer
    Private currentCell As DataGridViewMultiColumnComboBoxCell = Nothing

    Public Property OwnerCell() As DataGridViewMultiColumnComboBoxCell
        Get
            Return currentCell
        End Get
        Set(ByVal value As DataGridViewMultiColumnComboBoxCell)
            currentCell = Nothing
            MyBase.SelectedIdValue = value.Value
        End Set
    End Property



    Public Property EditingControlFormattedValue() As Object _
        Implements IDataGridViewEditingControl.EditingControlFormattedValue

        Get
            Return MyBase.SelectedIdValue
        End Get

        Set(ByVal value As Object)
            MyBase.SelectedIdValue = value            
        End Set

    End Property

    Public Function GetEditingControlFormattedValue(ByVal context _
        As DataGridViewDataErrorContexts) As Object _
        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

        Return MyBase.SelectedIdValue

    End Function

    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As  _
        DataGridViewCellStyle) _
        Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

        Me.Font = dataGridViewCellStyle.Font
        Me.ForeColor = dataGridViewCellStyle.ForeColor
        Me.BackColor = dataGridViewCellStyle.BackColor

    End Sub

    Public Property EditingControlRowIndex() As Integer _
        Implements IDataGridViewEditingControl.EditingControlRowIndex

        Get
            Return rowIndexNum
        End Get
        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set

    End Property

    Public Function EditingControlWantsInputKey(ByVal key As Keys, _
        ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
        Implements IDataGridViewEditingControl.EditingControlWantsInputKey

        ' Let the DateTimePicker handle the keys listed.
        Select Case key And Keys.KeyCode
            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
                Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

                Return True

            Case Else
                Return Not dataGridViewWantsInputKey
        End Select

    End Function

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
        Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

        ' No preparation needs to be done.

    End Sub

    Public ReadOnly Property RepositionEditingControlOnValueChange() _
        As Boolean Implements _
        IDataGridViewEditingControl.RepositionEditingControlOnValueChange

        Get
            Return False
        End Get

    End Property

    Public Property EditingControlDataGridView() As DataGridView _
        Implements IDataGridViewEditingControl.EditingControlDataGridView

        Get
            Return dataGridViewControl
        End Get
        Set(ByVal value As DataGridView)
            dataGridViewControl = value
        End Set

    End Property

    Public Property EditingControlValueChanged() As Boolean _
        Implements IDataGridViewEditingControl.EditingControlValueChanged

        Get
            Return valueIsChanged
        End Get
        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set

    End Property

    Public ReadOnly Property EditingControlCursor() As Cursor _
        Implements IDataGridViewEditingControl.EditingPanelCursor

        Get
            Return MyBase.Cursor
        End Get

    End Property

    Private Sub DoSelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SelectedValueChanged
        If currentCell IsNot Nothing Then
            valueIsChanged = True
            currentCell.Value = MyBase.SelectedIdValue
            currentCell.setDisplayValue(MyBase.Text)
        End If
    End Sub

End Class
