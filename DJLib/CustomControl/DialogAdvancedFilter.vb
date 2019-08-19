Imports System.Windows.Forms
Imports System.ComponentModel

Public Class DialogAdvancedFilter
    Private data As New BindingSource

    Private FilterOperatorHash As New Hashtable
    Private isDateTime As Boolean
    Private Dg As DataGridView
    Dim _myfilter As String

    Public ReadOnly Property myFilter As String
        Get
            Return _myfilter
        End Get
    End Property


    Public Sub New(ByVal DG As DataGridView)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        InitFilterOperatorHash()

        data = CType(DG.DataSource, BindingSource)
        
        Me.Dg = DG

        InitDataLayout()

    End Sub
    Private Sub InitFilterOperatorHash()
        FilterOperatorHash.Add(0, "None")
        FilterOperatorHash.Add(1, "=")
        FilterOperatorHash.Add(2, "Like")
        FilterOperatorHash.Add(3, "<")
        FilterOperatorHash.Add(4, "<=")
        FilterOperatorHash.Add(5, ">")
        FilterOperatorHash.Add(6, ">=")
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        executefilter()

        Me.Close()
    End Sub

    Private Sub executefilter()
        If data.List.Count <= 0 OrElse
            FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 OrElse
            OperatorComboBox.SelectedIndex <= 0 Then
            Return
        End If

        If String.IsNullOrEmpty(FilterTextBox.Text) Then Return

        'inFilterMode = True

        '##getpropertyname##
        '1.get columnname from combo
        Dim filterMember As String = CType(FieldComboBox.SelectedItem, FieldClass).id.ToString
        '2.Get dataitem from bindinglist.list(0)
        Dim DataItem As Object = data.List(0)
        '3.Get Propertiescollection from dataitem
        Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(DataItem)
        '4.Get Selected PropertyDescriptor based on filtermember
        Dim propDesc As PropertyDescriptor = props.Find(filterMember, True)

        'getoperator
        Dim stringoperator As String = FilterOperatorHash(OperatorComboBox.SelectedIndex).ToString
        'putbindingfilter
        'Check for different format
        Dim filterValue As String = Nothing
        Dim JoinFilter As String = "AND "
        If data.Filter <> "" AndAlso data.Filter IsNot Nothing Then
            If RadioButton2.Checked Then JoinFilter = "OR "
        Else
            JoinFilter = ""
        End If

        Select Case OperatorComboBox.SelectedIndex
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
                If isDateTime Then
                    filterValue = String.Format("{0}{1} {2} '#{3}#'", JoinFilter, propDesc.Name, stringoperator, FilterTextBox.Text)
                Else
                    filterValue = String.Format("{0}{1} {2} '{3}'", JoinFilter, propDesc.Name, stringoperator, FilterTextBox.Text)
                End If

        End Select
        _myfilter = filterValue

    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub InitDataLayout()
        BindFilterFields()
        BuildAutoCompleteString()
        OperatorComboBox.DataSource = System.Enum.GetNames(GetType(FilterOperator))
    End Sub

    Private Sub BindFilterFields()
        Dim cols As List(Of FieldClass) = New List(Of FieldClass)
        If Dg.Columns.Count > 0 Then

            For i = 0 To Dg.Columns.Count - 1
                If Dg.Columns(i).Visible Then
                    cols.Add(New FieldClass With {.id = Dg.Columns(i).Name,
                                                  .name = Dg.Columns(i).HeaderText})
                End If
            Next
            'cols.Sort()
            cols.Insert(0, New FieldClass With {.id = "None", .name = "None"})
        End If
        FieldComboBox.DataSource = cols
        FieldComboBox.DisplayMember = "Name"
        FieldComboBox.SelectedItem = "id"
    End Sub
    Private Sub BuildAutoCompleteString()
        Dim myfilter As String
        isDateTime = False
        'clear first
        FilterTextBox.AutoCompleteCustomSource.Clear()

        If data.List.Count <= 0 OrElse FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 Then Return

        'Get Column Name
        myfilter = data.Filter
        If RadioButton2.Checked Then
            data.Filter = ""
        End If
        
        Dim FilterField As String = CType(FieldComboBox.SelectedItem, FieldClass).id.ToString
        Dim filterVals As AutoCompleteStringCollection = New AutoCompleteStringCollection
        For Each dataitem As Object In data.List
            Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(dataitem)
            Dim propdesc As PropertyDescriptor = props.Find(FilterField, True)
            Dim fieldval As String = propdesc.GetValue(dataitem).ToString
            If propdesc.PropertyType.Name = "DateTime" Then
                isDateTime = True
            End If
            filterVals.Add(fieldval)
        Next
        data.Filter = myfilter
        FilterTextBox.AutoCompleteCustomSource = filterVals
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        FilterOperatorHash.Clear()
    End Sub

    Private Sub FieldComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FieldComboBox.SelectedIndexChanged
        BuildAutoCompleteString()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged
        BuildAutoCompleteString()
    End Sub
End Class
Public Enum FilterOperator
    None
    IsEqualTo
    IsLike
    IsLessThan
    isLessThanOrEqualTo
    IsGreaterThan
    isGreaterThanOrEqualTo
End Enum
Public Class FieldClass
    Public Property name
    Public Property id

    Public Overrides Function ToString() As String
        'Return MyBase.ToString()
        Return name.ToString
    End Function
End Class