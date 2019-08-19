Imports System.Windows.Forms
Imports System.Drawing

Public Class DynamicMenu
    Implements IDisposable

    Private ObjForm As Form
    Private mainMenu As New MenuStrip()
    Private objMenu As MenuStrip
    Private mDataTable1 As DataTable
    Private ImageList1 As ImageList
    Public Sub New(ByVal mObjForm As Form, ByVal DataTable As DataTable)
        ObjForm = mObjForm
        mDataTable1 = DataTable
    End Sub
    Public Sub New(ByVal mObjForm As Form, ByVal DataTable As DataTable, ByVal mImageList As ImageList)
        ObjForm = mObjForm
        mDataTable1 = DataTable
        ImageList1 = mImageList
    End Sub
    Public Function LoadMenu(ByRef mylistmenu As MenuStrip) As Boolean
        Dim mItem As New ToolStripMenuItem()
        Dim myReturn As Boolean = False
        Dim myParent As String = String.Empty
        Dim myErrorString As String = String.Empty
        Dim myEvent As String = String.Empty
        Dim DataRows1() As DataRow = mDataTable1.Select("parentid=0", "myorder asc")
        If DataRows1.Length > 0 Then
            For Each dr As DataRow In DataRows1
                mItem = New ToolStripMenuItem(Trim(dr.ItemArray(4).ToString))
                myParent = dr.ItemArray(1).ToString
                myEvent = Trim(dr.ItemArray(5).ToString)

                mainMenu.Items.Add(mItem)
                If myEvent <> "" Then
                    Call FindEventsByName(mItem, ObjForm, True, "MenuItemOn", "_" & myEvent)
                End If
                Call GenerateMenus(myParent, mItem)
            Next
            mylistmenu = objMenu
            myReturn = True
        End If
        Return myReturn
    End Function

    Private Function GenerateMenus(ByVal myParent As String, ByVal mItm As ToolStripMenuItem)
        Dim sMenu As New ToolStripMenuItem()
        Dim myReturn As Boolean = False
        Dim myEvent As String = String.Empty
        Dim DataRows1() As DataRow = mDataTable1.Select("parentid=" & myParent, "myorder asc")
        If DataRows1.Length > 0 Then
            For Each dr As DataRow In DataRows1
                If Trim(dr.ItemArray(4).ToString) = "-" Then
                    mItm.DropDownItems.Add(New ToolStripSeparator)
                Else
                    sMenu = New ToolStripMenuItem(Trim(dr.ItemArray(4).ToString))
                    If Not dr.ItemArray(6).ToString = "" Then
                        sMenu.Image = ImageList1.Images(CInt(dr.ItemArray(7))) 'Image.FromFile(Application.ExecutablePath & dr.ItemArray(6).ToString)
                    End If
                    mItm.DropDownItems.Add(sMenu)
                End If
                myParent = dr.ItemArray(1).ToString
                myEvent = Trim(dr.ItemArray(5).ToString)
                If myEvent <> "" Then
                    Call FindEventsByName(sMenu, ObjForm, True, "MenuItemOn", "_" & myEvent)
                End If
                Call GenerateMenus(myParent, sMenu)
            Next
            objMenu = mainMenu
            myReturn = True
        End If
        Return myReturn
    End Function


    Private Sub FindEventsByName(ByVal sender As Object, _
                                 ByVal receiver As Object, ByVal bind As Boolean, _
                                 ByVal handlerPrefix As String, ByVal handlerSufix As String)

        Dim SenderEvents() As System.Reflection.EventInfo = sender.GetType().GetEvents()
        Dim ReceiverType As Type = receiver.GetType()
        Dim E As System.Reflection.EventInfo
        Dim Method As System.Reflection.MethodInfo
        Dim mySW As New Stopwatch

        For Each E In SenderEvents
            'only interest in Click
            If E.Name = "Click" Then
                'mySW.sWStart()

                Method = ReceiverType.GetMethod(handlerPrefix & E.Name & handlerSufix, _
                                                System.Reflection.BindingFlags.IgnoreCase Or _
                                                System.Reflection.BindingFlags.Instance Or _
                                                System.Reflection.BindingFlags.NonPublic)
                'mySW.sWStop()


                If Not Method Is Nothing Then
                    Dim D As System.Delegate = System.Delegate.CreateDelegate(E.EventHandlerType, receiver, Method.Name)
                    If bind Then
                        E.AddEventHandler(sender, D)
                    Else
                        E.RemoveEventHandler(sender, D)
                    End If
                End If

            End If
        Next


    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
