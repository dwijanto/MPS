Imports System.Runtime.InteropServices

Public Class PublicClass
    Public Shared dbtools1 As New DJLib.Dbtools
    Public Shared HelperClass1 As New SSP.HelperClass
    Public Shared DBAdapter1 As DBAdapter
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean        
    End Function
End Class
