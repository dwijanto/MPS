Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class ToolStripCustom
    Inherits System.Windows.Forms.ToolStrip

    'Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    '    'MyBase.OnPaintBackground(e)
    '    Dim formGraphics As System.Drawing.Graphics = e.Graphics
    '    Dim myLineargradienMode As New System.Drawing.Drawing2D.LinearGradientMode
    '    myLineargradienMode = Drawing2D.LinearGradientMode.Vertical
    '    Dim myrectangle = New Rectangle(Me.Location.X, Me.Location.Y, Width, Height)
    '    Dim gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(myrectangle, Color.AliceBlue, Color.Crimson, Drawing2D.LinearGradientMode.Vertical)


    '    formGraphics.FillRectangle(gradientBrush, ClientRectangle)
    'End Sub

    Public Sub New()
        MyBase.new()
        ' This call is required by the designer.
        InitializeComponent()
        Me.RenderMode = ToolStripRenderMode.Professional
        Me.Renderer = New ToolStripProfessionalRenderer(New CustomToolStripColorTable)
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    <Category("Style")>
    <DisplayName("ToolStripForeColor")>
    Public Property ToolStripForeColor As Color
        Get
            Return My.MySettings.Default.ToolStripForeColor
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripForeColor = value
            Me.ForeColor = value
        End Set
    End Property

    <Category("Style")>
    <DisplayName("ToolStripBorder")>
    Public Property ToolStripBorder As Color
        Get
            Return My.MySettings.Default.ToolStripBorder
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripBorder = value
        End Set
    End Property

    <Category("Style")>
    <DisplayName("ToolStripContentPanelGradientBegin")>
    Public Property ToolStripContentPanelGradientBegin As Color
        Get
            Return My.MySettings.Default.ToolStripContentPanelGradientBegin
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripContentPanelGradientBegin = value
        End Set
    End Property
    <Category("Style")>
    <DisplayName("ToolStripContentPanelGradientEnd")>
    Public Property ToolStripContentPanelGradientEnd As Color
        Get
            Return My.MySettings.Default.ToolStripContentPanelGradientEnd
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripContentPanelGradientEnd = value
        End Set
    End Property

    <Category("Style")>
    <DisplayName("ToolStripDropDownBackground")>
    Public Property ToolStripDropDownBackground As Color
        Get
            Return My.MySettings.Default.ToolStripDropDownBackground
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripDropDownBackground = value
        End Set
    End Property

    <Category("Style")>
    <DisplayName("ToolStripGradientBegin")>
    Public Property ToolStripGradientBegin As Color
        Get
            Return My.MySettings.Default.ToolStripGradientBegin
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripGradientBegin = value
        End Set
    End Property

    <Category("Style")>
    <DisplayName("ToolStripGradientEnd")>
    Public Property ToolStripGradientEnd As Color
        Get
            Return My.MySettings.Default.ToolStripGradientEnd
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripGradientEnd = value
        End Set
    End Property

    <Category("Style")>
   <DisplayName("ToolStripGradientMiddle")>
    Public Property ToolStripGradientMiddle As Color
        Get
            Return My.MySettings.Default.ToolStripGradientMiddle
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripGradientMiddle = value
        End Set
    End Property

    '<Category("Style")>
    '<DisplayName("ToolStripPanelGradientBegin")>
    'Public Property ToolStripPanelGradientBegin As Color
    '    Get
    '        Return My.MySettings.Default.ToolStripPanelGradientBegin
    '    End Get
    '    Set(ByVal value As Color)
    '        My.MySettings.Default.ToolStripPanelGradientBegin = value
    '    End Set
    'End Property


    '<Category("Style")>
    '<DisplayName("ToolStripPanelGradientEnd")>
    'Public Property ToolStripPanelGradientEnd As Color
    '    Get
    '        Return My.MySettings.Default.ToolStripPanelGradientEnd
    '    End Get
    '    Set(ByVal value As Color)
    '        My.MySettings.Default.ToolStripPanelGradientEnd = value
    '    End Set
    'End Property
End Class

Friend Class CustomToolStripColorTable
    Inherits ProfessionalColorTable


    Public Overrides ReadOnly Property ToolStripBorder As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripBorder


        End Get
    End Property

    Public Overrides ReadOnly Property ToolStripContentPanelGradientBegin As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripContentPanelGradientBegin
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripContentPanelGradientEnd As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripContentPanelGradientEnd
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripDropDownBackground As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripDropDownBackground
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripGradientBegin As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripGradientBegin
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripGradientEnd As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripGradientEnd
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripGradientMiddle As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripGradientMiddle
        End Get
    End Property
    'Public Overrides ReadOnly Property ToolStripPanelGradientBegin As System.Drawing.Color
    '    Get
    '        Return My.MySettings.Default.ToolStripPanelGradientBegin
    '    End Get
    'End Property
    'Public Overrides ReadOnly Property ToolStripPanelGradientEnd As System.Drawing.Color
    '    Get
    '        Return My.MySettings.Default.ToolStripPanelGradientEnd
    '    End Get
    'End Property

End Class