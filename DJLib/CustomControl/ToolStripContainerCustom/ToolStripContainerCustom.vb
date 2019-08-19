
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class ToolStripContainerCustom
    Inherits System.Windows.Forms.ToolStripContainer

    Public Sub New()
        Me.LeftToolStripPanel.RenderMode = ToolStripRenderMode.Professional
        Me.LeftToolStripPanel.Renderer = New ToolStripProfessionalRenderer(New CustomToolStripContainerColorTable)
        Me.TopToolStripPanel.RenderMode = ToolStripRenderMode.Professional
        Me.TopToolStripPanel.Renderer = New ToolStripProfessionalRenderer(New CustomToolStripContainerColorTable)
        Me.RightToolStripPanel.RenderMode = ToolStripRenderMode.Professional
        Me.RightToolStripPanel.Renderer = New ToolStripProfessionalRenderer(New CustomToolStripContainerColorTable)
        Me.BottomToolStripPanel.RenderMode = ToolStripRenderMode.Professional
        Me.BottomToolStripPanel.Renderer = New ToolStripProfessionalRenderer(New CustomToolStripContainerColorTable)
    End Sub

    Public Property ToolStripPanelGradientBegin As Color
        Get
            Return My.MySettings.Default.ToolStripPanelGradientBegin
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripPanelGradientBegin = value
            Me.LeftToolStripPanel.BackColor = value
        End Set
    End Property
    Public Property ToolStripPanelGradientEnd As Color
        Get
            Return My.MySettings.Default.ToolStripPanelGradientEnd
        End Get
        Set(ByVal value As Color)
            My.MySettings.Default.ToolStripPanelGradientEnd = value
        End Set
    End Property

End Class



