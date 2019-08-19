Imports System.Windows.Forms

Class CustomToolStripContainerColorTable
    Inherits ProfessionalColorTable

    Public Overrides ReadOnly Property ToolStripPanelGradientBegin As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripPanelGradientBegin
        End Get
    End Property
    Public Overrides ReadOnly Property ToolStripPanelGradientEnd As System.Drawing.Color
        Get
            Return My.MySettings.Default.ToolStripPanelGradientEnd
        End Get
    End Property
End Class