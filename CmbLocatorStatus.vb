Public Class CmbLocatorStatus
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    Public intNotFound As Integer = 0
    Public intEnabled As Integer = 0

    Public Sub New()
        intNotFound = Me.Add("Not Found")
        intEnabled = Me.Add("Enabled")
        Me.Select(intNotFound)
        Me.Enabled = False
    End Sub

    Public Sub ChangeStatus(ByVal cookie As Integer)
        Me.Select(cookie)
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
