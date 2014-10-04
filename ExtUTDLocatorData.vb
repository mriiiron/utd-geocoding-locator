Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase

Public Class ExtUTDLocatorData
    Inherits ESRI.ArcGIS.Desktop.AddIns.Extension

    Private Shared extUTDLocator As ExtUTDLocatorData

    Public WorkspacePath As String = Nothing
    Public Locator As ILocator = Nothing

    Public Sub New()
        extUTDLocator = Me
    End Sub

    Friend Shared Function GetExtension() As ExtUTDLocatorData
        If extUTDLocator Is Nothing Then
            Dim extID As UID = New UIDClass()
            extID.Value = My.ThisAddIn.IDs.ExtUTDLocatorData
            My.ArcMap.Application.FindExtensionByCLSID(extID)
        End If
        Return extUTDLocator
    End Function

End Class
