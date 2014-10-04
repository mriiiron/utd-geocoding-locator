Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

Public Class BtnClearMarkers
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pGraphicsContainer As IGraphicsContainer = pMap
        pGraphicsContainer.DeleteAllElements()
        pMxDoc.ActiveView.Refresh()
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
