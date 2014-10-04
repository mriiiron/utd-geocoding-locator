Imports ESRI.ArcGIS.Location
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Display

Public Class BtnGeocodeSingle
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Public Shared Sub GeocodeSingle(ByVal Address As String, ByRef AddressGeocodingLocator As IAddressGeocoding, Optional ByVal IsShowMessage As Boolean = True)

        ' Set input address properties
        ' All hardcoded (UTD is completely in Richardson, TX 75080) except the detailed address
        Dim pAddressProperties As IPropertySet = New PropertySet
        With pAddressProperties
            .SetProperty("Street", Address)
            .SetProperty("City", "Richardson")
            .SetProperty("State", "TX")
            .SetProperty("ZIP", "75080")
        End With

        ' Try geocoding
        Dim pMatchProperties As IPropertySet
        Try
            pMatchProperties = AddressGeocodingLocator.MatchAddress(pAddressProperties)
        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        ' Show results
        Dim pMatchPoint As IPoint = Nothing
        Dim pMatchField As IField = Nothing
        Dim pMatchFields As IFields = AddressGeocodingLocator.MatchFields
        Dim strResult As String = ""
        For i As Long = 0 To pMatchFields.FieldCount - 1
            pMatchField = pMatchFields.Field(i)
            If pMatchField.Type = esriFieldType.esriFieldTypeGeometry Then
                pMatchPoint = pMatchProperties.GetProperty(pMatchField.Name)
                If Not pMatchPoint.IsEmpty Then
                    strResult = strResult & "Find Match: " & Address & vbCrLf & "X: " & pMatchPoint.X & vbCrLf & "Y: " & pMatchPoint.Y
                Else
                    If IsShowMessage Then
                        Windows.Forms.MessageBox.Show("Cannot find match: " & Address, "Single Match Result")
                    End If
                    Exit Sub
                End If
            Else
                strResult = strResult & pMatchField.AliasName & ": " & pMatchProperties.GetProperty(pMatchField.Name)
            End If
            If i < pMatchFields.FieldCount - 1 Then
                strResult = strResult & vbCrLf
            End If
        Next
        If IsShowMessage Then
            Windows.Forms.MessageBox.Show(strResult, "Single Match Result")
        End If

        ' Add a marker to show the result on map
        ' This would only applicable when geocoding success
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pGraphicsContainer As IGraphicsContainer = pMap
        Dim pResultElement As IElement = New MarkerElement
        pResultElement.Geometry = pMatchPoint
        Dim pMarkerSymbol As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        Dim pRedColor As IRgbColor = New RgbColor
        pRedColor.Red = 200
        pMarkerSymbol.Color = pRedColor
        pMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSDiamond
        Dim pMarkerElement As IMarkerElement = pResultElement
        pMarkerElement.Symbol = pMarkerSymbol
        pGraphicsContainer.AddElement(pMarkerElement, 0)
        pMxDoc.ActiveView.Refresh()

    End Sub

    Protected Overrides Sub OnClick()

        ' Get locator from extension
        Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
        If extUTDLocator.Locator Is Nothing Then
            Windows.Forms.MessageBox.Show("UTD Locator is not found." & vbCrLf & "Open or create it first.", "Error")
            Exit Sub
        End If

        ' Input a single address
        Dim strAddress As String = InputBox("Input your address:", "Geocode Single Address", "800 W Campbell Rd")

        ' Switch to geocoding interface
        Dim pAddressGeocodingLocator As IAddressGeocoding = extUTDLocator.Locator
        pAddressGeocodingLocator.Validate()

        ' Clear markers
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pGraphicsContainer As IGraphicsContainer = pMap
        pGraphicsContainer.DeleteAllElements()

        ' Geocode
        GeocodeSingle(strAddress, pAddressGeocodingLocator)

    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
