Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Location
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase

Public Class TolReverseGeocodeSingle
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Public Sub New()

    End Sub

    Public Shared Sub ReverseGeocodeSingle(ByVal Point As IPoint, ByRef ReverseGeocodingLocator As IReverseGeocoding)

        ' Set tolerance, hardcoded to 100 feet
        Dim pReverseGeocodingProperties As IReverseGeocodingProperties = ReverseGeocodingLocator
        pReverseGeocodingProperties.SearchDistance = 100
        pReverseGeocodingProperties.SearchDistanceUnits = esriUnits.esriFeet

        ' Try reverse geocoding
        Dim pAddressProperties As IPropertySet
        Try
            pAddressProperties = ReverseGeocodingLocator.ReverseGeocode(Point, False)
        Catch ex As Exception
            Windows.Forms.MessageBox.Show("Cannot find match at that point.")
            Exit Sub
        End Try

        ' Show results
        Dim pAddressInputs As IAddressInputs = ReverseGeocodingLocator
        Dim pAddressFields As IFields = pAddressInputs.AddressFields
        Dim strResult As String = ""
        For i As Integer = 0 To pAddressFields.FieldCount - 1
            Dim pAddressField As IField = pAddressFields.Field(i)
            strResult = strResult & pAddressField.AliasName & ": " + pAddressProperties.GetProperty(pAddressField.Name)
            If i < pAddressFields.FieldCount - 1 Then
                strResult = strResult & vbCrLf
            End If
        Next
        Windows.Forms.MessageBox.Show(strResult, "Single Address Result")

    End Sub

    Protected Overrides Sub OnActivate()

        ' Check if locator exists
        ' If not, deactivate the tool (not like buttons, tools should be deactivated)
        Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
        If extUTDLocator.Locator Is Nothing Then
            Windows.Forms.MessageBox.Show("UTD Locator is not found." & vbCrLf & "Open or create it first.", "Error")
            My.ArcMap.Application.CurrentTool = Nothing
            Exit Sub
        End If

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)

        ' Get the point where user clicks and convert it to map coordinates
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pActiveView As IActiveView = TryCast(pMxDoc.FocusMap, IActiveView)
        Dim pPoint As IPoint = TryCast(pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(arg.X, arg.Y), IPoint)

        ' Switch to reverse geocoding interface
        Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
        Dim pReverseGeocodingLocator As IReverseGeocoding = extUTDLocator.Locator

        ' Reverse geocode
        ReverseGeocodeSingle(pPoint, pReverseGeocodingLocator)

        ' Deactivate the tool
        My.ArcMap.Application.CurrentTool = Nothing

    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
