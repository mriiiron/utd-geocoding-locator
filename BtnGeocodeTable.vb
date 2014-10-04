Imports ESRI.ArcGIS.Location
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesFile

Public Class BtnGeocodeTable
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    ' Using locator's MatchTable() method
    Public Shared Sub GeocodeTable_MatchTable(ByVal AddressTable As ITable, ByRef AddressGeocodingLocator As IAddressGeocoding)

        ' Create a feature class to contain the geocoding results
        ' The feature class must contain an ObjectID Field
        Dim pFieldsEdit As IFieldsEdit = New Fields
        Dim pFieldEdit As IFieldEdit = New Field
        pFieldEdit.Name_2 = "OBJECTID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        pFieldsEdit.AddField(pFieldEdit)

        ' Add the match fields
        Dim pMatchFields As IFields = AddressGeocodingLocator.MatchFields
        Dim intFieldCount As Integer = pMatchFields.FieldCount
        Dim strOutputFieldNames As String = ""
        For i As Integer = 0 To intFieldCount - 1
            pFieldsEdit.AddField(pMatchFields.Field(i))
            If (strOutputFieldNames.Length <> 0) Then
                strOutputFieldNames = strOutputFieldNames & ","
            End If
            strOutputFieldNames = strOutputFieldNames & pMatchFields.Field(i).Name
        Next i

        ' Add fields that will be copied over during geocoding
        ' In this case, we only copy address and city fields
        Dim pAddressTableFields As IFields = AddressTable.Fields
        Dim pFieldsToCopy As IPropertySet = New PropertySet
        intFieldCount = pAddressTableFields.FieldCount
        For i As Integer = 0 To intFieldCount - 1
            Dim pField As IField = pAddressTableFields.Field(i)
            If pField.Type <> esriFieldType.esriFieldTypeOID Then
                pFieldsEdit.AddField(pField)
                pFieldsToCopy.SetProperty(pField.Name, pField.Name)
            End If
        Next i

        ' Create the new feature class that will contain all of the geocoded addresses.
        Dim pUID As New UID
        pUID.Value = "esriGeodatabase.Feature"
        Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
        Dim pWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactory
        Dim pWorkspace As IWorkspace = pWorkspaceFactory.OpenFromFile(extUTDLocator.WorkspacePath, 0)
        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim pFeatureClassMatches As IFeatureClass = pFeatureWorkspace.CreateFeatureClass("Matched_Locations", pFieldsEdit, pUID, Nothing, esriFeatureType.esriFTSimple, "Shape", "")

        ' Set the fields which would composite an address
        ' City must be inculded here, e.g. 800 W Campbell Rd Richardson
        Dim strAddressFieldNames As String = "Address,City"

        ' Geocode the table into the feature class.
        AddressGeocodingLocator.MatchTable(AddressTable, strAddressFieldNames, "", pFeatureClassMatches, strOutputFieldNames, pFieldsToCopy, Nothing)

        ' Add matched points to map
        Dim pFeatureLayerMatches As IFeatureLayer = New FeatureLayer
        pFeatureLayerMatches.FeatureClass = pFeatureClassMatches
        pFeatureLayerMatches.Name = pFeatureClassMatches.AliasName
        Dim pMap As IMap = My.ArcMap.Document.FocusMap
        pMap.AddLayer(pFeatureLayerMatches)

    End Sub

    ' Use geocoding method from BtnGeocodeSingle multiple times
    Public Shared Sub GeocodeTable_IterateSingle(ByVal AddressTable As ITable, ByRef AddressGeocodingLocator As IAddressGeocoding)

        ' Remove markers
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pGraphicsContainer As IGraphicsContainer = pMap
        pGraphicsContainer.DeleteAllElements()

        ' Get the field index of Address
        ' City field not needed here because it has been hardcoded
        Dim pFields As IFields = AddressTable.Fields
        Dim intPosAddress As Integer = pFields.FindField("Address")

        ' Loop through the address table
        Dim pCursor As ICursor = AddressTable.Search(Nothing, True)
        Dim pRow As IRow = pCursor.NextRow()
        Dim strAddress As String
        Do Until pRow Is Nothing
            strAddress = pRow.Value(intPosAddress)
            BtnGeocodeSingle.GeocodeSingle(strAddress, AddressGeocodingLocator, False)
            pRow = pCursor.NextRow()
        Loop

    End Sub

    Protected Overrides Sub OnClick()

        ' Get locator from extension
        Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
        If extUTDLocator.Locator Is Nothing Then
            Windows.Forms.MessageBox.Show("UTD Locator is not found." & vbCrLf & "Open or create it first.", "Error")
            Exit Sub
        End If

        ' Switch to geocoding interface
        Dim pAddressGeocodingLocator As IAddressGeocoding = extUTDLocator.Locator
        pAddressGeocodingLocator.Validate()

        ' Open the table to geocode.
        ' User must select a standalone table in Table of Contents
        Dim pMxDoc As IMxDocument = My.ArcMap.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pTable As ITable
        Try
            pTable = pMxDoc.SelectedItem
            If TypeOf pTable Is IStandaloneTable Then
                If Windows.Forms.MessageBox.Show("Do you want to save matches to feature class?", "Info", Windows.Forms.MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    GeocodeTable_MatchTable(pTable, pAddressGeocodingLocator)
                Else
                    GeocodeTable_IterateSingle(pTable, pAddressGeocodingLocator)
                End If
            Else
                Windows.Forms.MessageBox.Show("You must selected a standalone table", "Error")
            End If
        Catch ex As Exception
            Windows.Forms.MessageBox.Show("You must selected a standalone table", "Error")
            Exit Sub
        End Try

    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
