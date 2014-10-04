Imports ESRI.ArcGIS.Location
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

Public Class BtnGetLocator
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()

        Dim pFolderBrowserDialog As New Windows.Forms.FolderBrowserDialog()
        pFolderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer
        pFolderBrowserDialog.Description = "Select the folder which contains utd_road.shp or UTD Locator (If applicable)."
        pFolderBrowserDialog.ShowNewFolderButton = False

        Dim pObj As System.Object = Activator.CreateInstance(Type.GetTypeFromProgID("esriLocation.LocatorManager"))
        Dim pLocatorManager As ILocatorManager2 = TryCast(pObj, ILocatorManager2)

        If pFolderBrowserDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim pLocatorWorkspace As ILocatorWorkspace = pLocatorManager.GetLocatorWorkspaceFromPath(pFolderBrowserDialog.SelectedPath)

            Dim isGonnaLoadLocator As Boolean = False
            Dim pLocator As ILocator = Nothing
            Try

                pLocator = pLocatorWorkspace.GetLocator("UTD Locator")
                Windows.Forms.MessageBox.Show("Seems that UTD Locator has been created already." & vbCrLf & "UTD Locator opening success.", "Info")

                isGonnaLoadLocator = True

            Catch ex1 As Exception

                If Windows.Forms.MessageBox.Show("UTD Locator not found. Create one?" & vbCrLf & "The reference shapefile must be in the folder you just selected.", "Info", Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then

                    ' Open the default local locator workspace to get the locator style.
                    pLocatorWorkspace = pLocatorManager.GetLocatorWorkspaceFromPath("")
                    Dim pLocatorStyle As ILocatorStyle = pLocatorWorkspace.GetLocatorStyle("US Address - Dual Ranges")

                    ' Open the utd_road feature class to use as reference data.
                    Dim pWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactory
                    Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(pFolderBrowserDialog.SelectedPath, 0)

                    Dim pUTDFeatureClass As IFeatureClass
                    Try
                        pUTDFeatureClass = pFeatureWorkspace.OpenFeatureClass("utd_road")
                    Catch ex2 As Exception
                        Windows.Forms.MessageBox.Show("Cannot found required reference data (must be named as utd_road) in the selected folder.", "Error")
                        Exit Sub
                    End Try

                    ' Set the feature class as the primary reference data table for the locator
                    ' Then set the reference data table fields
                    Dim pUTDDataset As IDataset = CType(pUTDFeatureClass, IDataset)
                    Dim pReferenceDataTables As IReferenceDataTables = CType(pLocatorStyle, IReferenceDataTables)
                    Dim pEnumReferenceDataTable As IEnumReferenceDataTable = pReferenceDataTables.Tables
                    pEnumReferenceDataTable.Reset()
                    Dim pReferenceDataTable As IReferenceDataTable = pEnumReferenceDataTable.Next()
                    Dim pEnumReferenceDataField As IEnumReferenceDataField = pReferenceDataTable.Fields
                    pEnumReferenceDataField.Reset()
                    Dim pReferenceDataFieldEdit As IReferenceDataFieldEdit = pEnumReferenceDataField.Next()
                    Do While pReferenceDataFieldEdit IsNot Nothing
                        Select Case pReferenceDataFieldEdit.InternalName
                            Case "Primary.PreType"
                                pReferenceDataFieldEdit.Name_2 = "PRE_TYPE"
                            Case "Primary.PreDir"
                                pReferenceDataFieldEdit.Name_2 = "PRE_DIR"
                            Case "Primary.StreetName"
                                pReferenceDataFieldEdit.Name_2 = "NAME"
                            Case "Primary.SufType"
                                pReferenceDataFieldEdit.Name_2 = "SUF_TYPE"
                            Case "Primary.SufDir"
                                pReferenceDataFieldEdit.Name_2 = "SUF_DIR"
                            Case "Primary.FromLeft"
                                pReferenceDataFieldEdit.Name_2 = "L_F_ADD"
                            Case "Primary.ToLeft"
                                pReferenceDataFieldEdit.Name_2 = "L_T_ADD"
                            Case "Primary.FromRight"
                                pReferenceDataFieldEdit.Name_2 = "R_F_ADD"
                            Case "Primary.ToRight"
                                pReferenceDataFieldEdit.Name_2 = "R_T_ADD"
                            Case "Primary.CityLeft"
                                pReferenceDataFieldEdit.Name_2 = "CITY_L"
                            Case "Primary.CityRight"
                                pReferenceDataFieldEdit.Name_2 = "CITY_R"
                            Case "Primary.ZIPLeft"
                                pReferenceDataFieldEdit.Name_2 = "ZIP_L"
                            Case "Primary.ZIPRight"
                                pReferenceDataFieldEdit.Name_2 = "ZIP_R"
                        End Select
                        pReferenceDataFieldEdit = pEnumReferenceDataField.Next()
                    Loop
                    Dim pReferenceDataTableEdit As IReferenceDataTableEdit = CType(pReferenceDataTable, IReferenceDataTableEdit)
                    Dim pName As IName = pUTDDataset.FullName
                    pReferenceDataTableEdit.Name_2 = CType(pName, ITableName)

                    ' Store the new locator in the same workspace as the reference data.
                    If pReferenceDataTables.HasEnoughInfo Then
                        pLocatorWorkspace = pLocatorManager.GetLocatorWorkspaceFromPath(pFolderBrowserDialog.SelectedPath)
                        pLocator = pLocatorWorkspace.AddLocator("UTD Locator", CType(pLocatorStyle, ILocator), "", Nothing)
                        Windows.Forms.MessageBox.Show("UTD Locator created at:" & vbCrLf & pFolderBrowserDialog.SelectedPath, "Info")
                        isGonnaLoadLocator = True
                    Else
                        Windows.Forms.MessageBox.Show("Reference data is not enough. Locator creating failed.", "Error")
                    End If

                End If

            End Try

            ' If the locator already exists, or successfully created, load it to this add-in's extension
            If isGonnaLoadLocator Then

                ' Load UTD Locator to extension
                Dim extUTDLocator As ExtUTDLocatorData = ExtUTDLocatorData.GetExtension()
                extUTDLocator.WorkspacePath = pFolderBrowserDialog.SelectedPath
                extUTDLocator.Locator = pLocator

                ' Change the "status bar"
                Dim cmbStatus As CmbLocatorStatus = AddIn.FromID(Of CmbLocatorStatus)(My.ThisAddIn.IDs.CmbLocatorStatus)
                cmbStatus.ChangeStatus(cmbStatus.intEnabled)

                ' If UTD road map does not exist in current map, add it
                Dim pMxDoc As IMxDocument = My.ArcMap.Document
                Dim pMap As IMap = pMxDoc.FocusMap
                Dim pWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactory
                Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(pFolderBrowserDialog.SelectedPath, 0)
                Dim pUTDFeatureClass As IFeatureClass = pFeatureWorkspace.OpenFeatureClass("utd_road")
                Dim isNotFound As Boolean = True
                Dim pEnumLayer As IEnumLayer = pMap.Layers
                pEnumLayer.Reset()
                Dim pLayer As ILayer = pEnumLayer.Next()
                Do While Not (pLayer Is Nothing)
                    If pLayer.Name = "utd_road" Then
                        isNotFound = False
                        Exit Do
                    End If
                    pLayer = pEnumLayer.Next()
                Loop
                If isNotFound Then
                    Dim pUTDFeatureLayer As IFeatureLayer = New FeatureLayer
                    pUTDFeatureLayer.FeatureClass = pUTDFeatureClass
                    pUTDFeatureLayer.Name = pUTDFeatureClass.AliasName
                    pMap.AddLayer(pUTDFeatureLayer)
                End If

            End If

        End If


    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

End Class
