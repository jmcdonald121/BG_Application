Attribute VB_Name = "ODNR_Common"
'********************************************************************************
'
'   Program:    ODNR_Common
'   Author:     Gregory Palovchik
'               Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229 ext. 202
'   Date:       March 21, 2002
'   Purpose:    Provides common routines and functions to the geology extension
'
' This module provides global variables representing the file paths to data
' necessary to run the geology extension.
'
'********************************************************************************

Option Explicit

Public g_strGeoDB_Path As String

'These global variables are used to connect to the Geology SDE database.
'Added 20051212, Jim McDonald
Public g_strGeoSDE_Server As String
Public g_strGeoSDE_User As String
Public g_strGeoSDE_Instance As String
Public g_strGeoSDE_Password As String
Public g_strGeoSDE_Version As String
Public g_pGeoSDEprops As IPropertySet

Public g_strProjectsDB_Path As String
Public g_strBedrockDB_Path As String
'Public g_strGeoDir_Path As String
Public g_strDRG_Path As String
Public g_strSCAN_Path As String
Public g_strExport_Path As String
'Public g_strSelect_Quad As String
Public g_strDRGLOC_Path As String ' added because they want to keep the layer files in a separate dir from the DRG's
'Public g_strCounty_Names As String 'holds the county names within the 24K BG select_quad
Public g_strBase_Path As String

'Public g_blnQuadsSelected As Boolean
'Public g_blnInDataView As Boolean
'Public g_blnDocChanged As Boolean
Public g_blnMapsChanging As Boolean
Public g_blnExportDialogOpen As Boolean
Public gODNRProjectDb As ODNRProjectsDatabase
Public gODNRProject As ODNRProject

Private m_pApp As esriFramework.IApplication
Private pProjectImages As ImageList
Private m_pStepProg As IStepProgressor
Private m_pStatBar As IStatusBar

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_HWNDPARENT = (-8)

Const c_strModuleName As String = "ODNR_Common"

Public Enum ODNRQuadScale
    odnrScaleUnknown = 0
    odnr24K = 1
    odnr100K = 2
    odnr62K = 3
    odnr250K = 4
    odnr500K = 5
End Enum

Public Enum ODNRProjectType
    odnrProjectTypeUnknown = 0
    odnrTopography = 1
    odnrGeology = 2
    odnrBedrockStructure = 3
    odnrAUM = 4
End Enum

Public Enum ODNRMapType
    odnrMapTypeUnknown = 0
    odnrLocation24KMap = 1
    odnrLocation100KMap = 2
    odnrGeologyMap = 3
End Enum

Public Enum ODNRQuadExportMethod
    odnrExportMethodAll = 0
    odnrExportMethodField = 1
    odnrExportMethodSpatial = 2
End Enum

Public Enum ODNRVisibilityLevel
    odnrZoomLevelUnknown = -1
    odnrZoomLevelAll = 0
    odnrZoomLevelOhio = 1
    odnrZoomLevelQuad = 2
    odnrZoomLevelNone = 3
End Enum

Public Sub HookApplication(pApp As IApplication)
'Hook Application. Called by Select_DataDir_Cmd.ICommand_OnCreate
'since this is the first tool created by the GEO1_Toolbar
    On Error GoTo ErrorHandler
    
    Set m_pApp = pApp

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".HookApplication " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ShowMessage(strMsg As String, blnAnimation As Boolean, Optional lngInc As Long, Optional lngMax As Long)
'Show a message and progress on the ArcMap statusbar
    On Error GoTo ErrorHandler
    
    If (m_pStatBar Is Nothing) Then Set m_pStatBar = m_pApp.StatusBar
    m_pStatBar.PlayProgressAnimation blnAnimation
    If (lngInc And lngMax) Then
        If (m_pStepProg Is Nothing) Then Set m_pStepProg = m_pStatBar.ProgressBar
        m_pStepProg.Position = lngInc
        If (m_pStatBar.Visible) Then
            m_pStatBar.ShowProgressBar strMsg, 0, lngMax, lngInc, False
        Else
            m_pStatBar.ProgressBar.Show
        End If
        If (lngInc >= lngMax) Or (lngInc = -1) Then
            m_pStatBar.HideProgressBar
        End If
    Else
        m_pStatBar.HideProgressBar
        m_pStatBar.Message(0) = strMsg
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowMessage " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function LoadProjectsDatabase() As Boolean
'This routine is responsible for loading the projects database
'that contains BLOBs of each project mxd file.
'This routine is called by the Directory setup form when the OK button
'is clicked or by the "On_Create" method of the Select_DataDir_Cmd.
'The ODNR extension will not work if the projects database is not
'loaded.
    On Error GoTo ErrorHandler
    
    Dim strPath As String, pResp As Integer
    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    
    g_strDRG_Path = GetSetting("ArcView", "ODNR_Geology", "DRGDirectory") & "\"
    'g_strGeoDir_Path = GetSetting("ArcView", "ODNR_Geology", "ProjectDirectory") & "\"
    g_strSCAN_Path = GetSetting("ArcView", "ODNR_Geology", "ScansDirectory") & "\"
    g_strExport_Path = GetSetting("ArcView", "ODNR_Geology", "ExportDirectory") & "\"
    g_strBase_Path = GetSetting("ArcView", "ODNR_Geology", "Base100Directory") & "\"
    g_strDRGLOC_Path = GetSetting("ArcView", "ODNR_Geology", "DRGLOCDirectory") & "\"
    g_strGeoDB_Path = GetSetting("ArcView", "ODNR_Geology", "GeologyDatabasePath")
    g_strBedrockDB_Path = GetSetting("ArcView", "ODNR_Geology", "BedrockDatabasePath")
    g_strProjectsDB_Path = GetSetting("ArcView", "ODNR_Geology", "ProjectsDatabasePath")
        
    pResp = 0
    strPath = g_strDRG_Path & "Bethesda.lyr"
    If (pFSO.FileExists(strPath) = False) Then
        pResp = MsgBox("Error locating path to DRG files. Would you like to look for the directory now?", vbYesNo)
    End If

    If (pResp = 0) And (pFSO.FileExists(g_strGeoDB_Path) = False) Then
        pResp = MsgBox("Error locating path to Geology Geodatabase.  Would you like to look for the directory now?", vbYesNo)
    End If
    
    strPath = g_strSCAN_Path & "alliance clarion.tif"
    If (pResp = 0) And (pFSO.FileExists(strPath) = False) Then
        pResp = MsgBox("Error locating path to Structure Scan files.  Would you like to look for the directory now?", vbYesNo)
    End If

    strPath = g_strDRGLOC_Path & "Sinking_spring.tif"
    If (pResp = 0) And (pFSO.FileExists(strPath) = False) Then
        pResp = MsgBox("Error locating path to the DRG TIFs.  Would you like to look for the directory now?", vbYesNo)
    End If
    
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pWs As IWorkspace
    Set pWsFact = New AccessWorkspaceFactory
    If (pWsFact.IsWorkspace(g_strProjectsDB_Path)) Then
        Set pWs = pWsFact.OpenFromFile(g_strProjectsDB_Path, 0)
        Set gODNRProjectDb = New ODNRProjectsDatabase
        If (gODNRProjectDb.LoadDatabase(m_pApp, pWs)) Then
            LoadProjectsDatabase = True
            ODNR_Common.LoadProjectCombo
        Else
            LoadProjectsDatabase = False
        End If
    End If
            
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadProjectsDatabase " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ButtonImage(strKey As String) As IPictureDisp
'Gets as button image from the frmImages form
    On Error GoTo ErrorHandler
        
        If (pProjectImages Is Nothing) Then Load frmImages
        Set pProjectImages = frmImages.ProjectImageList
        Set ButtonImage = pProjectImages.ListImages.Item(strKey).Picture

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ButtonImage " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub LoadProjectCombo()
'Populates the project combo with the list of projects from the projects database
    On Error GoTo ErrorHandler
    
    Dim pCodeList As Collection, vKey As Variant
    Set pCodeList = gODNRProjectDb.ProjectCodeList
    frmToolbarControls.cboProject.Clear
    frmToolbarControls.cboProject.Text = "Open Project"
    For Each vKey In pCodeList
        frmToolbarControls.cboProject.AddItem CStr(vKey)
    Next
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadCombo " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UpdateProjectCombo()
'Updates the project combo when the document changes or a new project is loaded
    On Error GoTo ErrorHandler
    
    frmToolbarControls.cboProject.Text = "Open Project"
    If Not (gODNRProjectDb Is Nothing) Then
        If Not (gODNRProject Is Nothing) Then
            frmToolbarControls.cboProject.Text = gODNRProjectDb.ActiveProjectCode
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UpdateProjectCombo " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadQuadCombo()
'Populates the QuadCombo with the list of project quads
    On Error GoTo ErrorHandler
    
    Dim pQuadList As Collection, vKey As Variant
    Set pQuadList = gODNRProject.Quads.QuadNameList
    frmToolbarControls.cboQuad.Clear
    frmToolbarControls.cboQuad.Text = "Select a Quadrangle"
    For Each vKey In pQuadList
        frmToolbarControls.cboQuad.AddItem CStr(vKey)
    Next
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadCombo " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function ControlEnabled(strName As String) As Boolean
'Assignes the current enabled state of a control
    On Error GoTo ErrorHandler
    
    ControlEnabled = False
    Dim blnEnabled As Boolean, blnDbLoaded As Boolean, blnProjectLoaded As Boolean
    blnDbLoaded = False
    blnProjectLoaded = False
    If Not (gODNRProjectDb Is Nothing) Then
        blnDbLoaded = True
        If Not (gODNRProject Is Nothing) Then
            blnProjectLoaded = True
            If (g_blnExportDialogOpen = False) Then blnEnabled = True
        End If
    End If
    Select Case strName
        Case "Bedrock_Layers_Cmd"
            If (blnProjectLoaded) Then
                If (gODNRProject.InDataView) Then
                    If (gODNRProject.QuadScale = odnr24K) Then
                        If ((gODNRProject.ProjectType = odnrBedrockStructure) Or (gODNRProject.ProjectType = odnrGeology)) And (gODNRProject.IsZoomedToQuadSelection) Then
                            ControlEnabled = blnEnabled
                        End If
                    End If
                End If
            End If
        Case "Export_Image_Cmd"
            ControlEnabled = blnEnabled
        Case "Export_Tool"
            If (blnProjectLoaded) Then
                ControlEnabled = True
            End If
        Case "GoDataView_Cmd"
            If (blnProjectLoaded) Then
                If (gODNRProject.InDataView = False) Then
                    ControlEnabled = blnEnabled
                End If
            End If
        Case "GoLayout_Cmd"
            If (blnProjectLoaded) Then
                If (gODNRProject.InDataView) And (gODNRProject.IsZoomedToQuadSelection) Then
                    ControlEnabled = blnEnabled
                End If
            End If
        Case "PickQuad_Tool"
            If (blnProjectLoaded) Then
                If (gODNRProject.IsZoomedToQuadSelection = False) Then
                    ControlEnabled = blnEnabled
                End If
            End If
        Case "Quad_Combo"
            If (blnProjectLoaded) Then
                If (gODNRProject.IsZoomedToQuadSelection = False) Then
                    ControlEnabled = blnEnabled
                End If
            End If
        Case "Select_DataDir_Cmd"
            If (g_blnExportDialogOpen = False) Then ControlEnabled = True
        Case "Switch_combo"
            If (blnDbLoaded) Then
                If (g_blnExportDialogOpen = False) Then ControlEnabled = True
            End If
        Case "Select100_Tool"
            If (blnProjectLoaded) Then
                If (gODNRProject.QuadScale = odnr100K) And (gODNRProject.IsZoomedToQuadSelection = False) Then
                    ControlEnabled = blnEnabled
                End If
            End If
        Case "ZoomToOhio_Cmd"
            If (blnProjectLoaded) Then
                If (gODNRProject.IsZoomedToQuadSelection) Then
                    ControlEnabled = blnEnabled
                End If
            End If
    End Select
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ControlEnabled " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub VerifyProjectLayers()
'The purpose of this procedure is to verify that the correct layers exist
'in each map in the map document.  It references the ODNRObjects database
'to reference each appropriate layer in the map
    On Error GoTo ErrorHandler
    
    Dim pStateLayer As ODNRStateLayer, pMap As IMap, pLyr As ILayer
    Dim pMapList As Collection, lngIdx As Long
    Dim pMxDoc As IMxDocument, pContentsView As IContentsView
    Set pMapList = New Collection
    pMapList.Add Item:=gODNRProject.ProjectMap(odnrGeologyMap)
    pMapList.Add Item:=gODNRProject.ProjectMap(odnrLocation24KMap)
    pMapList.Add Item:=gODNRProject.ProjectMap(odnrLocation100KMap)
    
    Set pMxDoc = m_pApp.Document
    Set pContentsView = pMxDoc.CurrentContentsView
    
    For lngIdx = 1 To pMapList.Count
        Set pMap = pMapList.Item(lngIdx)
        If Not (pMap Is Nothing) Then
            gODNRProject.StateLayers.ActiveMapName = pMap.Name
            gODNRProject.StateLayers.Reset
            Set pStateLayer = gODNRProject.StateLayers.NextLayer
            Do While Not pStateLayer Is Nothing
                If (pStateLayer.InMap = False) Then
                    Set pLyr = pStateLayer.ESRILayer
                    If Not (pLyr Is Nothing) Then
                        pLyr.Visible = False
                        pMap.AddLayer pLyr
                    End If
                End If
                Set pStateLayer = gODNRProject.StateLayers.NextLayer
            Loop
            pContentsView.Refresh pMap
        End If
    Next
    
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".VerifyProjectLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function LoadFromGxFile(strFilePath As String, Optional strRepairPath As String) As ILayer
'Load a Raster or feature layer (.lyr file) from a GxFile
    On Error GoTo ErrorHandler
    
    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    If (pFSO.FileExists(strFilePath)) Then
        strFilePath = pFSO.GetFile(strFilePath).Path
        Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
        Dim pLayer As ILayer, pRLayer As IRasterLayer, pFLayer As IFeatureLayer
        Dim pDataLayer As IDataLayer, pDsName As IDatasetName
        Dim lngB1 As Long, lngB2 As Long, strDsPath As String, strCheckDsPath As String
        
        Set pGxLayer = New esriCatalog.GxLayer
        Set pGxFile = pGxLayer
        pGxFile.Path = strFilePath
        pGxFile.Open
        Set pLayer = pGxLayer.Layer
        If (strRepairPath <> "") Then
            If (TypeOf pLayer Is IRasterLayer) Then
                Set pRLayer = pLayer
                If (Strings.Left(pRLayer.FilePath, 6) = "RASTER") Then
                    lngB1 = Strings.InStr(20, pRLayer.FilePath, ";")
                    lngB2 = Strings.InStr(lngB1 + 1, pRLayer.FilePath, ";")
                    strDsPath = Strings.Mid(pRLayer.FilePath, 21, lngB1 - 21)
                    strDsPath = strDsPath & Strings.Mid(pRLayer.FilePath, lngB1 + 18, lngB2 - (lngB1 + 18))
                Else
                    strDsPath = pRLayer.FilePath
                End If
                strCheckDsPath = strRepairPath & pFSO.GetFileName(strDsPath)
                If (strDsPath <> strCheckDsPath) Then
                    Set pGxNewLayer = RepairRasterLayerDataSource(strFilePath, strCheckDsPath)
                    If Not (pGxNewLayer Is Nothing) Then
                        Set pLayer = pGxNewLayer.Layer
                    End If
                End If
            ElseIf (TypeOf pLayer Is IFeatureLayer) Then
                Set pDataLayer = pLayer
                Set pDsName = pDataLayer.DataSourceName
                If (g_strGeoDB_Path <> pDsName.WorkspaceName.PathName) Then
                    Set pGxNewLayer = RepairFeatureLayerDataSource(strFilePath, g_strGeoDB_Path, pDsName.Name)
                    If Not (pGxNewLayer Is Nothing) Then
                        Set pLayer = pGxNewLayer.Layer
                    End If
                End If
            End If
        End If
        Set LoadFromGxFile = pLayer
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadFromGxFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function RepairRasterLayerDataSource(strLayerPath As String, strNewDsPath As String) As IGxLayer
'Repair a raster layer data source
    On Error GoTo ErrorHandler

    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    If (pFSO.FileExists(strLayerPath)) And (pFSO.FileExists(strNewDsPath)) Then
        Dim pLayer As ILayer, pGxLayer As IGxLayer, pGxFile As IGxFile
        Set pGxLayer = New esriCatalog.GxLayer
        Set pGxFile = pGxLayer
        pGxFile.Path = strLayerPath
        pGxFile.Open
        Set pLayer = pGxLayer.Layer
        pGxFile.Close False
        If (TypeOf pLayer Is IRasterLayer) Then
            Dim pRLayer As IRasterLayer, pNewRLayer As IRasterLayer
            Set pRLayer = pLayer
            Set pNewRLayer = New RasterLayer
            pNewRLayer.CreateFromFilePath strNewDsPath
            If Not (pRLayer.Renderer Is Nothing) Then
                Set pNewRLayer.Renderer = pRLayer.Renderer
                pNewRLayer.Renderer.Update
            End If
            Set pGxLayer = New GxLayer
            Set pGxFile = pGxLayer
            pGxFile.Path = strLayerPath
            Set pGxLayer.Layer = pNewRLayer
            pGxFile.Save
            Set RepairRasterLayerDataSource = pGxLayer
        End If
    Else
        Set RepairRasterLayerDataSource = Nothing
    End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".RepairRasterLayerDataSource " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function RepairFeatureLayerDataSource(strLayerPath As String, strWsPath As String, strFcName As String) As IGxLayer
'Repair a featurelayer data source
    On Error GoTo ErrorHandler
    
    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    If (pFSO.FileExists(strLayerPath)) And (pFSO.FileExists(strWsPath)) Then
        Dim pLayer As ILayer, pGxLayer As IGxLayer, pGxFile As IGxFile
        Set pGxLayer = New esriCatalog.GxLayer
        Set pGxFile = pGxLayer
        pGxFile.Path = strLayerPath
        pGxFile.Open
        Set pLayer = pGxLayer.Layer
        pGxFile.Close False
        If (TypeOf pLayer Is IFeatureLayer) Then
            Dim pFLayer As IFeatureLayer, pFc As IFeatureClass
            Dim pWsFact As IWorkspaceFactory, pFWs As IFeatureWorkspace
            Set pFLayer = pLayer
            Set pWsFact = New AccessWorkspaceFactory
            Set pFWs = pWsFact.OpenFromFile(strWsPath, 0)
            Set pFc = pFWs.OpenFeatureClass(strFcName)
            Set pFLayer.FeatureClass = pFc
            Set pGxLayer = New GxLayer
            Set pGxFile = pGxLayer
            pGxFile.Path = strLayerPath
            Set pGxLayer.Layer = pFLayer
            pGxFile.Save
            Set RepairFeatureLayerDataSource = pGxLayer
        End If
    Else
        Set RepairFeatureLayerDataSource = Nothing
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".RepairFeatureLayerDataSource " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub CollapseLegend(pLayer As ILayer)
'Collapse the legend of a layer
    On Error GoTo ErrorHandler

    If Not (pLayer Is Nothing) Then
        Dim pFlyr As IFeatureLayer, pRLyr As IRasterLayer
        Dim pLegendInfo As ILegendInfo, lngLegendIdx As Long
        If (TypeOf pLayer Is IFeatureLayer) Then
            Set pFlyr = pLayer
            Set pLegendInfo = pFlyr
        ElseIf (TypeOf pLayer Is IRasterLayer) Then
            Set pRLyr = pLayer
            Set pLegendInfo = pRLyr
        End If
        For lngLegendIdx = 0 To pLegendInfo.LegendGroupCount - 1
            pLegendInfo.LegendGroup(lngLegendIdx).Visible = False
        Next
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CollapseLegend " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function ODNRGeologyWorkspace() As IWorkspace
'Return the workspace object representing the geology database
    On Error GoTo ErrorHandler
    
    Dim pWs As IWorkspace, pWsFact As IWorkspaceFactory
    Set pWsFact = New AccessWorkspaceFactory
    
    'This is the temporary hack to test the SDE connection
    '20051215, Jim McDonald
    
'    Dim pWsSDE As IWorkspace
'    Dim pWsFactSDE As IWorkspaceFactory
'    Dim pConnProp As IPropertySet
    
    'Added SDE connection properties, 20051212, Jim McDonald
'    Set pConnProp = New PropertySet
'    With pConnProp
'        .SetProperty "SERVER", "NRSDE"
'        .SetProperty "USER", "mcdonald"
'        .SetProperty "DATABASE", "GEOLOGY"
'        .SetProperty "INSTANCE", "5151"
'        .SetProperty "PASSWORD", "jimsde83"
'        .SetProperty "VERSION", "SDE.DEFAULT"
'    End With
'    Set pWsFactSDE = New SdeWorkspaceFactory
'    Set pWsSDE = pWsFactSDE.Open(pConnProp, 0)
    
    Set pWs = pWsFact.OpenFromFile(g_strGeoDB_Path, 0)
    Set ODNRGeologyWorkspace = pWs

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ODNRGeologyWorkspace " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ObjectClassExists(strClassName As String, pWs As IWorkspace) As Boolean
'Verify the existence of an object class in a geodatabase
    On Error GoTo ErrorHandler
    
    If (Not pWs Is Nothing) And (strClassName <> "") Then
        Dim pFWsManage As IFeatureWorkspaceManage
        Set pFWsManage = pWs
        If (pFWsManage.IsRegisteredAsObjectClass(strClassName)) Then
            ObjectClassExists = True
        Else
            ObjectClassExists = False
        End If
    End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ObjectClassExists " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub SelectPointerTool()
'Select ArcMap pointer tool (in order to unselect some other tool).
    On Error GoTo ErrorHandler

    Dim pUID As New UID
    Dim pCmdItem As ICommandItem
    pUID.Value = "{C22579D1-BC17-11D0-8667-0000F8751720}" 'Select_Elements tool
    Set pCmdItem = m_pApp.Document.CommandBars.Find(pUID)
    pCmdItem.Execute
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CloseDialogs " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'****************************************************
