VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pLayer As ILayer
Private m_pVisibilityLevel As ODNRVisibilityLevel
Private m_blnIsLoaded As Boolean

Const c_strModuleName As String = "ODNRLayer"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pLayer = Nothing
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get Name() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then Name = m_pLayer.Name
        
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Set ESRILayer = m_pLayer
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayerType() As esriDatasetType
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pDs As IDataset, pFlyr As IFeatureLayer, pRlyr As IRasterLayer
        If (TypeOf m_pLayer Is IFeatureLayer) Then
            Set pFlyr = m_pLayer
            Set pDs = pFlyr
        ElseIf (TypeOf m_pLayer Is IRasterLayer) Then
            Set pRlyr = m_pLayer
            Set pDs = pRlyr
        End If
        If Not (pDs Is Nothing) Then
            ESRILayerType = pDs.Type
        Else
            ESRILayerType = esriDTAny
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayerType " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get VisibilityLevel() As ODNRVisibilityLevel
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then VisibilityLevel = m_pVisibilityLevel
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".VisibilityLevel " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub Export(strExportDir As String, strName As String)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If (TypeOf m_pLayer Is IFeatureLayer) Then
            Dim pFlyr As IFeatureLayer, pFlyrDef As IFeatureLayerDefinition
            Dim pFc As IFeatureClass, pDs As IDataset, pDsName As IDatasetName
            Dim strQryDef As String, pQf As IQueryFilter, pSelSet As ISelectionSet
            Dim pInFcName As IFeatureClassName
            
            Set pFlyr = m_pLayer
            Set pFlyrDef = pFlyr
            Set pFc = pFlyr.FeatureClass
            If (pFlyrDef.DefinitionExpression <> "") Then
                strQryDef = pFlyrDef.DefinitionExpression
                Set pQf = New QueryFilter
                pQf.WhereClause = strQryDef
                Set pSelSet = pFc.Select(pQf, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
            Else
                Set pSelSet = Nothing
            End If
            
            Set pDs = pFeatureClass
            Set pInFcName = pDs.FullName
            
            Dim pInWs As IWorkspace
            Set pInWs = pDs.Workspace
            
            Dim pOutFc As IFeatureClass, pOutFcName As IFeatureClassName
            Dim pOutDs As IDataset, pOutWs As IWorkspace, pWsFact As IWorkspaceFactory
            Dim pOutWsName As IWorkspaceName
            Set pWsFact = New ShapefileWorkspaceFactory
            
            Set pOutWs = pWsFact.OpenFromFile(strExportDir, 0)
            Set pDs = pOutWs
            Set pOutWsName = pDs.FullName
            
            Set pOutFcName = New FeatureClassName
            Set pDsName = pOutFcName
            pDsName.Name = strName
            Set pDsName.WorkspaceName = pOutWsName
            
            Dim pInFields As IFields, pOutFields As IFields
            Dim pFieldChecker As IFieldChecker
            Set pInFields = pFc.Fields
            Set pFieldChecker = New FieldChecker
            pFieldChecker.InputWorkspace = pInWs
            Set pFieldChecker.ValidateWorkspace = pOutWs
            pFieldChecker.Validate pInFields, Nothing, pOutFields
            
            Dim pFDataCon As IFeatureDataConverter2
            Set pFDataCon = New FeatureDataConverter
            pFDataCon.ConvertFeatureClass pInFcName, Nothing, pSelSet, Nothing, pOutFcName, Nothing, pOutFields, "", 100, 0
        ElseIf (TypeOf m_pLayer Is IRasterLayer) Then
            
        
        End If
    Exit Sub
    
ErrorHandler:
    HandleError True, c_strModuleName & ".Export " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LimitToQuads(pQuadList As Collection, strQuadField As String)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If (TypeOf m_pLayer Is IFeatureLayer) Then
            Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
            Dim pFlyrDef As IFeatureLayerDefinition
            Dim strQryDef As String, lngIdx As Long
            Set pFlyr = m_pLayer
            Set pFc = pFlyr.FeatureClass
            If (pFc.FindField(strQuadField) > -1) And (pQuadList.Count > 0) Then
                Set pFlyrDef = pFlyr
                For lngIdx = 1 To pQuadList.Count
                    If (strQryDef = "") Then
                        strQryDef = strQuadField & " in (" & pQuadList.Item(lngIdx)
                    Else
                        strQryDef = strQryDef & "," & pQuadList.Item(lngIdx)
                    End If
                Next
                stryQryDef = strQryDef & ")"
                pFlyrDef.DefinitionExpression = strQryDef
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LimitToQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ShowAllFeatures()
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If (TypeOf m_pLayer Is IFeatureLayer) Then
            Dim pFlyr As IFeatureLayer, pFlyrDef As IFeatureLayerDefinition
            Set pFlyr = m_pLayer
            Set pFlyrDef = pFlyr
            pFlyrDef.DefinitionExpression = ""
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowAllFeatures " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadFromLayer(pLayer As ILayer, pVisibilityLevel As ODNRVisibilityLevel)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded = False) Then
        If Not (pLayer Is Nothing) Then
            m_pVisibilityLevel = pVisibilityLevel
            Set m_pLayer = pLayer
            If Not (m_pLayer Is Nothing) Then
                m_blnIsLoaded = True
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadFromGxFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadFromGxFile(strFilePath As String, pVisibilityLevel As ODNRVisibilityLevel, Optional strRepairPath As String)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded = False) Then
        Dim pFSO As FileSystemObject
        Set pFSO = New FileSystemObject
        If (pFSO.FileExists(strFilePath)) Then
            Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
            Dim pLayer As ILayer, pRLayer As IRasterLayer, pFLayer As IFeatureLayer
            Dim pDataLayer As IDataLayer, pDsName As IDatasetName
            Dim lngB1 As Long, lngB2 As Long, strDsPath As String, strCheckDsPath As String
            
            m_pVisibilityLevel = pVisibilityLevel
            Set pGxLayer = New esriCore.GxLayer
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
            Set m_pLayer = pLayer
            If Not (m_pLayer Is Nothing) Then
                m_blnIsLoaded = True
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadFromGxFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function RepairRasterLayerDataSource(strLayerPath As String, strNewDsPath As String) As IGxLayer
    On Error GoTo ErrorHandler

    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    If (pFSO.FileExists(strLayerPath)) And (pFSO.FileExists(strNewDsPath)) Then
        Dim pLayer As ILayer, pGxLayer As IGxLayer, pGxFile As IGxFile
        Set pGxLayer = New esriCore.GxLayer
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
    On Error GoTo ErrorHandler
    
    Dim pFSO As FileSystemObject
    Set pFSO = New FileSystemObject
    If (pFSO.FileExists(strLayerPath)) And (pFSO.FileExists(strWsPath)) Then
        Dim pLayer As ILayer, pGxLayer As IGxLayer, pGxFile As IGxFile
        Set pGxLayer = New esriCore.GxLayer
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

