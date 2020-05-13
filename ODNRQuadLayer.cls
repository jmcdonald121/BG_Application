VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRQuadLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRQuadLayer
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object representing the map layers in the
'               project that are specific to a selected set of quads.
'               These layers only cover the extent of the quad, not the
'               state.  This object is owned by ODNRQuadLayers.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pParent As ODNRQuadLayers
Private m_pLayer As ILayer
Private m_pDsType As esriDatasetType
Private m_pVisibilityLevel As ODNRVisibilityLevel
Private m_lngMapType As ODNRMapType
Private m_blnCanExport As Boolean
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRQuadLayer"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pLayer = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get Name() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Name = m_pLayer.Name
        
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set ESRILayer = m_pLayer
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayerType() As esriDatasetType
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then ESRILayerType = m_pDsType
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayerType " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get CanExport() As Boolean
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then CanExport = m_blnCanExport
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRIFeatureType " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get VisibilityLevel() As ODNRVisibilityLevel
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then VisibilityLevel = m_pVisibilityLevel
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".VisibilityLevel " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get InMap() As Boolean
    On Error GoTo ErrorHandler
    
    InMap = False
    If (m_blnIsloaded) Then
        Dim pMap As IMap, pLayer As ILayer, lngIdx As Long
        Set pMap = m_pParent.Parent.ProjectMap(m_lngMapType)
        For lngIdx = 0 To pMap.LayerCount - 1
            Set pLayer = pMap.Layer(lngIdx)
            If (pLayer Is m_pLayer) Then
                InMap = True
                Exit For
            End If
        Next
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".VisibilityLevel " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function Export(strExportDir As String) As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (m_blnCanExport) Then
            Dim pLyr As ILayer, pGlyr As IGroupLayer, pClyr As ICompositeLayer, lngIdx As Long
            Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pOutFc As IFeatureClass
            Dim pFlyrDef As IFeatureLayerDefinition, pQf As IQueryFilter
            Dim pDs As IDataset, pDsName As IDatasetName, pOutDs As IDataset
            Dim pInFcName As IFeatureClassName, pOutFcName As IFeatureClassName
            Dim pInWs As IWorkspace, pOutWs As IWorkspace, pWsFact As IWorkspaceFactory
            Dim pSelSet As ISelectionSet, pOutWsName As IWorkspaceName
            Dim pEnumInvalidObj As IEnumInvalidObject, pInvalidObjInfo As IInvalidObjectInfo
            Dim pInFields As IFields, pOutFields As IFields, blnDoExport As Boolean
            Dim pFieldChecker As IFieldChecker, pFDataCon As IFeatureDataConverter2
            
            Dim pRLyr As IRasterLayer, strFilePath As String, strDest As String
            Dim pFSO As FileSystemObject, strBaseFileName As String
            Dim strMsg As String, strExt As String
            
            strMsg = "Initializing export of grouplayer: " & Me.Name & "..."
            If (m_pLayer.Valid = False) Then
                strMsg = strMsg & vbCrLf & vbTab & "layer: " & m_pLayer.Name & " is invalid."
            ElseIf (TypeOf m_pLayer Is IGroupLayer) Then
                Set pGlyr = m_pLayer
                Set pClyr = pGlyr
                If (pClyr.Count > 0) Then
                    For lngIdx = 0 To pClyr.Count - 1
                        blnDoExport = True
                        Set pLyr = pClyr.Layer(lngIdx)
                        If (pLyr.Valid) Then
                            ODNR_Common.ShowMessage "Exporting " & m_pLayer.Name & " (" & pLyr.Name & ")...", False
                            If (TypeOf pLyr Is IFeatureLayer) Then
                                strMsg = strMsg & vbCrLf & vbTab & "exported featurelayer: " & pLyr.Name
                                Set pFlyr = pLyr
                                Set pFlyrDef = pFlyr
                                Set pFc = pFlyr.FeatureClass
                                If (pFlyrDef.DefinitionExpression <> "") Then
                                    Set pQf = New QueryFilter
                                    pQf.WhereClause = pFlyrDef.DefinitionExpression
                                    Set pSelSet = pFc.Select(pQf, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
                                    If (pSelSet.Count = 0) Then
                                        blnDoExport = False
                                    End If
                                Else
                                    Set pSelSet = Nothing
                                End If
                                If (blnDoExport) Then
                                    Set pDs = pFc
                                    Set pInFcName = pDs.FullName
                                    
                                    Set pInWs = pDs.Workspace
                                    Set pWsFact = New ShapefileWorkspaceFactory
                                    Set pOutWs = pWsFact.OpenFromFile(strExportDir, 0)
                                    Set pDs = pOutWs
                                    Set pOutWsName = pDs.FullName
                                    
                                    Set pOutFcName = New FeatureClassName
                                    Set pDsName = pOutFcName
                                    pDsName.Name = pLyr.Name
                                    Set pDsName.WorkspaceName = pOutWsName
                                    
                                    Set pInFields = pFc.Fields
                                    Set pFieldChecker = New FieldChecker
                                    pFieldChecker.InputWorkspace = pInWs
                                    Set pFieldChecker.ValidateWorkspace = pOutWs
                                    pFieldChecker.Validate pInFields, Nothing, pOutFields
                                    
                                    Set pFDataCon = New FeatureDataConverter
                                    Set pEnumInvalidObj = pFDataCon.ConvertFeatureClass(pInFcName, Nothing, pSelSet, Nothing, pOutFcName, Nothing, pOutFields, "", 100, 0)
                                    Set pInvalidObjInfo = pEnumInvalidObj.Next
                                    If (pInvalidObjInfo Is Nothing) Then
                                        strMsg = strMsg & " succesfully."
                                    Else
                                        strMsg = strMsg & " unsuccesfully."
                                    End If
                                    Set pInWs = Nothing
                                    Set pOutWs = Nothing
                                    Set pFieldChecker = Nothing
                                    Set pFDataCon = Nothing
                                    Set pEnumInvalidObj = Nothing
                                Else
                                    strMsg = strMsg & " unsuccesfully."
                                End If
                            ElseIf (TypeOf pLyr Is IRasterLayer) Then
                                strMsg = strMsg & vbCrLf & vbTab & "exported rasterlayer: " & pLyr.Name
                                Set pRLyr = pLyr
                                Set pFSO = New FileSystemObject
                                strFilePath = pRLyr.FilePath
                                strBaseFileName = pFSO.GetBaseName(strFilePath)
                                strExt = Right(pFSO.GetFileName(strFilePath), 3)
                                strDest = strExportDir & "\"
                                strFilePath = Left(strFilePath, Len(strFilePath) - 3) & "*"
                                pFSO.CopyFile strFilePath, strDest, True
                                If (pFSO.FileExists(strDest & strBaseFileName & "." & strExt)) Then
                                    strMsg = strMsg & " succesfully."
                                Else
                                    strMsg = strMsg & " unsuccesfully."
                                End If
                                Set pFSO = Nothing
                            End If
                        Else
                            strMsg = strMsg & vbCrLf & vbTab & "layer: " & pLyr.Name & " is invalid."
                        End If
                        ODNR_Common.ShowMessage "", False
                    Next
                    Export = strMsg
                Else
                    strMsg = strMsg & vbCrLf & vbTab & "grouplayer is empty."
                End If
            End If
        Else
            Export = "Unable to export " & Me.Name & "."
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError True, c_strModuleName & ".Export " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub LoadFromLayer(pApp As esriFramework.IApplication, pParent As ODNRQuadLayers, pLayer As ILayer, pDsType As esriDatasetType, pVisibilityLevel As ODNRVisibilityLevel, lngMapType As ODNRMapType, blnCanExport As Boolean)
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded = False) Then
        If (Not pApp Is Nothing) And (Not pParent Is Nothing) And (Not pLayer Is Nothing) Then
            Set m_pApp = pApp
            Set m_pParent = pParent
            m_pDsType = pDsType
            m_pVisibilityLevel = pVisibilityLevel
            m_lngMapType = lngMapType
            m_blnCanExport = blnCanExport
            Set m_pLayer = pLayer
            If Not (m_pLayer Is Nothing) Then
                m_blnIsloaded = True
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadFromLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

