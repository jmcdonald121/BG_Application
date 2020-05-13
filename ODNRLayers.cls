VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pApp As esriCore.IApplication
Private m_pParent As ODNRProject
Private m_pProjectLayers As Dictionary
Private m_pQuadLayerTypes As Dictionary
Private m_pQuadLayers As Dictionary
Private m_pFilePaths As Dictionary
Private m_strActiveMap As String
Private m_strActiveQuad As String
Private m_lngLayerIndex As Long
Private m_lngQuadLayerIndex As Long
Private m_lngQuadLayerTypeIndex As Long
Private m_blnIsLoaded As Boolean

Const c_strModuleName As String = "ODNRLayers"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsLoaded = False
    m_lngLayerIndex = 0
    m_lngQuadLayerIndex = 0
    m_lngQuadLayerTypeIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pProjectLayers = Nothing
    Set m_pQuadLayers = Nothing
    Set m_pFilePaths = Nothing
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get ActiveMapName() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then ActiveMapName = m_strActiveMap

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMapName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let ActiveMapName(RHS As String)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If (RHS <> m_strActiveMap) Then
            Dim pMapLayers As Dictionary
            Set pMapLayers = m_pProjectLayers.Item(m_pParent.ProjectName)
            If (pMapLayers.Exists(RHS)) Then
                m_strActiveMap = RHS
            Else
                m_strActiveMap = m_pParent.GeologyMapName
            End If
            m_lngLayerIndex = 0
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMapName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get LayerCount(strMapName As String) As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pMapLayers As Dictionary, pLayerList As Dictionary
        Set pMapLayers = m_pProjectLayers.Item(m_pParent.ProjectName)
        Set pLayerList = pMapLayers.Item(m_strActiveMap)
        LayerCount = pLayerList.Count
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function NextLayer() As ODNRLayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pMapLayers As Dictionary, pLayerList As Dictionary
        Dim pODNRLayer As ODNRLayer, pLayer As ILayer, strName As String
        Set pMapLayers = m_pProjectLayers.Item(m_pParent.ProjectName)
        Set pLayerList = pMapLayers.Item(m_strActiveMap)
        If (m_lngLayerIndex < pLayerList.Count) Then
            strName = pLayerList.Keys(m_lngLayerIndex)
            Set pODNRLayer = New ODNRLayer
            Set pLayer = FindLayer(strName)
            If Not (pLayer Is Nothing) Then
                pODNRLayer.LoadFromLayer pLayer, pLayerList.Item(CVar(strName))
            Else
                pODNRLayer.LoadFromGxFile m_pFilePaths.Item(CVar(strName)), pLayerList.Item(CVar(strName)), g_strGeoDB_Path
            End If
            Set NextLayer = pODNRLayer
            m_lngLayerIndex = m_lngLayerIndex + 1
        Else
            Set NextLayer = Nothing
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".NextLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub Reset()
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then m_lngLayerIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Reset " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function NextQuadLayer() As ODNRLayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If Not (m_pQuadLayers Is Nothing) Then
            If (m_pQuadLayers.Count > 0) Then
                Dim pODNRLayer As ODNRLayer
                If (m_lngQuadLayerIndex < m_pQuadLayers.Count) Then
                    Set pODNRLayer = m_pQuadLayers.Item(m_pQuadLayers.Keys(m_lngQuadLayerIndex))
                    Set NextQuadLayer = pODNRLayer
                    m_lngQuadLayerIndex = m_lngQuadLayerIndex + 1
                Else
                    Set NextQuadLayer = Nothing
                End If
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".NextQuadLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ResetQuadLayers()
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then m_lngQuadLayerIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ResetQuadLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get QuadLayerTypeCount() As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pLayerList As Dictionary
        Set pLayerList = m_pQuadLayerTypes.Item(m_pParent.QuadScale)
        If Not (pLayerList Is Nothing) Then
            QuadLayerTypeCount = pLayerList.Count
        Else
            QuadLayerTypeCount = 0
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadLayerTypeCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function NextQuadLayerType() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        If Not (m_pQuadLayerTypes Is Nothing) Then
            Dim pLayerList As Dictionary
            Set pLayerList = m_pQuadLayerTypes.Item(m_pParent.QuadScale)
            If (pLayerList.Count > 0) Then
                If (m_lngQuadLayerTypeIndex < pLayerList.Count) Then
                    NextQuadLayerType = pLayerList.Keys(m_lngQuadLayerTypeIndex)
                    m_lngQuadLayerTypeIndex = m_lngQuadLayerTypeIndex + 1
                Else
                    NextQuadLayerType = ""
                End If
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".NextQuadLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ResetQuadLayerTypes()
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then m_lngQuadLayerTypeIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ResetQuadLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub Load(pApp As esriCore.IApplication, pParent As ODNRProject)
    On Error GoTo ErrorHandler
    
    If Not (pParent Is Nothing) Then
        Set m_pApp = pApp
        Set m_pParent = pParent
        LoadProjectLayers
        LoadQuadSpecificLayerTypes
        LoadQuadSpecificLayers
        LoadLayerFilePaths
        m_strActiveMap = m_pParent.GeologyMapName
        m_blnIsLoaded = True
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function FindLayer(strName As String) As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, pLayer As ILayer, lngIdx As Long
        Set pMxDoc = m_pApp.Document
        For lngIdx = 0 To pMxDoc.Maps.Count - 1
            If (m_strActiveMap = pMxDoc.Maps.Item(lngIdx).Name) Then
                Set pMap = pMxDoc.Maps.Item(lngIdx)
                Exit For
            End If
        Next
        If Not (pMap Is Nothing) Then
            For lngIdx = 0 To pMap.LayerCount - 1
                If (strName = pMap.Layer(lngIdx).Name) Then
                    Set pLayer = pMap.Layer(lngIdx)
                    Exit For
                End If
            Next
        End If
        If Not (pLayer Is Nothing) Then
            Set FindLayer = pLayer
        Else
            Set FindLayer = Nothing
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".FindLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub LoadQuadSpecificLayers()
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pLayerTypes As Dictionary, vKey As Variant, strPath As String
        Dim strName As String, pODNRLayer As ODNRLayer, strRepairPath As String
        Dim pWsFact As IWorkspaceFactory, pRWs As IRasterWorkspace
        Dim pRDs As IRasterDataset, pRLayer As IRasterLayer
        Set m_pQuadLayers = New Dictionary
        
        m_lngQuadLayerIndex = 0
        If (m_strActiveQuad <> "") Then
            Set pLayerTypes = m_pQuadLayerTypes.Item(m_pParent.QuadScale)
            For Each vKey In pLayerTypes
                Select Case CStr(vKey)
                    Case "DRGs"
                        strPath = pLayerTypes.Item(vKey) & m_strActiveQuad & ".lyr"
                        strRepairPath = g_strDRGLOC_Path & ".tif"
                        Set pODNRLayer = New ODNRLayer
                        pODNRLayer.LoadFromGxFile strPath, odnrZoomLevelQuad, strRepairPath
                        If Not (pODNRLayer.ESRILayer Is Nothing) Then
                            m_pQuadLayers.Add Key:=pODNRLayer.Name, Item:=pODNRLayer
                        End If
                    Case "Bedrock Structure Scans"
                        Set pWsFact = New RasterWorkspaceFactory
                        Set pRWs = pWsFact.OpenFromFile(g_strSCAN_Path, 0)
                        strPath = pLayerTypes.Item(vKey)
                        strPath = strPath & m_strActiveQuad & " *.tif"
                        strName = Dir$(strPath)
                        Do While strName <> ""
                            Set pRDs = pRWs.OpenRasterDataset(strName)
                            Set pRLayer = New RasterLayer
                            pRLayer.CreateFromDataset pRDs
                            Set pODNRLayer = New ODNRLayer
                            pODNRLayer.LoadFromLayer pRLayer, odnrZoomLevelQuad
                            If Not (pODNRLayer.ESRILayer Is Nothing) Then
                                m_pQuadLayers.Add Key:=pODNRLayer.Name, Item:=pODNRLayer
                            End If
                        Loop
                    Case "Base 100 Points"
                        strPath = pLayerTypes.Item(vKey) & m_strActiveQuad & "_X.lyr"
                        Set pODNRLayer = New ODNRLayer
                        pODNRLayer.LoadFromGxFile strPath, odnrZoomLevelQuad, g_strGeoDB_Path
                        If Not (pODNRLayer.ESRILayer Is Nothing) Then
                            m_pQuadLayers.Add Key:=pODNRLayer.Name, Item:=pODNRLayer
                        End If
                    Case "Base 100 Lines"
                        strPath = pLayerTypes.Item(vKey) & m_strActiveQuad & "_L.lyr"
                        Set pODNRLayer = New ODNRLayer
                        pODNRLayer.LoadFromGxFile strPath, odnrZoomLevelQuad, g_strGeoDB_Path
                        If Not (pODNRLayer.ESRILayer Is Nothing) Then
                            m_pQuadLayers.Add Key:=pODNRLayer.Name, Item:=pODNRLayer
                        End If
                    Case "Base 100 Polygons"
                        strPath = pLayerTypes.Item(vKey) & m_strActiveQuad & "_P.lyr"
                        Set pODNRLayer = New ODNRLayer
                        pODNRLayer.LoadFromGxFile strPath, odnrZoomLevelQuad, g_strGeoDB_Path
                        If Not (pODNRLayer.ESRILayer Is Nothing) Then
                            m_pQuadLayers.Add Key:=pODNRLayer.Name, Item:=pODNRLayer
                        End If
                End Select
            Next
        End If
    End If
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadSpecificLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadProjectLayers()
    On Error GoTo ErrorHandler
    
    Dim pLayerList As Dictionary, pMapList As Dictionary
    Set m_pProjectLayers = New Dictionary
    
    'BT24K_Quad.mxd
    Set pMapList = New Dictionary
    'Load Location 24K Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Quad24K Index", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Location 24K", Item:=pLayerList
    'Load Geology Map24BT Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="Boreholes", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Bedrock_Topography", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad24K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelOhio
    pLayerList.Add Key:="Topo_Annotation", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Geology Map24BT", Item:=pLayerList
    'Load BT24K_Quad.mxd Project
    m_pProjectLayers.Add Key:="BT24K_Quad.mxd", Item:=pMapList

    'BG24K_Quad.mxd
    Set pMapList = New Dictionary
    'Load Location 24K Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Quad24K Index", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Location 24K", Item:=pLayerList
    'Load Geology Map24BG Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="BG Units Ply", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Units Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad24K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="County_Bndry", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="BG Units Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Misc Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Structure Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Structure Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Ply", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Anno", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Demarcation", Item:=odnrZoomLevelNone
    pMapList.Add Key:="Geology Map24BG", Item:=pLayerList
    'Load BG24K_Quad.mxd Project
    m_pProjectLayers.Add Key:="BG24K_Quad.mxd", Item:=pMapList


    'BT100K_Quad.mxd
    Set pMapList = New Dictionary
    'Load Location 100K Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="County_Bndry", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Location 100K", Item:=pLayerList
    'Load Loc 24 Index Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Quad24KBLK", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Anno24KQ", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Loc 24 Index", Item:=pLayerList
    'Load Geology Map100BT Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="Boreholes", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Bedrock_Topography", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad24K", Item:=odnrZoomLevelOhio
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Geology Map100BT", Item:=pLayerList
    'Load BT100K_Quad.mxd Project
    m_pProjectLayers.Add Key:="BT100K_Quad.mxd", Item:=pMapList


    'BG100K_Quad.mxd
    Set pMapList = New Dictionary
    'Load Location 100K Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="County_Bndry", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Location 100K", Item:=pLayerList
    'Load Loc 24 Index Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Quad24K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="Anno24KQ", Item:=odnrZoomLevelAll
    pMapList.Add Key:="Loc 24 Index", Item:=pLayerList
    'Load Geology Map100BG Map
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="BG Units Ply", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Units Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="Quad24K", Item:=odnrZoomLevelOhio
    pLayerList.Add Key:="Quad100K", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="State_Bndry", Item:=odnrZoomLevelAll
    pLayerList.Add Key:="BG Units Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Misc Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Structure Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Structure Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Line", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Pnt", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Mines Ply", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Anno", Item:=odnrZoomLevelNone
    pLayerList.Add Key:="BG Demarcation", Item:=odnrZoomLevelNone
    pMapList.Add Key:="Geology Map100BG", Item:=pLayerList
    'Load BG100K_Quad.mxd Project
    m_pProjectLayers.Add Key:="BG100K_Quad.mxd", Item:=pMapList

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadProjectLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadQuadSpecificLayerTypes()
    On Error GoTo ErrorHandler

    Dim pLayerList As Dictionary
    Set m_pQuadLayerTypes = New Dictionary
    
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="DRGs", Item:=g_strDRG_Path
    pLayerList.Add Key:="Bedrock Structure Scans", Item:=g_strSCAN_Path
    m_pQuadLayerTypes.Add Key:=odnr24K, Item:=pLayerList
    
    Set pLayerList = New Dictionary
    pLayerList.Add Key:="Base 100 Points", Item:=g_strBase_Path
    pLayerList.Add Key:="Base 100 Lines", Item:=g_strBase_Path
    pLayerList.Add Key:="Base 100 Polygons", Item:=g_strBase_Path
    m_pQuadLayerTypes.Add Key:=odnr100K, Item:=pLayerList
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadSpecificLayerTypes " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadLayerFilePaths()
    On Error GoTo ErrorHandler

    Dim strPath As String
    strPath = g_strGeoDir_Path & "Layer_Files\"
    Set m_pFilePaths = New Dictionary
    m_pFilePaths.Add Key:="Anno24KQ", Item:=strPath & "Anno24KQ.lyr"
    m_pFilePaths.Add Key:="Bedrock_Topography", Item:=strPath & "Bedrock_Topography.lyr"
    m_pFilePaths.Add Key:="BG Anno", Item:=strPath & "BG Anno.lyr"
    m_pFilePaths.Add Key:="BG Demarcation", Item:=strPath & "BG Demarcation.lyr"
    m_pFilePaths.Add Key:="BG Mines Line", Item:=strPath & "BG Mines Line.lyr"
    m_pFilePaths.Add Key:="BG Mines Ply", Item:=strPath & "BG Mines Ply.lyr"
    m_pFilePaths.Add Key:="BG Mines Pnt", Item:=strPath & "BG Mines Pnt.lyr"
    m_pFilePaths.Add Key:="BG Misc Pnt", Item:=strPath & "BG Misc Pnt.lyr"
    m_pFilePaths.Add Key:="BG Structure Line", Item:=strPath & "BG Structure Line.lyr"
    m_pFilePaths.Add Key:="BG Structure Pnt", Item:=strPath & "BG Structure Pnt.lyr"
    m_pFilePaths.Add Key:="BG Units Line", Item:=strPath & "BG Units Line.lyr"
    m_pFilePaths.Add Key:="BG Units Ply", Item:=strPath & "BG Units Ply.lyr"
    m_pFilePaths.Add Key:="BG Units Pnt", Item:=strPath & "BG Units Pnt.lyr"
    m_pFilePaths.Add Key:="Boreholes", Item:=strPath & "Boreholes.lyr"
    m_pFilePaths.Add Key:="County_Bndry", Item:=strPath & "County_Bndry.lyr"
    m_pFilePaths.Add Key:="Quad100K", Item:=strPath & "Quad100K.lyr"
    m_pFilePaths.Add Key:="Quad24K", Item:=strPath & "Quad24K.lyr"
    m_pFilePaths.Add Key:="Quad24K Index", Item:=strPath & "Quad24K Index.lyr"
    m_pFilePaths.Add Key:="Quad24KBLK", Item:=strPath & "Quad24KBLK.lyr"
    m_pFilePaths.Add Key:="State_Bndry", Item:=strPath & "State_Bndry.lyr"
    m_pFilePaths.Add Key:="Topo_Annotation", Item:=strPath & "Topo_Annotation.lyr"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadLayerFilePaths " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
