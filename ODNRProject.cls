VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRProject
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object representing the map project.
'               It serves various information to forms, buttons and other
'               objects in order to make the extension work.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pQuads As ODNRQuads
Private m_pBedrockLayers As ODNRBedrockLayers
Private m_pQuadLayers As ODNRQuadLayers
Private m_pStateLayers As ODNRStateLayers
Private m_pLayoutElements As Dictionary
Private m_pQuadScale As ODNRQuadScale
Private m_pProjectType As ODNRProjectType
Private m_blnIsloaded As Boolean
Private m_blnIsZoomedToQuads As Boolean
Private m_blnInDataView As Boolean
Private m_blnDocChanged As Boolean
Private m_blnDoResetLayout As Boolean
Private m_blnDoResetDataView As Boolean

Const c_strModuleName As String = "ODNRProject"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    m_pQuadScale = odnrScaleUnknown
    m_blnInDataView = True
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Me.ResetLayout
        Me.ResetDataView
    End If
    If Not (m_pLayoutElements Is Nothing) Then
        m_pLayoutElements.RemoveAll
        Set m_pLayoutElements = Nothing
    End If
    Set m_pStateLayers = Nothing
    Set m_pQuadLayers = Nothing
    Set m_pBedrockLayers = Nothing
    m_pQuads.RemoveAll
    Set m_pQuads = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get ProjectType() As ODNRProjectType
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then ProjectType = m_pProjectType
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectType " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadScale() As ODNRQuadScale
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then QuadScale = m_pQuadScale
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadScale " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get BedrockLayers() As ODNRBedrockLayers
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If Not (m_pBedrockLayers Is Nothing) Then
            Set BedrockLayers = m_pBedrockLayers
        Else
            Set BedrockLayers = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BedrockLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get StateLayers() As ODNRStateLayers
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If Not (m_pStateLayers Is Nothing) Then
            Set StateLayers = m_pStateLayers
        Else
            Set StateLayers = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".StateLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadLayers() As ODNRQuadLayers
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If Not (m_pQuadLayers Is Nothing) Then
            Set QuadLayers = m_pQuadLayers
        Else
            Set QuadLayers = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get InDataView() As Boolean
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        RefreshViewStatus
        InDataView = m_blnInDataView
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".InDataView " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get IsZoomedToQuadSelection() As Boolean
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then IsZoomedToQuadSelection = m_blnIsZoomedToQuads
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IsZoomedToQuadSelection " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ProjectName() As String
    On Error GoTo ErrorHandler
    
    ProjectName = gODNRProjectDb.ActiveProjectName
        
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ProjectMap(lngMapType As ODNRMapType) As IMap
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim strMapName As String
        If (lngMapType = odnrGeologyMap) Then
            strMapName = "Geology Map"
        ElseIf (lngMapType = odnrLocation24KMap) Then
            strMapName = "Location 24K"
        ElseIf (lngMapType = odnrLocation100KMap) Then
            strMapName = "Location 100K"
        End If
        Set ProjectMap = Me.GetMapByName(strMapName)
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectMap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function GetMapByName(strMapName As String) As IMap
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, lngIdx As Long
        Set pMxDoc = m_pApp.Document
        If (pMxDoc.Maps.Count > 0) And (strMapName <> "") Then
            For lngIdx = 0 To pMxDoc.Maps.Count - 1
                Set pMap = pMxDoc.Maps.Item(lngIdx)
                If (pMap.Name = strMapName) Then
                    Set GetMapByName = pMap
                    Exit For
                End If
            Next
        Else
            Set GetMapByName = Nothing
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".GetMapByName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Property Get QuadFeatureLayer() As IFeatureLayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.QuadScale = odnr24K) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Quad24K")
            Set QuadFeatureLayer = pStateLayer.ESRILayer
        ElseIf (Me.QuadScale = odnr100K) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Quad100K")
            Set QuadFeatureLayer = pStateLayer.ESRILayer
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadFeatureLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get GeologyUnitsPolygonLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Units Ply")
            Set GeologyUnitsPolygonLayer = pStateLayer.ESRILayer
        Else
            Set GeologyUnitsPolygonLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".GeologyUnitsPolygonLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
Public Property Get GeologyUnitsLineLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Units Line")
            Set GeologyUnitsLineLayer = pStateLayer.ESRILayer
        Else
            Set GeologyUnitsLineLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".GeologyUnitsLineLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get GeologyUnitsPointLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Units Pnt")
            Set GeologyUnitsPointLayer = pStateLayer.ESRILayer
        Else
            Set GeologyUnitsPointLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".GeologyUnitsPointLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
Public Property Get BTContourLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrTopography) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BT Contours")
            Set BTContourLayer = pStateLayer.ESRILayer
        Else
            Set BTContourLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BTContourLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get BTPointsLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrTopography) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BT Data Points")
            Set BTPointsLayer = pStateLayer.ESRILayer
        Else
            Set BTPointsLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BTPointsLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get AUMPointsLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrAUM) Then
'            Set pStateLayer = m_pStateLayers.GetLayerByName("AUM_PTS")
            Set pStateLayer = m_pStateLayers.GetLayerByName("Mine Opening")
            Set AUMPointsLayer = pStateLayer.ESRILayer
        Else
            Set AUMPointsLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".AUMPointsLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get AUMMinePointsLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrAUM) Then
'            Set pStateLayer = m_pStateLayers.GetLayerByName("AUM_PTS")
            Set pStateLayer = m_pStateLayers.GetLayerByName("Mine Location - Extent Unknown")
            Set AUMMinePointsLayer = pStateLayer.ESRILayer
        Else
            Set AUMMinePointsLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".AUMPointsLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get AUMMinesLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrAUM) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Underground Mine")
            Set AUMMinesLayer = pStateLayer.ESRILayer
        Else
            Set AUMMinesLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".AUMMinesLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get AUMMinesOverlapLayer() As ILayer
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrAUM) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Superimposed Underground Mine")
            Set AUMMinesOverlapLayer = pStateLayer.ESRILayer
        Else
            Set AUMMinesOverlapLayer = Nothing
        End If
    End If

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".AUMMinesOverlapLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get AUMHatchuredLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrAUM) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Underground Mine - Extent Partially Unknown")
            Set AUMHatchuredLayer = pStateLayer.ESRILayer
        Else
            Set AUMHatchuredLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".AUMHatchuredLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get GeologyUnitsContactLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Units Contacts")
            Set GeologyUnitsContactLayer = pStateLayer.ESRILayer
        Else
            Set GeologyUnitsContactLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".GeologyUnitsContactLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get GeologyFaciesLineLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Facies")
            Set GeologyFaciesLineLayer = pStateLayer.ESRILayer
        Else
            Set GeologyFaciesLineLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".GeologyFaciesLineLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get MinesPolygonLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Mines Ply")
            Set MinesPolygonLayer = pStateLayer.ESRILayer
        Else
            Set MinesPolygonLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".MinesPolygonLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get MinesLineLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Mines Line")
            Set MinesLineLayer = pStateLayer.ESRILayer
        Else
            Set MinesLineLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".MinesLineLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get MinesPointLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Mines Pnt")
            Set MinesPointLayer = pStateLayer.ESRILayer
        Else
            Set MinesPointLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".MinesPointLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get BGStructurePointLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Structure Pnt")
            Set BGStructurePointLayer = pStateLayer.ESRILayer
        Else
            Set BGStructurePointLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BGStructurePointLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get BGStructureLineLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Structure Line")
            Set BGStructureLineLayer = pStateLayer.ESRILayer
        Else
            Set BGStructureLineLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BGStructureLineLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get BGMiscPntLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrGeology) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("BG Misc Pnt")
            Set BGMiscPntLayer = pStateLayer.ESRILayer
        Else
            Set BGMiscPntLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".BGMiscLineLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get CountyLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        Set pStateLayer = m_pStateLayers.GetLayerByName("County_Bndry")
        Set CountyLayer = pStateLayer.ESRILayer
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".CountyLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get Quads() As ODNRQuads
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Set Quads = m_pQuads
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Quads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property


Public Sub OpenProject(pApp As IApplication, lngQuadScale As ODNRQuadScale, lngProjectType As ODNRProjectType)
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        MsgBox "Please unload the existing ODNR project before opening another project.", vbInformation, "ODNRProject.OpenProject"
        Exit Sub
    ElseIf (lngQuadScale <> odnrScaleUnknown) And (lngProjectType <> odnrProjectTypeUnknown) Then
        Set m_pApp = pApp
        m_pQuadScale = lngQuadScale
        m_pProjectType = lngProjectType
        m_blnIsloaded = True
        Set m_pStateLayers = New ODNRStateLayers
        m_pStateLayers.Load m_pApp, Me
        Set m_pQuads = New ODNRQuads
        m_pQuads.Load m_pApp, Me
        If (m_pQuadScale = odnr24K) Then
            If (m_pProjectType = odnrBedrockStructure) Or (m_pProjectType = odnrGeology) Then
                Set m_pBedrockLayers = New ODNRBedrockLayers
                m_pBedrockLayers.Load m_pApp, Me
            End If
        End If
        Set m_pQuadLayers = New ODNRQuadLayers
        m_pQuadLayers.Load m_pApp, Me
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".OpenProject " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ZoomToQuadsExtent()
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pActiveView As IActiveView, pDisplayTransform As IDisplayTransformation
        Dim pDispEnv As IEnvelope, pQuadEnv As IEnvelope
        Set pMxDoc = m_pApp.Document
        Set pActiveView = pMxDoc.ActiveView
        Set pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation
        Set pDispEnv = m_pQuads.ExtentEnvelope
        If Not (pDispEnv Is Nothing) Then
            Set pQuadEnv = pDispEnv
            pDispEnv.Expand 1.07, 1.07, True
            pDisplayTransform.VisibleBounds = pDispEnv
            If Not (m_pQuads.FocusQuad Is Nothing) And (Me.QuadScale = odnr24K) Then
                pQuadEnv.Expand 1.2, 1.2, True
            End If
            m_blnIsZoomedToQuads = True
        End If
        pActiveView.Refresh
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ZoomToQuadsExtent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ShowBedrockLayers()
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) And (Not m_pBedrockLayers Is Nothing) Then
        Dim pMxDoc As IMxDocument, pMap As IMap
        Dim pBedrockLayer As ODNRBedrockLayer, pLayer As ILayer
        Set pMxDoc = m_pApp.Document
        Set pMap = Me.ProjectMap(odnrGeologyMap)
        m_pBedrockLayers.Refresh
        m_pBedrockLayers.Reset
        Set pBedrockLayer = m_pBedrockLayers.NextLayer
        Do While Not pBedrockLayer Is Nothing
            Set pLayer = pBedrockLayer.ESRILayer
            pMap.AddLayer pLayer
            pMap.MoveLayer pLayer, pMap.LayerCount - 1
            Set pBedrockLayer = m_pBedrockLayers.NextLayer
        Loop
        pMxDoc.ActiveView.Refresh
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowBedrockLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ShowQuadLayers()
    On Error GoTo ErrorHandler
    
    Dim pFLayer As IFeatureLayer
    Dim pGFlayer As IGeoFeatureLayer

    If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap
        Dim pQuadLayer As ODNRQuadLayer, pLayer As ILayer, pAUMFLyr As IFeatureLayer
        Set pMxDoc = m_pApp.Document
        Set pMap = Me.ProjectMap(odnrGeologyMap)
        m_pQuadLayers.Refresh
        m_pQuadLayers.Reset
        Set pQuadLayer = m_pQuadLayers.NextLayer
        Do While Not pQuadLayer Is Nothing
            Set pLayer = pQuadLayer.ESRILayer
            pMap.AddLayer pLayer
            pMap.MoveLayer pLayer, pMap.LayerCount - 1
            Set pQuadLayer = m_pQuadLayers.NextLayer
        Loop
        'This displays the BT, AUM, and OGWELLS layers when zoomed into the quad
        'Jim McDonald, 20051207, 20060519
        If (Me.ProjectType = odnrTopography) Then
            BTPointsLayer.Visible = True
            BTContourLayer.Visible = True
        ElseIf (Me.ProjectType = odnrAUM) Then
            LoadAUMMinesOutline
            AUMHatchuredLayer.Visible = True
            AUMMinesLayer.Visible = True
            AUMPointsLayer.Visible = True
            AUMMinePointsLayer.Visible = True
            AUMMinesOverlapLayer.Visible = True
            Set pAUMFLyr = AUMMinesLayer
            Call LabelFeatures(pAUMFLyr)
        ElseIf (Me.ProjectType = odnrOGWells) Then
            LoadQuadOGWellLayers
            DirDrilledWellboreLayer.Visible = True
            OLSLayer.Visible = True
        End If
        pMxDoc.UpdateContents
        pMxDoc.ActiveView.Refresh
        
        'This IF statement calls the procedure LoadGeologyLayer to load the
        'Quad Geology Units Ply layer, Jim McDonald, 20051201
        If ((Me.ProjectType = odnrGeology) And Me.ProjectName <> "Bedrock Geology 24K for Editing") Then
            LoadQuadBGLayers
        ElseIf ((Me.ProjectType = odnrGeology) And Me.ProjectName = "Bedrock Geology 24K for Editing") Then
            Dim pEnumLyr As IEnumLayer
            Set pEnumLyr = pMap.Layers
            Set pLayer = pEnumLyr.Next
            Do Until pLayer Is Nothing
                If ((TypeOf pLayer Is IFeatureLayer) And (pLayer.Visible = False) And pLayer.Name Like "Editing*") Then
                    pLayer.Visible = True
                    If (pLayer.Name = "Editing-BG Units Ply") Then
                        pMap.MoveLayer pLayer, pMap.LayerCount - 1
                    End If
                End If
                Set pLayer = pEnumLyr.Next
            Loop
            pMxDoc.UpdateContents
            pMxDoc.ActiveView.Refresh
        End If
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowQuadLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ShowOverviewMapQuad()
    On Error GoTo ErrorHandler
    
    Dim pOverviewMap As IMap, pOverviewMapAlt As IMap, pActiveView As IActiveView
    Dim pQuadG As IGeometry, pQuadEnv As IEnvelope, pArea As IArea, pPolygon As IPointCollection
    Dim pRectElem As IFillShapeElement, pElem As IElement, pEnumElem As IEnumElement
    Dim pGraCont As IGraphicsContainer, pGraContSel As IGraphicsContainerSelect
    Dim pFill As IFillSymbol, pLine As ILineSymbol, pColor As IColor
    Dim pLayer As ILayer
    Dim pEnumLyr As IEnumLayer
    Dim pGeoFLayer As IGeoFeatureLayer
    Dim pFDef As IFeatureLayerDefinition
    Dim pFLayer As IFeatureLayer
    Dim pAnnoProps As IAnnotateLayerPropertiesCollection
    Dim pAnnoLayerProps As IAnnotateLayerProperties
    Dim pLabelEngine As ILabelEngineLayerProperties
    Dim pTextSym As ITextSymbol
    Dim pFTextSym As IFormattedTextSymbol
    Dim pFont As stdole.IFontDisp
    Dim pDs As IDataset
    Dim pFClass As IFeatureClass
    Dim pWs As IWorkspace
    
    If (m_pQuads.FocusQuad Is Nothing) Then
        Set pQuadEnv = m_pQuads.ExtentEnvelope
        Set pPolygon = New Polygon
        pPolygon.AddPoint pQuadEnv.UpperLeft
        pPolygon.AddPoint pQuadEnv.UpperRight
        pPolygon.AddPoint pQuadEnv.LowerRight
        pPolygon.AddPoint pQuadEnv.LowerLeft
        pPolygon.AddPoint pQuadEnv.UpperLeft
        Set pQuadG = pPolygon
    Else
        Set pQuadG = m_pQuads.FocusQuad.QuadBoundary
    End If
    If (Me.QuadScale = odnr24K) Then
        Set pOverviewMap = Me.ProjectMap(odnrLocation24KMap)
    ElseIf (Me.QuadScale = odnr100K) Then
        Set pOverviewMap = Me.ProjectMap(odnrLocation100KMap)
        Set pOverviewMapAlt = Me.ProjectMap(odnrLocation24KMap)
    End If
    
    If (Not pQuadG Is Nothing) And (Not pOverviewMap Is Nothing) Then
        Set pColor = New RgbColor
        pColor.RGB = RGB(0, 0, 0)
        
        Set pQuadEnv = pQuadG.Envelope
        Set pElem = New PolygonElement
        pElem.Geometry = pQuadG
        Set pRectElem = pElem
        Set pFill = pRectElem.Symbol
        Set pLine = pRectElem.Symbol.Outline
        pFill.Color = pColor
        pLine.Color = pColor
        pFill.Outline = pLine
        pRectElem.Symbol = pFill
        
        Set pActiveView = pOverviewMap
        'delete all graphics in this index map
        Set pGraCont = pActiveView
        pGraCont.DeleteAllElements
             
        'add the new one
        pGraCont.AddElement pElem, 0
        pActiveView.Refresh
        Set pGraContSel = pGraCont
        pGraContSel.UnselectAllElements
                    
        If Not (pOverviewMapAlt Is Nothing) Then
            Set pActiveView = pOverviewMapAlt
            Set pArea = pQuadEnv
            pQuadEnv.CenterAt pArea.Centroid
            pActiveView.Extent = pQuadEnv
            pActiveView.Refresh
        End If
        
        Set pEnumLyr = pOverviewMap.Layers
        pEnumLyr.Reset
        Set pLayer = pEnumLyr.Next
        Do While Not pLayer Is Nothing
            If pLayer.Name = "County_Bndry" Then
                Set pFDef = pLayer
                Set pFLayer = pLayer
                Set pFClass = pFLayer.FeatureClass
                Set pDs = pFClass
                Set pWs = pDs.Workspace
                If (pWs.Type = esriRemoteDatabaseWorkspace) Then
                    pFDef.DefinitionExpression = "STATUS = 'Land' OR STATUS = 'Island'"
                ElseIf (pWs.Type = esriLocalDatabaseWorkspace) Then
                    pFDef.DefinitionExpression = "[STATUS] = 'Land' OR [STATUS] = 'Island'"
                ElseIf (pWs.Type = esriFileSystemWorkspace) Then
                    pFDef.DefinitionExpression = """STATUS"" = 'Land' OR ""STATUS"" = 'Island'"
                End If
                Exit Do
            End If
            Set pLayer = pEnumLyr.Next
        Loop
    End If
    
    'Added 20060028, Jim McDonald
    'This section will label the 1:24,000 scale quadrangles in the 24K
    'Overview Map.
    If (Not pQuadG Is Nothing) And (Not pOverviewMapAlt Is Nothing) Then
        Set pEnumLyr = pOverviewMapAlt.Layers
        Set pLayer = pEnumLyr.Next
        Do While Not pLayer Is Nothing
            If pLayer.Name = "Quad24K" Then
                Set pGeoFLayer = pLayer
                Exit Do
            End If
            Set pLayer = pEnumLyr.Next
        Loop
        
        'Set the Font
        Set pFont = New stdole.StdFont
        pFont.Name = "Arial"
        
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        pTextSym.Size = 4
'        pTextSym.Angle = 45
        
        'Create a new formatted text symbol.
        Set pFTextSym = pTextSym
'        pFTextSym.Direction = esriTDAngle
        
        'Create a spatial query
        Dim pSQf As ISpatialFilter
        Dim pFc As IFeatureCursor
        Dim pQuadFc As IFeatureClass
        Set pQuadFc = pGeoFLayer.FeatureClass
        Set pSQf = New SpatialFilter
        pSQf.SpatialRel = esriSpatialRelIntersects
        Set pSQf.Geometry = pQuadG
        Set pFc = pQuadFc.Search(pSQf, True)
        
        'Cycle through the feature cursor and create new graphic elements
        Dim pFeature As IFeature
        Dim pPoint As IPoint
        Dim pGC As IGraphicsContainer
        
        Set pGC = pOverviewMapAlt
        
        Dim lngFieldIdx As Long
        Dim strQuadName As String
        Dim pElement As IElement
        Dim pAreaQuad As IArea
        Dim pTextElement As ITextElement
        Dim pTrans2D As ITransform2D
        
        
        lngFieldIdx = pFc.FindField("QUADNAME")
                
        Set pFeature = pFc.NextFeature
        Do While Not pFeature Is Nothing
            strQuadName = pFeature.Value(lngFieldIdx)
            Set pAreaQuad = pFeature.Shape
            pFTextSym.Text = strQuadName
            Set pTextElement = New TextElement
            pTextElement.Symbol = pFTextSym
            Set pElement = pTextElement
            pTextElement.Text = strQuadName
            pElement.Geometry = pAreaQuad.Centroid
            Set pTrans2D = pElement
            pTrans2D.Rotate pElement.Geometry, 45
            pGC.AddElement pElement, 0
            Set pFeature = pFc.NextFeature
        Loop
        
        Set pEnumLyr = pOverviewMapAlt.Layers
        pEnumLyr.Reset
        Set pLayer = pEnumLyr.Next
        Do While Not pLayer Is Nothing
            If pLayer.Name = "County_Bndry" Then
                Set pFDef = pLayer
                Set pFLayer = pLayer
                Set pFClass = pFLayer.FeatureClass
                Set pDs = pFClass
                Set pWs = pDs.Workspace
                If (pWs.Type = esriRemoteDatabaseWorkspace) Then
                    pFDef.DefinitionExpression = "STATUS = 'Land' OR STATUS = 'Island'"
                ElseIf (pWs.Type = esriLocalDatabaseWorkspace) Then
                    pFDef.DefinitionExpression = "[STATUS] = 'Land' OR [STATUS] = 'Island'"
                ElseIf (pWs.Type = esriFileSystemWorkspace) Then
                    pFDef.DefinitionExpression = """STATUS"" = 'Land' OR ""STATUS"" = 'Island'"
                End If
                Exit Do
            End If
            Set pLayer = pEnumLyr.Next
        Loop
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowOverviewMapQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ResetLayout()
    On Error GoTo ErrorHandler
    'FUNCTION: this procedure resets all the elements in the layout,
    'zooms out to the State of Ohio and unloads the image layers
    If (m_blnIsloaded) And (m_blnDoResetLayout) Then
        Dim pMxDoc As IMxDocument, pPageLayout As IPageLayout, pGraCont As IGraphicsContainer
        Dim pODNRLayoutTextEl As ODNRLayoutTextElement, vKey As Variant
        Dim pODNRLegendEl As ODNRLayoutLegendElement
        Dim pColorBlack As IRgbColor
        Dim pOverviewMapAlt As IMap
        Dim pGC As IGraphicsContainer
        
        If Not (m_pLayoutElements Is Nothing) Then
            Set pMxDoc = m_pApp.Document
            Set pPageLayout = pMxDoc.PageLayout
            Set pGraCont = pPageLayout
            pGraCont.Reset
            Set pColorBlack = New RgbColor
            pColorBlack.RGB = RGB(0, 0, 0)
            For Each vKey In m_pLayoutElements.Keys
                If (TypeOf m_pLayoutElements.Item(vKey) Is ODNRLayoutTextElement) Then
                    Set pODNRLayoutTextEl = m_pLayoutElements.Item(vKey)
                    pODNRLayoutTextEl.SetTextToKey
                    If (m_pQuads.QuadCount > 1) Then Set pODNRLayoutTextEl.SymbolColor = pColorBlack
                    pODNRLayoutTextEl.UpdateGraphicsContainer
                    m_pLayoutElements.Remove vKey
                    Set pODNRLayoutTextEl = Nothing
                ElseIf (TypeOf m_pLayoutElements.Item(vKey) Is ODNRLayoutLegendElement) Then
                    Set pODNRLegendEl = m_pLayoutElements.Item(vKey)
                    pODNRLegendEl.SetLegendTitleTextToKey
                    pODNRLegendEl.ClearLegend
                    m_pLayoutElements.Remove vKey
                    Set pODNRLegendEl = Nothing
                End If
            Next
        End If
        'Empty the list once everything has been reset and then make the list nothing
        Set m_pLayoutElements = Nothing
        m_blnDoResetLayout = False
        
        If (Me.QuadScale = odnr100K) Then
            Set pOverviewMapAlt = Me.ProjectMap(odnrLocation24KMap)
            Set pGC = pOverviewMapAlt
            pGC.DeleteAllElements
        End If
    
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ResetLayout " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub ResetDataView()
    On Error GoTo ErrorHandler
    'FUNCTION: this procedure resets all the elements in the layout,
    'zooms out to the State of Ohio and unloads the image layers
    
    Dim pMxDoc As IMxDocument, pMap As IMap, pActiveView As IActiveView
    Dim pLayer As ILayer, pEnumLayer As IEnumLayer, pFeatureLayer As IFeatureLayer
    Dim strMapScale As String
    Dim pOverviewMapAlt As IMap
    Dim pGC As IGraphicsContainer

    Set pMxDoc = m_pApp.Document
    Set pMap = pMxDoc.FocusMap
    Set pActiveView = pMxDoc.ActiveView
    
    Set pMap = Me.ProjectMap(odnrGeologyMap)
    Set pActiveView = pMap
    pMap.ClearSelection
    Set pEnumLayer = pMap.Layers
    pEnumLayer.Reset
    Set pLayer = pEnumLayer.Next
    Do While Not pLayer Is Nothing
        If TypeOf pLayer Is IRasterLayer Then
            pMap.DeleteLayer pLayer
        ElseIf (TypeOf pLayer Is IGroupLayer) Then
            pMap.DeleteLayer pLayer
        ElseIf (TypeOf pLayer Is IFeatureLayer) Then
            Set pFeatureLayer = pLayer
            pFeatureLayer.Visible = False
            pFeatureLayer.Selectable = False
            If (pLayer.Name = "State_Bndry") Then
                pFeatureLayer.Visible = True
            ElseIf (pLayer.Name = "Quad24K") And (Me.QuadScale = odnr24K) Then
                pFeatureLayer.Selectable = True
                pFeatureLayer.Visible = True
            ElseIf (pLayer.Name = "Quad100K") And (Me.QuadScale = odnr100K) Then
                pFeatureLayer.Selectable = True
                pFeatureLayer.Visible = True
            End If
            If (pLayer.Name = "Quad Geology Units Ply") Then
                UnloadGeologyLayer
                pMap.DeleteLayer pLayer
            End If
            If (pLayer.Name = "Quad Oil & Gas Wells") Then
                pMap.DeleteLayer pLayer
            End If
            If (pLayer.Name = "Underground Mines - Outline") Then
                pMap.DeleteLayer pLayer
            End If
            If (Me.QuadScale = odnr100K) Then
                strMapScale = Right(pLayer.Name, 2) 'delete the 100K basemap layers
                Select Case strMapScale
                    Case "_A"
                        pMap.DeleteLayer pLayer
                    Case "_X"
                        pMap.DeleteLayer pLayer
                    Case "_P"
                        pMap.DeleteLayer pLayer
                    Case "_L"
                        pMap.DeleteLayer pLayer
                End Select
            End If
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    'This turns off display the BT and AUM layers when zoomed out to the state
    'Jim McDonald, 20051207
    If (Me.ProjectType = odnrTopography) Then
        BTPointsLayer.Visible = False
        BTContourLayer.Visible = False
    ElseIf (Me.ProjectType = odnrAUM) Then
        AUMHatchuredLayer.Visible = False
        AUMMinesLayer.Visible = False
        AUMMinePointsLayer.Visible = False
        AUMPointsLayer.Visible = False
        AUMMinesOverlapLayer.Visible = False
    ElseIf (Me.ProjectType = odnrOGWells) Then
        OGWellsLayer.Visible = False
        DirDrilledWellboreLayer.Visible = False
        OLSLayer.Visible = False
    End If
    
    If (Me.QuadScale = odnr100K) Then
        Set pOverviewMapAlt = Me.ProjectMap(odnrLocation24KMap)
        Set pGC = pOverviewMapAlt
        pGC.DeleteAllElements
    End If
    
    pMxDoc.UpdateContents
    
    pActiveView.Extent = pMxDoc.ActiveView.FullExtent
    pActiveView.Refresh
    
    If Not (m_pQuads Is Nothing) Then
        If (m_pQuads.QuadCount > 0) Then
            m_pQuads.RemoveAll
        End If
    End If
    m_blnIsZoomedToQuads = False
    frmToolbarControls.cboQuad.Text = "Select a Quadrangle"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ResetDataView " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub GoDataView()
    On Error GoTo ErrorHandler
    
    If (m_blnInDataView = False) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, pActiveView As IActiveView
        Me.ResetLayout
        Set pMxDoc = m_pApp.Document
        Set pMxDoc.ActiveView = Me.ProjectMap(odnrGeologyMap)
        m_blnInDataView = True
        pMxDoc.ActiveView.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".GoDataView " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub GoLayout()
    On Error GoTo ErrorHandler
    'find out which project we are on and which template
    If (m_blnInDataView) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, pActiveView As IActiveView
        Dim pRow As IRow, pPageLayout As IPageLayout, pGraCont As IGraphicsContainer
        Dim pElem As IElement, pTextElement As ITextElement
        Dim pGroupElement As IGroupElement
        Dim blnSingleQuad As Boolean, blnListCounties As Boolean
        Dim pCountyList As Collection, lngIdx As Long
        Dim pFWs As IFeatureWorkspace, pWsFact As IWorkspaceFactory, pTypesTbl As ITable
        Dim pQf As IQueryFilter, pTypesRow As IRow, pCursor As ICursor
        Dim pQuadLayer As ODNRQuadLayer, pBedrockLayer As ODNRBedrockLayer
        Dim pGlyr As IGroupLayer, pClyr As ICompositeLayer
        
        Dim pGFlyr As IGeoFeatureLayer, pLayer As ILayer, pFLayer As IFeatureLayer
        Dim pUniqueValueRend As IUniqueValueRenderer, pEnumLayer As IEnumLayer
        Dim lngValueIdx As Long
        
        Dim lngLyrIdx As Long, strBedrockType As String
        Dim dblWidth As Double, dblHeight As Double
        
        Dim strText As String, strMapType As String
        Dim strTitle As String, strResYear As String
        Dim strBedrock As String, strQuadName As String 'remember to preface with the word Map
        Dim strLatN As String, strLatS As String
        Dim strLongW As String, strLongE As String
        Dim strAuthor As String, strCounties As String
        Dim strInterval As String
        
        Dim pODNRLayoutTextEl As ODNRLayoutTextElement, pColor As IRgbColor
        Dim pODNRLegendEl As ODNRLayoutLegendElement
        Dim pQuad As ODNRQuad
        
        Set pMxDoc = m_pApp.Document
        Set pMap = pMxDoc.FocusMap
        Set pActiveView = pMap
        
        blnSingleQuad = False
        If Not (m_pQuads.FocusQuad Is Nothing) Then
            blnSingleQuad = True
        End If
        
        'For quadrangles that have metadata record, read the metadata database record
        If (blnSingleQuad) Then
            Set pRow = m_pQuads.QuadLayoutTextRecord
            If (pRow Is Nothing And Me.ProjectType <> odnrOGWells) Then 'Edited to include oil and gas wells titles, 20060520
                MsgBox "Unable to locate title record to customize layout text."
                Exit Sub
            ElseIf (pRow Is Nothing And Me.ProjectType = odnrOGWells) Then
                'Continue on with generating Oil- and-Gas Wells Map
            Else
                If (pRow.Fields.FindField("TITLE") > -1) Then
                    strTitle = FieldValue(pRow, pRow.Fields.FindField("TITLE"))
                    pMxDoc.PageLayout.Page.QuerySize dblWidth, dblHeight
                    If (dblHeight > dblWidth) Then strTitle = ParseString(70, strTitle)
                End If
                If (pRow.Fields.FindField("NLAT") > -1) Then strLatN = FieldValue(pRow, pRow.Fields.FindField("NLAT"))
                If (pRow.Fields.FindField("SLAT") > -1) Then strLatS = FieldValue(pRow, pRow.Fields.FindField("SLAT"))
                If (pRow.Fields.FindField("WLONG") > -1) Then strLongW = FieldValue(pRow, pRow.Fields.FindField("WLONG"))
                If (pRow.Fields.FindField("ELONG") > -1) Then strLongE = FieldValue(pRow, pRow.Fields.FindField("ELONG"))
                If (Me.QuadScale = odnr24K) Then
                    If (pRow.Fields.FindField("TYPE") > -1) Then strMapType = FieldValue(pRow, pRow.Fields.FindField("TYPE"))
                    If (pRow.Fields.FindField("REVISION_DATE") > -1) Then strResYear = FieldValue(pRow, pRow.Fields.FindField("REVISION_DATE"))
                    If (pRow.Fields.FindField("AUTHOR") > -1) Then strAuthor = FieldValue(pRow, pRow.Fields.FindField("AUTHOR"))
                    If (pRow.Fields.FindField("QUAD") > -1) Then strQuadName = "Map " & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
                    If (Me.ProjectType = odnrBedrockStructure) Then
                        'Added section, to be used in the future once BS is loaded into SDE. 20060106, Jim McDonald
                        'If (g_strSDEPGB = "SDE") Then
                        '    pFWs = ODNRGeologyWorkspace
                        'ElseIf (g_strSDEPGB = "PGB") Then
                        '    Set pWsFact = New AccessWorkspaceFactory
                        '    Set pFWs = pWsFact.OpenFromFile(g_strBedrockDB_Path, 0)
                        'End If
                        Set pWsFact = New AccessWorkspaceFactory
                        Set pFWs = pWsFact.OpenFromFile(g_strBedrockDB_Path, 0)
                        Set pTypesTbl = pFWs.OpenTable("QuadUnitCodes")
                        If Not (pFWs Is Nothing) Then
                            Set pBedrockLayer = m_pBedrockLayers.GetLayerByName("BS Contours")
                            Set pGlyr = pBedrockLayer.ESRILayer
                            Set pClyr = pGlyr
                            For lngLyrIdx = 0 To pClyr.Count - 1
                                Set pLayer = pClyr.Layer(lngLyrIdx)
                                If (pLayer.Visible) Then
                                    strBedrockType = Left(pLayer.Name, Len(pLayer.Name) - 8)
                                    Exit For
                                End If
                            Next
                            Set pQf = New QueryFilter
                            pQf.WhereClause = "[QUAD_ID] = '" & m_pQuads.FocusQuad.ODNRQuadId & "' And [UNIT_CODE] = '" & strBedrockType & "'"
                            Set pCursor = pTypesTbl.Search(pQf, False)
                            Set pTypesRow = pCursor.NextRow
                            strInterval = CStr(pTypesRow.Value(3)) & " foot Contour Interval"
                        End If
                    End If
                ElseIf (Me.QuadScale = odnr100K) Then
                    If (pRow.Fields.FindField("TYPE") > -1) Then strMapType = FieldValue(pRow, pRow.Fields.FindField("TYPE"))
                    If (pRow.Fields.FindField("AUTHOR") > -1) Then strAuthor = FieldValue(pRow, pRow.Fields.FindField("AUTHOR"))
                    If (pRow.Fields.FindField("REV_DATE") > -1) Then strResYear = FieldValue(pRow, pRow.Fields.FindField("REV_DATE"))
                    If (pRow.Fields.FindField("QUAD") > -1) Then strQuadName = "Map " & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
                End If
            End If
        End If
    
        Me.QuadFeatureLayer.Visible = True
        
        'Go to layout
        Set pPageLayout = pMxDoc.PageLayout
        Set pMxDoc.ActiveView = pPageLayout
        Set pMxDoc = m_pApp.Document
        
        'Set the map layout scale
        Set pActiveView = Me.ProjectMap(odnrGeologyMap)
        If (Me.QuadScale = odnr24K) Then
            pActiveView.FocusMap.MapScale = 24000
        ElseIf (Me.QuadScale = odnr100K) Then
            pActiveView.FocusMap.MapScale = 100000
        End If
        
        m_blnInDataView = False
        m_blnDoResetLayout = True
        Set pActiveView = pMxDoc.PageLayout
        Set pGraCont = pPageLayout
        pGraCont.Reset
        
        Set m_pLayoutElements = New Dictionary
        
        'Build the title for the floating 100K box
        Dim strQuad100KNames As String
        Dim intQuadCnt As Integer
        Dim strQuadNamesArray() As String
        Dim intQuadNameCount As Integer
        Dim intNN As Integer
        Dim intN As Integer
        Dim intK As Integer
        Dim intJJ As Integer
        Dim intL As Integer
        Dim strTemp As String
        If (Me.QuadScale = odnr100K) Then
            If (m_pQuads.FocusQuad Is Nothing) Then
                m_pQuads.Reset
                Set pQuad = m_pQuads.NextQuad
                ReDim strQuadNamesArray(m_pQuads.QuadCount - 1) As String
                intQuadNameCount = 0
                'Place the quad names into a dynamic array.  The dynamic array will be used for sorting.
                Do While Not pQuad Is Nothing
                    strQuadNamesArray(intQuadNameCount) = pQuad.QuadName
                    intQuadNameCount = intQuadNameCount + 1
                    Set pQuad = m_pQuads.NextQuad
                Loop
                'Do a bubble sort
                intN = m_pQuads.QuadCount - 1
                intNN = intN - 1
                For intK = 0 To intNN
                    intJJ = intN - intK
                    For intL = 0 To intJJ - 1
                        If (strQuadNamesArray(intL) < strQuadNamesArray(intL + 1)) Then GoTo SKIPFOR
                        strTemp = strQuadNamesArray(intL)
                        strQuadNamesArray(intL) = strQuadNamesArray(intL + 1)
                        strQuadNamesArray(intL + 1) = strTemp
SKIPFOR:
                    Next intL
                Next intK
                strQuad100KNames = strQuadNamesArray(0)
                For intQuadNameCount = 1 To m_pQuads.QuadCount - 2
                    strQuad100KNames = strQuad100KNames & ", " & strQuadNamesArray(intQuadNameCount)
                Next intQuadNameCount
                strQuad100KNames = strQuad100KNames & " and " & strQuadNamesArray(m_pQuads.QuadCount - 1) & " 30x60 minute quadrangles"
            End If
        End If
'        Dim strQuad100KNames As String
'        Dim intQuadCnt As Integer
'        If (Me.QuadScale = odnr100K) Then
'            If (m_pQuads.FocusQuad Is Nothing) Then
'                m_pQuads.Reset
'                Set pQuad = m_pQuads.NextQuad
'                intQuadCnt = 0
'                Do While Not pQuad Is Nothing
'                    If (intQuadCnt = 0) Then
'                        strQuad100KNames = pQuad.QuadName
'                    ElseIf (intQuadCnt = m_pQuads.QuadCount - 1) Then
'                        strQuad100KNames = strQuad100KNames & ", and " & pQuad.QuadName & " 30x60 minute quadrangles"
'                    Else
'                        strQuad100KNames = strQuad100KNames & ", " & pQuad.QuadName
'                    End If
'                    intQuadCnt = intQuadCnt + 1
'                    Set pQuad = m_pQuads.NextQuad
'                Loop
'            End If
'        End If

        
        'When a value is missing from the metadata record, set the color of the layout element to white
        'This hides the element on the layout
        Set pColor = New RgbColor
        pColor.RGB = RGB(255, 255, 255)
        
        'Update the TextElements in the Layout with the values from the metadata database
        Dim pMapFrame As IMapFrame
        Dim pMapGeom As IGeometry

        Set pElem = pGraCont.Next
        Do While Not pElem Is Nothing
            'This changes the MapFrame geometry so as to fit the quad boundaries.
            'Currently in prototype mode. 20060601, Jim McDonald
'            If TypeOf pElem Is IMapFrame Then
'                Set pMapFrame = pElem
'                If (pMapFrame.Map.Name = pMap.Name) Then
'                    pElem.Geometry = ChangeMapFrameGeometry(pMapFrame)
'                End If
'            End If
            If TypeOf pElem Is ITextElement Then
                Set pTextElement = pElem
                strText = pTextElement.Text
                If (strText = "Revised 02/2002") Then
                    If (blnSingleQuad) Then
                        If strResYear <> "" Then
                            strResYear = "Revised " & strResYear
                            pTextElement.Text = strResYear
                            pGraCont.UpdateElement pElem
                        Else
                            strResYear = strText
                        End If
                    End If
                End If
                Set pODNRLayoutTextEl = New ODNRLayoutTextElement
                pODNRLayoutTextEl.LoadElement pElem, strText, pGraCont
                Select Case strText
                    Case "Northing1"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLatN
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Easting1"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLongW
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Northing2"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLatN
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Easting2"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLongE
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Northing3"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLatS
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Easting3"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLongW
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Northing4"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLatS
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Easting4"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strLongE
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Mapquadrangle"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strQuadName
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "quadname"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strQuadName
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                    Case "Revised 02/2002"
                        If (Me.ProjectType <> odnrOGWells) Then
                            If (blnSingleQuad) Then
                                pODNRLayoutTextEl.ElementText = strResYear
                            Else
                                Set pODNRLayoutTextEl.SymbolColor = pColor
                            End If
                        ElseIf (Me.ProjectType = odnrOGWells) Then
                            If (blnSingleQuad) Then
                                pODNRLayoutTextEl.ElementText = "Printed " & Date
                            Else
                                Set pODNRLayoutTextEl.SymbolColor = pColor
                            End If
                        End If
                    Case "Author"
                        If (blnSingleQuad) Then
                                pODNRLayoutTextEl.ElementText = strAuthor
                        Else
                                pODNRLayoutTextEl.ElementText = "ODGS - Geologic Mapping Group"
                        End If
                    Case "Title Line 1"
                        If (blnSingleQuad) Then
                            If (Me.ProjectType = odnrOGWells) Then
                                pODNRLayoutTextEl.ElementText = UCase("Oil and Gas Wells of the " & m_pQuads.FocusQuad.QuadName & " quadrangle")
                            Else
                                pODNRLayoutTextEl.ElementText = UCase(strTitle)
                            End If
                        Else
                            If (Me.ProjectType = odnrGeology) Then
                                If (Me.QuadScale = odnr100K) Then
                                    strQuad100KNames = "Bedrock geology of portions of the " & strQuad100KNames
                                    If (dblHeight > dblWidth) Then strQuad100KNames = ParseString(70, strQuad100KNames)
                                    pODNRLayoutTextEl.ElementText = UCase(strQuad100KNames)
                                Else
                                    pODNRLayoutTextEl.ElementText = UCase("Bedrock Geology")
                                End If
                            ElseIf (Me.ProjectType = odnrTopography) Then
                                If (Me.QuadScale = odnr100K) Then
                                    strQuad100KNames = "Bedrock topography of portions of the " & strQuad100KNames
                                    If (dblHeight > dblWidth) Then strQuad100KNames = ParseString(70, strQuad100KNames)
                                    pODNRLayoutTextEl.ElementText = UCase(strQuad100KNames)
                                Else
                                    pODNRLayoutTextEl.ElementText = UCase("Bedrock Topography")
                                End If
                            ElseIf (Me.ProjectType = odnrBedrockStructure) Then
                                pODNRLayoutTextEl.ElementText = UCase("Bedrock Structure")
                            ElseIf (Me.ProjectType = odnrOGWells) Then
                                pODNRLayoutTextEl.ElementText = UCase("Oil-and-Gas Wells")
                            End If
                        End If
                    Case "Contour Interval"
                        If (blnSingleQuad) Then
                            pODNRLayoutTextEl.ElementText = strInterval
                        Else
                            Set pODNRLayoutTextEl.SymbolColor = pColor
                        End If
                End Select
                
                m_pLayoutElements.Add Key:=strText, Item:=pODNRLayoutTextEl
                
            ElseIf TypeOf pElem Is IGroupElement Then
                Set pEnumLayer = pMap.Layers
                Set pLayer = pEnumLayer.Next
                Do While Not pLayer Is Nothing
                    If (pLayer.Name = "Quad Geology Units Ply" Or pLayer.Name = "Editing-BG Units Ply") Then
                        Set pFLayer = pLayer
                        Set pGFlyr = pFLayer
                        Exit Do
                    End If
                    Set pLayer = pEnumLayer.Next
                Loop
                If (Not pGFlyr Is Nothing) Then
                    Set pGroupElement = pElem
                    If (pGroupElement.ElementCount = 2) Then
                        If (TypeOf pGroupElement.Element(0) Is ITextElement) Then
                            Set pTextElement = pGroupElement.Element(0)
                        ElseIf (TypeOf pGroupElement.Element(1) Is ITextElement) Then
                            Set pTextElement = pGroupElement.Element(1)
                        End If
                        If Not (pTextElement Is Nothing) Then
                            'Changed to EXPLANATION, 20051005, James McDonald
'                            strText = "Bedrock Geology Units Legend"
                            strText = "EXPLANATION"
                            If (pTextElement.Text = strText) Then
                                Set pODNRLegendEl = New ODNRLayoutLegendElement
                                pODNRLegendEl.LoadElement pElem, strText, 0.3, 0.5, 10, "Arial", pGraCont
                                'Added new section to detect Bedrock Geology 24K for Editing
                                If ((Me.ProjectType = odnrGeology) And Me.ProjectName = "Bedrock Geology 24K for Editing") Then
                                    Set pUniqueValueRend = EditingBGUnitsPolyUVR(pLayer)
                                Else
                                    Set pUniqueValueRend = pGFlyr.Renderer
                                End If
                                'End of Added Section
'                                Set pUniqueValueRend = pGFlyr.Renderer
                                For lngValueIdx = 0 To pUniqueValueRend.ValueCount - 1
                                    pODNRLegendEl.AddLegendItem pUniqueValueRend.Symbol(pUniqueValueRend.Value(lngValueIdx)), pUniqueValueRend.Value(lngValueIdx) & ": " & pUniqueValueRend.Description(pUniqueValueRend.Value(lngValueIdx))
                                Next
                                pODNRLegendEl.BuildLegend
                                pODNRLegendEl.UpdateGraphicsContainer
                                m_pLayoutElements.Add Key:=strText, Item:=pODNRLegendEl
                            End If
                        End If
                    End If
                End If
            End If
            Set pElem = pGraCont.Next
        Loop
        
        ShowOverviewMapQuad
        pActiveView.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".GoLayout " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub RefreshViewStatus()
    On Error GoTo ErrorHandler
    
    Dim pMxDoc As IMxDocument, pActiveView As IActiveView
    If (Not m_pApp Is Nothing) Then
        Set pMxDoc = m_pApp.Document
        Set pActiveView = pMxDoc.ActiveView
        m_blnInDataView = True
        If (TypeOf pActiveView Is IPageLayout) Then
            m_blnInDataView = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".RefreshViewStatus " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadQuadBGLayers()
'******************************************************************************
'
'   Program:    LoadQuadBGLayers
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       November 30, 2005
'   Purpose:    Creates all the the individual quadrangle BG layers.  Renders
'               them, and then groups most of the layers.  Finally, the layers
'               are loaded
'   Called from: ODNRProject.ShowQuadLayers
'
'******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim pNewFlyr As IFeatureLayer, pGFlyr As IGeoFeatureLayer
    Dim pLayer As ILayer, pFLayer As IFeatureLayer
    Dim pUniqueValueRend As IUniqueValueRenderer, pFlyrDef As IFeatureLayerDefinition
    Dim strDefQry As String, pDs As IDataset, pFWs As IFeatureWorkspace
    Dim pQDef As IQueryDef, pRow As IRow, pCursor As ICursor, pNewUniqueValueRend As IUniqueValueRenderer
    Dim pUniqueUnitCodes As Dictionary, pCodeOrder As Dictionary, vKey As Variant
    Dim pTblSort As ITableSort, pQf As IQueryFilter
    Dim lngCSDAGEFieldIdx As Long, lngSTRATCODEFieldIdx As Long
    Dim lngNAMEFieldIdx As Long, lngUNITCODEFieldIdx As Long
    Dim pSym As ISymbol, pRendererPropPage As IRendererPropertyPage
    Dim pDataLayer As IDataLayer, pDsName As IDatasetName
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)

    ODNR_Common.ShowMessage "Loading Geology Layer...", False
    
    'Start adding BG map layers
    Dim pBGUnitsPlyLyr As IFeatureLayer, pQuadBGUnitsPlyLyr As IFeatureLayer
    Dim pBGUnitsLineLyr As IFeatureLayer, pQuadBGUnitsLineLyr As IFeatureLayer
    Dim pBGUnitsPntLyr As IFeatureLayer, pQuadBGUnitsPntLyr As IFeatureLayer
    Dim pBGContactsLineLyr As IFeatureLayer, pQuadBGContactsLineLyr As IFeatureLayer
    Dim pBGMinesPlyLyr As IFeatureLayer, pQuadBGMinesPlyLyr As IFeatureLayer
    Dim pBGMinesLineLyr As IFeatureLayer, pQuadBGMinesLineLyr As IFeatureLayer
    Dim pBGMinesPntLyr As IFeatureLayer, pQuadBGMinesPntLyr As IFeatureLayer
    Dim pBGMiscPntLyr As IFeatureLayer, pQuadBGMiscPntLyr As IFeatureLayer
    Dim pBGStructurePntLyr As IFeatureLayer, pQuadBGStructurePntLyr As IFeatureLayer
    Dim pBGStructureLineLyr As IFeatureLayer, pQuadBGStructureLineLyr As IFeatureLayer
    Dim pBGFaciesLineLyr As IFeatureLayer, pQuadBGFaciesLineLyr As IFeatureLayer
    
    Dim pQuadGrpLyr As IGroupLayer
    Set pQuadGrpLyr = New GroupLayer
    pQuadGrpLyr.Name = "Quad BG Layers"
    pQuadGrpLyr.Visible = True
    pQuadGrpLyr.Expanded = False
    
    'Create the Quad BG Units Polygon
    Set pBGUnitsPlyLyr = Me.GeologyUnitsPolygonLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGUnitsPlyLyr = CreateQuadBGLayer(pBGUnitsPlyLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGUnitsPlyLyr = CreateQuadBGLayerByClip(pBGUnitsPlyLyr)
    End If
    Dim pQGrpLyr As IGeoFeatureLayer
    Set pQGrpLyr = pQuadBGUnitsPlyLyr
    Set pQGrpLyr.Renderer = BGUnitsPolyUVR(pQuadBGUnitsPlyLyr)
    Call LabelFeatures(pQuadBGUnitsPlyLyr)
    
    'Create the Quad BG Units Point
    Set pBGUnitsPntLyr = Me.GeologyUnitsPointLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGUnitsPntLyr = CreateQuadBGLayer(pBGUnitsPntLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGUnitsPntLyr = CreateQuadBGLayerByClip(pBGUnitsPntLyr)
    End If
    If (Not pQuadBGUnitsPntLyr Is Nothing) Then
        pQuadBGUnitsPntLyr.Visible = True
        pQuadBGUnitsPntLyr.Name = "Quad BG Units-Points"
        ODNR_Common.CollapseLegend pQuadBGUnitsPntLyr
        pQuadGrpLyr.Add pQuadBGUnitsPntLyr
    End If
    
    'Create the Quad BG Structure Point
    Set pBGStructurePntLyr = Me.BGStructurePointLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGStructurePntLyr = CreateQuadBGLayer(pBGStructurePntLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGStructurePntLyr = CreateQuadBGLayerByClip(pBGStructurePntLyr)
    End If
    If (Not pQuadBGStructurePntLyr Is Nothing) Then
        pQuadBGStructurePntLyr.Visible = True
        pQuadBGStructurePntLyr.Name = "Quad BG Structure-Points"
        ODNR_Common.CollapseLegend pQuadBGStructurePntLyr
        pQuadGrpLyr.Add pQuadBGStructurePntLyr
    End If
    
    'Create the Quad BG Mines Points
    Set pBGMinesPntLyr = Me.MinesPointLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGMinesPntLyr = CreateQuadBGLayer(pBGMinesPntLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGMinesPntLyr = CreateQuadBGLayerByClip(pBGMinesPntLyr)
    End If
    If (Not pQuadBGMinesPntLyr Is Nothing) Then
        pQuadBGMinesPntLyr.Visible = True
        pQuadBGMinesPntLyr.Name = "Quad BG Mines-Points"
        ODNR_Common.CollapseLegend pQuadBGMinesPntLyr
        pQuadGrpLyr.Add pQuadBGMinesPntLyr
    End If
    
    'Create the Quad BG Miscellaneous Point
    Set pBGMiscPntLyr = Me.BGMiscPntLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGMiscPntLyr = CreateQuadBGLayer(pBGMiscPntLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGMiscPntLyr = CreateQuadBGLayerByClip(pBGMiscPntLyr)
    End If
    If (Not pQuadBGMiscPntLyr Is Nothing) Then
        pQuadBGMiscPntLyr.Visible = True
        pQuadBGMiscPntLyr.Name = "Quad BG Miscellaneous Points"
        ODNR_Common.CollapseLegend pQuadBGMiscPntLyr
        pQuadGrpLyr.Add pQuadBGMiscPntLyr
    End If
    
    'Create the Quad BG Structure Line
    Set pBGStructureLineLyr = Me.BGStructureLineLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGStructureLineLyr = CreateQuadBGLayer(pBGStructureLineLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGStructureLineLyr = CreateQuadBGLayerByClip(pBGStructureLineLyr)
    End If
    If (Not pQuadBGStructureLineLyr Is Nothing) Then
        pQuadBGStructureLineLyr.Visible = True
        pQuadBGStructureLineLyr.Name = "Quad BG Structure-Lines"
        ODNR_Common.CollapseLegend pQuadBGStructureLineLyr
        pQuadGrpLyr.Add pQuadBGStructureLineLyr
    End If
    
    'Create the Quad BG Facies Line
    Set pBGFaciesLineLyr = Me.GeologyFaciesLineLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGFaciesLineLyr = CreateQuadBGLayer(pBGFaciesLineLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGFaciesLineLyr = CreateQuadBGLayerByClip(pBGFaciesLineLyr)
    End If
    If (Not pQuadBGFaciesLineLyr Is Nothing) Then
        pQuadBGFaciesLineLyr.Visible = True
        pQuadBGFaciesLineLyr.Name = "Quad BG Facies Lines"
        ODNR_Common.CollapseLegend pQuadBGFaciesLineLyr
        pQuadGrpLyr.Add pQuadBGFaciesLineLyr
    End If
    
    'Create the Quad BG Contacts Line
    Set pBGContactsLineLyr = Me.GeologyUnitsContactLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGContactsLineLyr = CreateQuadBGLayer(pBGContactsLineLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGContactsLineLyr = CreateQuadBGLayerByClip(pBGContactsLineLyr)
    End If
    If (Not pQuadBGContactsLineLyr Is Nothing) Then
        pQuadBGContactsLineLyr.Visible = True
        pQuadBGContactsLineLyr.Name = "Quad BG Contacts"
        ODNR_Common.CollapseLegend pQuadBGContactsLineLyr
        pQuadGrpLyr.Add pQuadBGContactsLineLyr
    End If
    
    'Create the Quad BG Units Line
    Set pBGUnitsLineLyr = Me.GeologyUnitsLineLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGUnitsLineLyr = CreateQuadBGLayer(pBGUnitsLineLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGUnitsLineLyr = CreateQuadBGLayerByClip(pBGUnitsLineLyr)
    End If
    If (Not pQuadBGUnitsLineLyr Is Nothing) Then
        pQuadBGUnitsLineLyr.Visible = True
        pQuadBGUnitsLineLyr.Name = "Quad BG Units Line"
        ODNR_Common.CollapseLegend pQuadBGUnitsLineLyr
        pQuadGrpLyr.Add pQuadBGUnitsLineLyr
    End If
   
    'Create the Quad BG Mines Line
    Set pBGMinesLineLyr = Me.MinesLineLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGMinesLineLyr = CreateQuadBGLayer(pBGMinesLineLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGMinesLineLyr = CreateQuadBGLayerByClip(pBGMinesLineLyr)
    End If
    If (Not pQuadBGMinesLineLyr Is Nothing) Then
        pQuadBGMinesLineLyr.Visible = True
        pQuadBGMinesLineLyr.Name = "Quad BG Mines-Lines"
        ODNR_Common.CollapseLegend pQuadBGMinesLineLyr
        pQuadGrpLyr.Add pQuadBGMinesLineLyr
    End If
    
    'Create the Quad BG Mines Polygon
    Set pBGMinesPlyLyr = Me.MinesPolygonLayer
    If (g_strClip = "PreCut") Then
        Set pQuadBGMinesPlyLyr = CreateQuadBGLayer(pBGMinesPlyLyr)
    ElseIf (g_strClip = "Cut on the fly") Then
        Set pQuadBGMinesPlyLyr = CreateQuadBGLayerByClip(pBGMinesPlyLyr)
    End If
    If (Not pQuadBGMinesPlyLyr Is Nothing) Then
        pQuadBGMinesPlyLyr.Visible = True
        pQuadBGMinesPlyLyr.Name = "Quad BG Mines-Polygons"
        ODNR_Common.CollapseLegend pQuadBGMinesPlyLyr
        pQuadGrpLyr.Add pQuadBGMinesPlyLyr
    End If
    
    'Add the Group Layer
    pMap.AddLayer pQuadGrpLyr
    pMap.MoveLayer pQuadGrpLyr, pMap.LayerCount - 3
        
    'Add the Quad BG Units Polygon
    Set pGFlyr = pQuadBGUnitsPlyLyr
    pGFlyr.Visible = True
    pGFlyr.Name = "Quad Geology Units Ply"
    Set pFLayer = pGFlyr
    ODNR_Common.CollapseLegend pFLayer
    pMap.AddLayer pGFlyr
    pMap.MoveLayer pGFlyr, pMap.LayerCount - 1
    
    'Update the Table of Contents and the Active View
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadBGLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function BGUnitsPolyUVR(pQuadBGUnitsPly As ILayer) As IUniqueValueRenderer
'******************************************************************************
'
'   Program:    BGUnitsPolyUVR
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       November 30, 2005
'   Purpose:    Unique value renderer for the BG Units Polygons
'   Called from: ODNRProject.LoadQuadBGLayers
'
'******************************************************************************
    On Error GoTo ErrorHandler

    Dim pNewFlyr As IFeatureLayer, pGFlyr As IGeoFeatureLayer
    Dim pLayer As ILayer, pFLayer As IFeatureLayer
    Dim pUniqueValueRend As IUniqueValueRenderer, pFlyrDef As IFeatureLayerDefinition
    Dim strDefQry As String, pDs As IDataset, pFWs As IFeatureWorkspace
    Dim pQDef As IQueryDef, pRow As IRow, pCursor As ICursor, pNewUniqueValueRend As IUniqueValueRenderer
    Dim pUniqueUnitCodes As Dictionary, pCodeOrder As Dictionary, vKey As Variant
    Dim pTblSort As ITableSort, pQf As IQueryFilter
    Dim lngCSDAGEFieldIdx As Long, lngSTRATCODEFieldIdx As Long
    Dim lngNAMEFieldIdx As Long, lngUNITCODEFieldIdx As Long
    Dim pSym As ISymbol, pRendererPropPage As IRendererPropertyPage
    Dim pDataLayer As IDataLayer, pDsName As IDatasetName
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    Dim strFldUnitCode As String
    Dim strFldArea As String
    Dim pFc As IFeatureClass
    Dim pDs1 As IDataset
    Dim pField As IField
    Dim strFldName As String
    Dim intFieldIdx As Integer
    Dim pFeature As IFeature
    Dim pFCursor As IFeatureCursor
    Dim pQF1 As IQueryFilter
    
    Set pGFlyr = pQuadBGUnitsPly
    Set pUniqueValueRend = pGFlyr.Renderer
    
    Set pFlyrDef = pQuadBGUnitsPly
    Set pFLayer = pQuadBGUnitsPly
    Set pFc = pFLayer.FeatureClass
    Set pDs1 = pFc
    
    intFieldIdx = pFc.Fields.FindField("UNIT_CODE")
    
    Set pDs = pGFlyr.FeatureClass
    Set pFWs = pDs.Workspace
    Set pUniqueUnitCodes = New Dictionary
    
    strDefQry = pFlyrDef.DefinitionExpression
    If (strDefQry <> "") Then
        strDefQry = strDefQry & " AND "
    End If
    If (pDs1.Workspace.Type = esriLocalDatabaseWorkspace) Then
        strDefQry = strDefQry & "([SHAPE_AREA] >= 2500)"
    ElseIf (pDs1.Workspace.Type = esriRemoteDatabaseWorkspace) Then
        strDefQry = strDefQry & "(shape.area >= 2500)"
    Else
        strDefQry = strDefQry & "(""SHAPE_AREA"" >= 2500)"
    End If
    
    Set pQF1 = New QueryFilter
    pQF1.WhereClause = strDefQry

    Set pFCursor = pFc.Search(pQF1, True)
    Set pFeature = pFCursor.NextFeature
    
    strDefQry = ""
    Do While Not pFeature Is Nothing
        If (IsNull(pFeature.Value(intFieldIdx)) = False) Then
            If (Trim(pFeature.Value(intFieldIdx)) <> "") Then
                If (pUniqueUnitCodes.Exists(pFeature.Value(intFieldIdx)) = False) Then
                    If (strDefQry = "") Then
                        If (pDs.Workspace.Type = esriLocalDatabaseWorkspace) Then
                            strDefQry = "[UNIT_CODE] in ('" & pFeature.Value(intFieldIdx) & "'"
                        ElseIf (pDs.Workspace.Type = esriRemoteDatabaseWorkspace) Then
                            strDefQry = "UNIT_CODE in ('" & pFeature.Value(intFieldIdx) & "'"
                        Else
                            strDefQry = """UNIT_CODE"" in ('" & pFeature.Value(intFieldIdx) & "'"
                        End If
                    Else
                        strDefQry = strDefQry & ",'" & pFeature.Value(intFieldIdx) & "'"
                    End If
                    pUniqueUnitCodes.Add Key:=pFeature.Value(intFieldIdx), Item:=""
                End If
            End If
        End If
        Set pFeature = pFCursor.NextFeature
    Loop
    strDefQry = strDefQry & ")"
    
'    strDefQry = ""
'
'    Do While Not pFeature Is Nothing
'        If (IsNull(pFeature.Value(intFieldIdx)) = False) Then
'            If (Trim(pFeature.Value(intFieldIdx)) <> "") Then
'                If (pUniqueUnitCodes.Exists(pFeature.Value(intFieldIdx)) = False) Then
'                    If (strDefQry = "") Then
'                        If (pDs.Workspace.Type = esriLocalDatabaseWorkspace) Then
'                            strDefQry = "[UNIT_CODE] in ('" & pFeature.Value(intFieldIdx) & "'"
'                        ElseIf (pDs.Workspace.Type = esriRemoteDatabaseWorkspace) Then
'                            strDefQry = "UNIT_CODE in ('" & pFeature.Value(intFieldIdx) & "'"
'                        Else
'                            strDefQry = """UNIT_CODE"" in ('" & pFeature.Value(intFieldIdx) & "'"
'                        End If
'                    Else
'                        strDefQry = strDefQry & ",'" & pFeature.Value(intFieldIdx) & "'"
'                    End If
'                    pUniqueUnitCodes.Add Key:=pFeature.Value(intFieldIdx), Item:=""
'                End If
'            End If
'        End If
'        Set pFeature = pFCursor.NextFeature
'    Loop
'    strDefQry = strDefQry & ")"
    
    lngCSDAGEFieldIdx = pGFlyr.FeatureClass.FindField("CSD_AGE")
    lngSTRATCODEFieldIdx = pGFlyr.FeatureClass.FindField("STRATCODE")
    lngNAMEFieldIdx = pGFlyr.FeatureClass.FindField("FMNAME")
    lngUNITCODEFieldIdx = pGFlyr.FeatureClass.FindField("UNIT_CODE")
    If (lngCSDAGEFieldIdx > 0) And (lngSTRATCODEFieldIdx > 0) And (lngNAMEFieldIdx > 0) Then
        Set pQf = New QueryFilter
        pQf.WhereClause = strDefQry
        Set pTblSort = New TableSort
        With pTblSort
            .Fields = "CSD_AGE, STRATCODE"
            .Ascending("CSD_AGE") = True
            .Ascending("STRATCODE") = True
            .CaseSensitive("STRATCODE") = False
            Set .QueryFilter = pQf
            Set .Table = pGFlyr.FeatureClass
        End With
        pTblSort.Sort Nothing
        Set pCursor = pTblSort.Rows
        Set pRow = pCursor.NextRow
        Set pCodeOrder = New Dictionary
        Do While Not pRow Is Nothing
            If (pUniqueUnitCodes.Exists(pRow.Value(lngUNITCODEFieldIdx))) Then
                If (pCodeOrder.Exists(pRow.Value(lngUNITCODEFieldIdx)) = False) Then
                    pCodeOrder.Add Key:=pRow.Value(lngUNITCODEFieldIdx), Item:=pRow.Value(lngNAMEFieldIdx)
                End If
            End If
            Set pRow = pCursor.NextRow
        Loop
    Else
        Set pCodeOrder = New Dictionary
        For Each vKey In pUniqueUnitCodes.Keys
            pCodeOrder.Add Key:=vKey, Item:=vKey
        Next
    End If

    Set pNewUniqueValueRend = New UniqueValueRenderer
    pNewUniqueValueRend.FieldCount = 1
    pNewUniqueValueRend.Field(0) = "UNIT_CODE"
    pNewUniqueValueRend.FieldType(0) = True
    For Each vKey In pCodeOrder.Keys
        Set pSym = pUniqueValueRend.Symbol(CStr(vKey))
        pNewUniqueValueRend.AddValue CStr(vKey), pUniqueValueRend.Heading(CStr(vKey)), pSym
        pNewUniqueValueRend.Description(CStr(vKey)) = pCodeOrder.Item(vKey)
    Next
    Set BGUnitsPolyUVR = pNewUniqueValueRend
    Set pRendererPropPage = New UniqueValuePropertyPage
    pGFlyr.RendererPropertyPageClassID = pRendererPropPage.ClassID
    
    ODNR_Common.ShowMessage "", False
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".BGUnitsPolyUVR " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function CreateQuadBGLayer(pQuadBGLayer As ILayer) As ILayer
'******************************************************************************
'
'   Program:    CreateQuadBGLayer
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       November 30, 2005
'   Purpose:    Creates an object representing the individual quadrangle BG layers.
'   Called from: ODNRProject.LoadQuadBGLayers
'
'******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim pNewFlyr As IFeatureLayer, pGFlyr As IGeoFeatureLayer, pNewGFlyr As IGeoFeatureLayer
    Dim pFRenderer As IFeatureRenderer, pFlyrDef As IFeatureLayerDefinition
    Dim strDefQry As String, pDs As IDataset, pFWs As IFeatureWorkspace
    Dim pQDef As IQueryDef, pRow As IRow, pCursor As ICursor, pNewFRenderer As IFeatureRenderer
    Dim pUniqueUnitCodes As Dictionary, pCodeOrder As Dictionary, vKey As Variant
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    Dim pAnnoProps As IAnnotateLayerPropertiesCollection
    Dim pAnnoLabelProps As IAnnotateLayerProperties
    Dim pLabelEngine As ILabelEngineLayerProperties
    Dim blnDispAnno As Boolean
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)
    
    'This section gets the 24K or 100K Quad attribute field name
    Set pStateLayer = m_pStateLayers.GetLayerByName(pQuadBGLayer.Name)
    If Not (pStateLayer Is Nothing) Then
        If (Me.QuadScale = odnr24K) Then
            strField = pStateLayer.Quad24KFieldName
        ElseIf (Me.QuadScale = odnr100K) Then
            strField = pStateLayer.Quad100KFieldName
        End If
    Else
        Exit Function
    End If
    
    'As long as number of quads selected is >0
    If (m_pQuads.QuadCount > 0) Then
        'If the focus quad is not selected, then use the first quad in the list.
        'This is used for the floating selection box in the 100K map templates.
        If (m_pQuads.FocusQuad Is Nothing) Then
            m_pQuads.Reset
            Set pQuad = m_pQuads.NextQuad
            Do While Not pQuad Is Nothing
                If (strDefQry = "") Then
                    strDefQry = strField & " = '" & pQuad.QuadName & "'"
                Else
                    strDefQry = strDefQry & " Or " & strField & " = '" & pQuad.QuadName & "'"
                End If
                Set pQuad = m_pQuads.NextQuad
            Loop
        'Use the focus quad.
        Else
            Set pQuad = m_pQuads.FocusQuad
            strDefQry = strField & " = '" & pQuad.QuadName & "'"
        End If
        
        Set pNewFlyr = New FeatureLayer
        Set pGFlyr = pQuadBGLayer
        Set pFRenderer = pGFlyr.Renderer
        Set pAnnoProps = pGFlyr.AnnotationProperties
        pAnnoProps.QueryItem 0, pAnnoLabelProps
        Set pLabelEngine = pAnnoLabelProps
        blnDispAnno = pGFlyr.DisplayAnnotation
        Set pNewFlyr.FeatureClass = pGFlyr.FeatureClass
        Set pFlyrDef = pNewFlyr
        pFlyrDef.DefinitionExpression = strDefQry
        
        Set pNewGFlyr = pNewFlyr
        Set pNewGFlyr.Renderer = pFRenderer
        
        pAnnoProps.Clear
        pAnnoProps.Add pAnnoLabelProps
        pNewGFlyr.AnnotationProperties = pAnnoProps
        pNewGFlyr.DisplayAnnotation = blnDispAnno
        
        Dim pFCCount As Long
        Dim pFCCQryFil As IQueryFilter
        Set pFCCQryFil = New QueryFilter
        pFCCQryFil.WhereClause = strDefQry
        
        pFCCount = pNewFlyr.FeatureClass.FeatureCount(pFCCQryFil)
        If (pFCCount <= 0) Then
            Set CreateQuadBGLayer = Nothing
            Exit Function
        Else
            Set CreateQuadBGLayer = pNewFlyr
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".CreateQuadBGLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub UnloadGeologyLayer()
    On Error GoTo ErrorHandler

    Dim pMapSurround As IMapSurround, pLegend As ILegend
    Dim pLegendItem As ILegendItem, lngMapSIdx As Long
    Dim pMxDoc As IMxDocument, pMap As IMap, pLayer As ILayer
    Dim pStateLayer As ODNRStateLayer
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)
    
    Set pStateLayer = m_pStateLayers.GetLayerByName(Me.GeologyUnitsPolygonLayer.Name)
    If Not (pStateLayer Is Nothing) Then
        For lngMapSIdx = 0 To pMap.MapSurroundCount - 1
            Set pMapSurround = pMap.MapSurround(lngMapSIdx)
            If (TypeOf pMapSurround Is ILegend) Then
                Set pLegend = pMapSurround
                Exit For
            End If
        Next
        If Not (pLegend Is Nothing) Then
            If (pLegend.ItemCount > 0) Then
                Set pLegendItem = pLegend.Item(0)
                Set pLayer = pStateLayer.ESRILayer
                Set pLegendItem.Layer = pLayer
                pLegend.Refresh
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UnloadGeologyLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function FieldValue(pRow As IRow, lngFieldIdx As Long) As Variant
    On Error GoTo ErrorHandler
        
        If Not (pRow Is Nothing) And (lngFieldIdx > -1) Then
            If (IsNull(pRow.Value(lngFieldIdx))) Then
                FieldValue = ""
            Else
                FieldValue = pRow.Value(lngFieldIdx)
            End If
        End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".FieldValue " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function ParseString(intBreak As Integer, strText As String) As String
    On Error GoTo ErrorHandler

    If (Len(strText) > intBreak) Then
        Dim pList As Collection, vKey As Variant, intSplit As Integer
        Dim blnDone As Boolean, strSplit As String
        Dim strReturn As String
        Set pList = New Collection
        blnDone = False
        Do While blnDone = False
            intSplit = InStr(intBreak, strText, " ")
            If (intSplit = 0) Then
                pList.Add Item:=strText
                blnDone = True
            Else
                strSplit = Left(strText, intSplit)
                pList.Add Item:=Trim(strSplit)
                strText = Right(strText, Len(strText) - intSplit)
            End If
        Loop
        strReturn = ""
        For Each vKey In pList
            If (strReturn = "") Then
                strReturn = CStr(vKey)
            Else
                strReturn = strReturn & vbCrLf & CStr(vKey)
            End If
        Next
        ParseString = strReturn
    Else
        ParseString = strText
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ParseString " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function EditingBGUnitsPolyUVR(pQuadBGUnitsPly As ILayer) As IUniqueValueRenderer
'******************************************************************************
'
'   Program:    EditingBGUnitsPolyUVR
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       November 30, 2005
'   Purpose:    Unique value renderer for the BG Units Polygons for Editing
'   Called from: ODNRProject.LoadQuadBGLayers
'
'******************************************************************************
'NOTE: Error Handler was removed since an unspecified error was occuring.  It currently
'does not affect operations.  It might do so in the future.  The error was posted to the
'ESRI EDN Forums 20060105.
    On Error GoTo ErrorHandler

    Dim pNewFlyr As IFeatureLayer, pGFlyr As IGeoFeatureLayer
    Dim pLayer As ILayer, pFLayer As IFeatureLayer
    Dim pUniqueValueRend As IUniqueValueRenderer, pFlyrDef As IFeatureLayerDefinition
    Dim strDefQry As String, pDs As IDataset, pFWs As IFeatureWorkspace
    Dim pQDef As IQueryDef, pRow As IRow, pCursor As ICursor, pNewUniqueValueRend As IUniqueValueRenderer
    Dim pUniqueUnitCodes As Dictionary, pCodeOrder As Dictionary, vKey As Variant
    Dim pTblSort As ITableSort, pQf As IQueryFilter
    Dim lngCSDAGEFieldIdx As Long, lngSTRATCODEFieldIdx As Long
    Dim lngNAMEFieldIdx As Long, lngUNITCODEFieldIdx As Long
    Dim lngUnitCodeIdx As Long
    Dim pSym As ISymbol, pRendererPropPage As IRendererPropertyPage
    Dim pDataLayer As IDataLayer, pDsName As IDatasetName
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    Dim strFldUnitCode As String
    Dim strFldArea As String
    Dim pFc As IFeatureClass
    
    Dim vArrayItems As Variant
    Dim vArrayKeys As Variant
        
    Set pGFlyr = pQuadBGUnitsPly
    Set pUniqueValueRend = pGFlyr.Renderer
    
    Set pFLayer = pQuadBGUnitsPly
    Set pFc = pFLayer.FeatureClass
    
    Set pDs = pGFlyr.FeatureClass
    Set pFWs = pDs.Workspace
    
    'Add section on creating a spatial query filter and feature cursor
    'Create Spatial Query Filter
    Dim pSQf As ISpatialFilter
    Dim pGeom As IGeometry
    Set pSQf = New SpatialFilter
    Set pGeom = m_pQuads.FocusQuad.QuadBoundary
    Set pSQf.Geometry = pGeom
    pSQf.SpatialRel = esriSpatialRelIntersects
    
    'Create feature cursor
    Dim pFCur As IFeatureCursor
    Set pFCur = pFc.Search(pSQf, True)
    
    Dim pFeature As IFeature
    lngUnitCodeIdx = pGFlyr.FeatureClass.FindField("UNIT_CODE")
    
    Set pFeature = pFCur.NextFeature
    Set pUniqueUnitCodes = New Dictionary
    strDefQry = ""
    Do While Not pFeature Is Nothing
        If (IsNull(pFeature.Value(lngUnitCodeIdx)) = False) Then
            If (Trim(pFeature.Value(lngUnitCodeIdx)) <> "") Then
                If (pUniqueUnitCodes.Exists(pFeature.Value(lngUnitCodeIdx)) = False) Then
                    If (strDefQry = "") Then
                        If (pDs.Workspace.Type = esriLocalDatabaseWorkspace) Then
                            strDefQry = "[UNIT_CODE] in ('" & pFeature.Value(lngUnitCodeIdx) & "'"
                        ElseIf (pDs.Workspace.Type = esriRemoteDatabaseWorkspace) Then
                            strDefQry = "UNIT_CODE in ('" & pFeature.Value(lngUnitCodeIdx) & "'"
                        End If
                    Else
                        strDefQry = strDefQry & ",'" & pFeature.Value(lngUnitCodeIdx) & "'"
                    End If
                    pUniqueUnitCodes.Add Key:=pFeature.Value(lngUnitCodeIdx), Item:=""
                End If
            End If
        End If
        Set pFeature = pFCur.NextFeature
    Loop
    strDefQry = strDefQry & ")"
    
    'This takes the original feature class, creates a new WHERE clause query
    'definition, and then does a table sort using just the features found in
    'using the previous spatial query definition.
    lngCSDAGEFieldIdx = pGFlyr.FeatureClass.FindField("CSD_AGE")
    lngSTRATCODEFieldIdx = pGFlyr.FeatureClass.FindField("STRATCODE")
    lngNAMEFieldIdx = pGFlyr.FeatureClass.FindField("FMNAME")
    lngUNITCODEFieldIdx = pGFlyr.FeatureClass.FindField("UNIT_CODE")
    If (lngCSDAGEFieldIdx > 0) And (lngSTRATCODEFieldIdx > 0) And (lngNAMEFieldIdx > 0) Then
        Set pQf = New QueryFilter
        pQf.WhereClause = strDefQry
        Set pTblSort = New TableSort
        With pTblSort
            .Fields = "CSD_AGE, STRATCODE"
            .Ascending("CSD_AGE") = True
            .Ascending("STRATCODE") = True
            .CaseSensitive("STRATCODE") = False
            Set .QueryFilter = pQf
            Set .Table = pGFlyr.FeatureClass
        End With
        pTblSort.Sort Nothing
        
        'Once the table is sorted, populate pUniqueUnitCodes dictionary
        Set pCursor = pTblSort.Rows
        Set pRow = pCursor.NextRow
        Set pCodeOrder = New Dictionary
        Do While Not pRow Is Nothing
            If (pUniqueUnitCodes.Exists(pRow.Value(lngUNITCODEFieldIdx))) Then
                If (pCodeOrder.Exists(pRow.Value(lngUNITCODEFieldIdx)) = False) Then
                    pCodeOrder.Add Key:=pRow.Value(lngUNITCODEFieldIdx), Item:=pRow.Value(lngNAMEFieldIdx)
                End If
            End If
            Set pRow = pCursor.NextRow
        Loop
    Else
        Set pCodeOrder = New Dictionary
        For Each vKey In pUniqueUnitCodes.Keys
            pCodeOrder.Add Key:=vKey, Item:=vKey
        Next
    End If
    
    Set pNewUniqueValueRend = New UniqueValueRenderer
    pNewUniqueValueRend.FieldCount = 1
    pNewUniqueValueRend.Field(0) = "UNIT_CODE"
    pNewUniqueValueRend.FieldType(0) = True
    For Each vKey In pCodeOrder.Keys
        Set pSym = pUniqueValueRend.Symbol(CStr(vKey))
        pNewUniqueValueRend.AddValue CStr(vKey), pUniqueValueRend.Heading(CStr(vKey)), pSym
        pNewUniqueValueRend.Description(CStr(vKey)) = pCodeOrder.Item(vKey)
    Next
    Set EditingBGUnitsPolyUVR = pNewUniqueValueRend
    
    ODNR_Common.ShowMessage "", False
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".EditingBGUnitsPolyUVR " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function CreateQuadBGLayerByClip(pOverlayLayer As ILayer) As ILayer
'******************************************************************************
'
'   Program:    CreateQuadBGLayerByClip
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       January 23, 2006
'   Purpose:    Creates shape files from the BG feature classes, by clipping
'               against the quadrangle boundaries.
'   Called from: ODNRProject.LoadQuadBGLayers
'
'******************************************************************************

    On Error GoTo ErrorHandler
    
    Dim pIntersectFClass As IFeatureClass
    Dim pFilter As IQueryFilter
    Dim pIntersectFCursor As IFeatureCursor
    Dim pIntersectFeature As IFeature
    Dim pIntersectTopo As ITopologicalOperator
    Dim pTopoOp2 As ITopologicalOperator2
    Dim pInterTopoOp2 As ITopologicalOperator2
    Dim pOverlayTopoOp As ITopologicalOperator
    Dim pPolygon As IPolygon
    Dim pSpatialFilter As ISpatialFilter
    Dim pIntersectFields As IFields
    Dim pOverlayFCursor As IFeatureCursor
    Dim pOverlayFClass As IFeatureClass
    Dim pOverlayFields As IFields
    Dim pOverlayFeature As IFeature
    Dim pOutputfClass As IFeatureClass
    Dim newGeometry As IGeometry
    Dim newGeomFix As IGeometry
    Dim pInterGeomFix As IGeometry
    Dim pOverlayIndex As Long
    Dim pOldFieldsCount As Long
    Dim Item As Long
    Dim pOldField As IField
    Dim pNewFeatureCursor As IFeatureCursor, pNewBuffer As IFeatureBuffer
    Dim intPolyCount As Integer
    Dim intNewBufferIdx As Integer
    Dim pRing() As IRing
    Dim pIntRing() As IRing
    Dim pNewPolygon As IPolygon2
    Dim l As Long
    Dim pArea As IArea
    Dim pPolygon2 As IGeometryCollection
    Dim pPolyline As IGeometryCollection
    Dim pPolyline2 As IGeometryCollection
    Dim pPolyline3 As IPolyline
    Dim pOverlayFLayer As IFeatureLayer
    Dim pQuad As ODNRQuad
    Dim pStateLayer As ODNRStateLayer
    Dim strDefQry As String
    Dim pNewGFlyr As IGeoFeatureLayer
    Dim pGFlyr As IGeoFeatureLayer
    Dim strField As String
    Dim pAnnoProps As IAnnotateLayerPropertiesCollection
    Dim pAnnoLabelProps As IAnnotateLayerProperties
    Dim pLabelEngine As ILabelEngineLayerProperties 'Delete in the future
    Dim blnDispAnno As Boolean
    Dim pFRenderer As IFeatureRenderer
    Dim n As Integer
    Dim intOverFeatCount As Integer
    Dim pNewlayer As IFeatureLayer

    Set pOverlayFLayer = pOverlayLayer
    Set pIntersectFClass = QuadFeatureLayer.FeatureClass 'Quadrangle feature class
    Set pOverlayFClass = pOverlayFLayer.FeatureClass 'Geology feature classes
    Set pOutputfClass = NewShapeFile(pOverlayFClass) 'New Geology Shape file feature classes
  
    'Maintain the original fields collection
    Set pOverlayFields = pOverlayFClass.Fields
    pOldFieldsCount = pOverlayFClass.Fields.FieldCount
    
    Set pStateLayer = m_pStateLayers.GetLayerByName(QuadFeatureLayer.Name)
    If Not (pStateLayer Is Nothing) Then
        If (Me.QuadScale = odnr24K) Then
            strField = pStateLayer.Quad24KFieldName
        ElseIf (Me.QuadScale = odnr100K) Then
            strField = pStateLayer.Quad100KFieldName
        End If
    Else
        Exit Function
    End If
    
    If (m_pQuads.QuadCount > 0) Then
        
        'If the focus quad is not selected, then use the first quad in the list.
        If (m_pQuads.FocusQuad Is Nothing) Then
            m_pQuads.Reset
            Set pQuad = m_pQuads.NextQuad
            Do While Not pQuad Is Nothing
                If (strDefQry = "") Then
                    strDefQry = strField & " = '" & pQuad.QuadName & "'"
                Else
                    strDefQry = strDefQry & " Or " & strField & " = '" & pQuad.QuadName & "'"
                End If
                Set pQuad = m_pQuads.NextQuad
            Loop
        'Use the focus quad.
        Else
            Set pQuad = m_pQuads.FocusQuad
            strDefQry = strField & " = '" & pQuad.QuadName & "'"
        End If
        
        ' get the buffer record for the Quad layer
        Set pFilter = New QueryFilter
        pFilter.WhereClause = strDefQry
        
        Set pIntersectFCursor = QuadFeatureLayer.Search(pFilter, False)
        
        'Create an insert cursor to hold the new features
        Set pNewFeatureCursor = pOutputfClass.Insert(True)
        
        'Start with buffer feature for the Quad layer
        Set pIntersectFeature = pIntersectFCursor.NextFeature
        While Not pIntersectFeature Is Nothing
            Set pInterTopoOp2 = pIntersectFeature.ShapeCopy
            pInterTopoOp2.IsKnownSimple = False
            pInterTopoOp2.Simplify
            Set pIntersectTopo = pInterTopoOp2
            Set pInterGeomFix = pIntersectTopo
            
            'get the intersecting layer features - Quad Layer
            Set pSpatialFilter = New SpatialFilter
            pSpatialFilter.GeometryField = pIntersectFClass.shapeFieldName
'            Set pSpatialFilter.Geometry = pIntersectFeature.Shape
            Set pSpatialFilter.Geometry = pInterGeomFix
            pSpatialFilter.SpatialRel = esriSpatialRelIntersects
            Set pOverlayFCursor = pOverlayFClass.Search(pSpatialFilter, False) 'Find all the Geology features that intersect the quad layer
            'Cycle through the features intersecting that buffer poly
            Set pOverlayFeature = pOverlayFCursor.NextFeature
            While Not pOverlayFeature Is Nothing
              
                If (pOutputfClass.ShapeType = esriGeometryPolygon) Then
                    Set pTopoOp2 = pOverlayFeature.ShapeCopy
                    pTopoOp2.IsKnownSimple = False
                    pTopoOp2.Simplify
                    
                    Set newGeomFix = pTopoOp2
                    Set newGeometry = pIntersectTopo.Intersect(newGeomFix, esriGeometry2Dimension)
                    
                    Set pOverlayTopoOp = newGeometry
                    
                    If (newGeometry.IsEmpty = False) Then
                        Set pPolygon = newGeometry
                        ReDim pRing(pPolygon.ExteriorRingCount - 1)
                        
                        Set pNewPolygon = pPolygon
                        pNewPolygon.QueryExteriorRingsEx pPolygon.ExteriorRingCount, pRing(0)
                        For l = 0 To pPolygon.ExteriorRingCount - 1
                            Set pPolygon2 = New Polygon
                            pPolygon2.AddGeometry pRing(l)
                            
                            If (pPolygon.InteriorRingCount(pRing(l)) > 0) Then
                                ReDim pIntRing(pPolygon.InteriorRingCount(pRing(l)) - 1)
                                pNewPolygon.QueryInteriorRingsEx pRing(l), pPolygon.InteriorRingCount(pRing(l)), pIntRing(0)
                                
                                For n = 0 To pPolygon.InteriorRingCount(pRing(l)) - 1
                                    pPolygon2.AddGeometry pIntRing(n)
                                Next
                            End If
                            
                            Set pArea = pPolygon2
                            'Calculate attributes in the new feature based on the overlay layer - The Quad layer
                            Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
                            Set pNewBuffer.Shape = pPolygon2
        
                            'Calculate the original attributes based on the overlay - The Geology Layer
                            For Item = 0 To pOldFieldsCount - 1
                                Set pOldField = pOverlayFClass.Fields.Field(Item)
                                intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
                                If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                        And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                        And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
                                    pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
                                    If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                        Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
                                        If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                            pNewBuffer.Value(intNewBufferIdx) = 0
                                        Else
                                            pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                        End If
                                    Else
                                        If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                            pNewBuffer.Value(intNewBufferIdx) = ""
                                        Else
                                            pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                        End If
                                    End If
                                ElseIf UCase(pOldField.Name) = "SHAPE_AREA" Or UCase(pOldField.Name) = "SHAPE.AREA" Then
                                    intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_area")
                                    pNewBuffer.Value(intNewBufferIdx) = pArea.Area
                                ElseIf UCase(pOldField.Name) = "SHAPE_LENGTH" Then
                                    intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_leng")
                                    pNewBuffer.Value(intNewBufferIdx) = pPolygon.Length
                                ElseIf UCase(pOldField.Name) = "SHAPE.LEN" Then
                                    intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_len")
                                    pNewBuffer.Value(intNewBufferIdx) = pPolygon.Length
                                End If
                            Next Item
        
                            'Store the feature
                            pNewFeatureCursor.InsertFeature pNewBuffer
                        Next l
                    End If
                ElseIf (pOutputfClass.ShapeType = esriGeometryPolyline) Then
                    Set newGeometry = pIntersectTopo.Intersect(pOverlayFeature.Shape, pOverlayFeature.Shape.Dimension)
                    Set pPolyline = newGeometry
                    For l = 0 To pPolyline.GeometryCount - 1
                        Set pPolyline2 = New Polyline
                        pPolyline2.AddGeometry pPolyline.Geometry(l)
                        Set pPolyline3 = pPolyline2
                        
                        'Calculate attributes in the new feature based on the overlay layer - The Quad layer
                        Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
                        Set pNewBuffer.Shape = pPolyline2
    
                        'Calculate the original attributes based on the overlay - The Geology Layer
                        For Item = 0 To pOldFieldsCount - 1
                            Set pOldField = pOverlayFClass.Fields.Field(Item)
                            intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
                            If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                    And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                    And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
                                pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
                                If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                    Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
                                    If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                        pNewBuffer.Value(intNewBufferIdx) = 0
                                    Else
                                        pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                    End If
                                Else
                                    If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                        pNewBuffer.Value(intNewBufferIdx) = ""
                                    Else
                                        pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                    End If
                                End If
                            ElseIf UCase(pOldField.Name) = "SHAPE_LENGTH" Then
                                intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_leng")
                                pNewBuffer.Value(intNewBufferIdx) = pPolyline3.Length
                            ElseIf UCase(pOldField.Name) = "SHAPE.LEN" Then
                                intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_len")
                                pNewBuffer.Value(intNewBufferIdx) = pPolyline3.Length
                            End If
                        Next Item
    
                        'Store the feature
                        pNewFeatureCursor.InsertFeature pNewBuffer
                        Next l
                ElseIf (pOutputfClass.ShapeType = esriGeometryPoint) Then
                    Set newGeometry = pIntersectTopo.Intersect(pOverlayFeature.Shape, pOverlayFeature.Shape.Dimension)
                    Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
                    Set pNewBuffer.Shape = newGeometry
    
                    'Calculate the original attributes based on the overlay - The Geology Layer
                    For Item = 0 To pOldFieldsCount - 1
                        Set pOldField = pOverlayFClass.Fields.Field(Item)
                        intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
                        If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                    And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                    And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
                            pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
                            If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
                                If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                    pNewBuffer.Value(intNewBufferIdx) = 0
                                Else
                                    pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                End If
                            Else
                                If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
                                    pNewBuffer.Value(intNewBufferIdx) = ""
                                Else
                                    pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
                                End If
                            End If
                        End If
                    Next Item
    
                    'Store the feature
                    pNewFeatureCursor.InsertFeature pNewBuffer
                End If
                Set pOverlayFeature = pOverlayFCursor.NextFeature
            Wend
            Set pIntersectFeature = pIntersectFCursor.NextFeature
        Wend
    End If
    
    'Check to see if there are any features in the new feature class
    If (pOutputfClass.FeatureCount(Nothing) > 0) Then
        'Finish seting up the Rendering and Labeling
        Set pGFlyr = pOverlayLayer
        'Get the old Renderer from the original BG Layers
        Set pFRenderer = pGFlyr.Renderer
        'Get the Labeling, if it exists, from the original BG Layers
        Set pAnnoProps = pGFlyr.AnnotationProperties
        pAnnoProps.QueryItem 0, pAnnoLabelProps
        Set pLabelEngine = pAnnoLabelProps
        blnDispAnno = pGFlyr.DisplayAnnotation
        'Create a new Featurelayer
        Set pNewlayer = New FeatureLayer
        Set pNewlayer.FeatureClass = pOutputfClass
        'Apply the Renderer to the new Geofeaturelayer
        Set pNewGFlyr = pNewlayer
        Set pNewGFlyr.Renderer = pFRenderer
        'Apply the Labeling
        pAnnoProps.Clear
        pAnnoProps.Add pAnnoLabelProps
        pNewGFlyr.AnnotationProperties = pAnnoProps
        pNewGFlyr.DisplayAnnotation = blnDispAnno
        'Create new layer
        Set CreateQuadBGLayerByClip = pNewlayer
    ElseIf (pOutputfClass.FeatureCount(Nothing) <= 0) Then
        'If there are no features in the feature class, then set the layer to nothing.
        Set CreateQuadBGLayerByClip = Nothing
    End If
    
    Exit Function
  
ErrorHandler:
    HandleError True, c_strModuleName & ".CreateQuadBGLayerByClip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function NewShapeFile(pTemplateFC As IFeatureClass) As IFeatureClass
'******************************************************************************
'
'   Program:    NewShapeFile
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       January 23, 2006
'   Purpose:    Creates a blank shape files for the clipping against the
'               quadrangle boundaries.
'   Called from: ODNRProject.CreateQuadBGLayerByClip
'
'******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim pFcExists As IFeatureClass
    Dim pFcDs As IDataset
    Dim pWsDs As IDataset
    Dim pTemplateDS As IDataset
    Dim pWsInDs As IDataset
    Dim pWsExists As IWorkspace
    Dim pWsIn As IWorkspace
    Dim pFWsIn As IFeatureWorkspace
    Dim pFWs As IFeatureWorkspace
    Dim pNewFeatClass As IFeatureClass
    Dim strFcName As String
    Dim i As Integer
    Dim intLen As Integer
    Dim intLoc As Integer
    Dim strPos As String
    'Objects used for the metadata
    Dim pNameInput As IName
    Dim pNameOutput As IName
    Dim pDsInput As IDataset
    Dim pDsOutput As IDataset
    Dim pMetadataInput As IMetadata
    Dim pMetadataOutput As IMetadata
    Dim pPsInput As IPropertySet
    Dim pPsOutput As IPropertySet
    
    'Open the input feature class workspace
    Set pTemplateDS = pTemplateFC
    Set pWsIn = pTemplateDS.Workspace
    
    'Get the template feature class name. Check to see if the feature class is in SDE, PGDB, or a file based name.
    If (pWsIn.Type = esriLocalDatabaseWorkspace) Then
        strFcName = pTemplateDS.Name
    ElseIf (pWsIn.Type = esriRemoteDatabaseWorkspace) Then
        strFcName = pTemplateDS.Name
        intLen = Len(strFcName)
        For i = 1 To intLen
            strPos = Mid(strFcName, i, 1)
            If (strPos = ".") Then intLoc = i
        Next i
        strFcName = Right(strFcName, intLen - intLoc)
    End If
    
    Dim strFcFullName As String
    strFcFullName = g_strExport_Path & "\" & strFcName
    
    'Open the shape file workspace
    Dim pOutShpWspName As IWorkspaceName
    Set pOutShpWspName = New WorkspaceName
    pOutShpWspName.PathName = g_strExport_Path
    pOutShpWspName.WorkspaceFactoryProgID = "esriCore.ShapefileWorkspaceFactory.1"
    Dim pName As IName
    Set pName = pOutShpWspName
    Dim pShapeWorkspace As IWorkspace
    Set pShapeWorkspace = pName.Open
    
    'Check to see if shape files exists, if so then delete
    Dim pWsFExists As IWorkspaceFactory
    Dim pTruncPath As String
    If InStr(1, strFcFullName, ".shp") > 0 Then
      pTruncPath = Left(strFcFullName, InStr(1, strFcFullName, ".shp") - 1)
    Else
      pTruncPath = strFcFullName
    End If
      
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(pTruncPath & ".shp") Or fs.FileExists(pTruncPath & ".dbf") Or _
        fs.FileExists(pTruncPath & ".shx") Then
        Set pWsFExists = New ShapefileWorkspaceFactory
        Set pWsExists = pWsFExists.OpenFromFile(g_strExport_Path, 0)
        'Note: Remove EntryName function
        Set pFWs = pWsExists
        Set pFcExists = pFWs.OpenFeatureClass(strFcName)
        Set pFcDs = pFcExists
        pFcDs.Delete
    End If

    'Validate the fields
    Dim pOutputFields As IFields
    Dim pFieldChecker As IFieldChecker
    Set pFieldChecker = New FieldChecker
    Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
    
    'Clone the fields of the Template feature class.  Create new FID field for the new shape file.
    Dim pErrorEnum As IEnumFieldError
    Dim pNewFields As IFields, pField As IField
    Dim pNewFieldsEdit As IFieldsEdit
    Set pNewFields = New Fields
    Set pNewFieldsEdit = pNewFields
    pNewFieldsEdit.FieldCount = pTemplateFC.Fields.FieldCount
    
    Dim pClone As IClone
    Dim pCloneField As IField
    Dim pOIDFld As IField
    Dim pOIDFldEdit As IFieldEdit
    Dim j As Integer
    For j = 0 To pTemplateFC.Fields.FieldCount - 1
        If (pTemplateFC.Fields.Field(j).Type = esriFieldTypeOID) Then
            Set pOIDFld = New Field
            Set pOIDFldEdit = pOIDFld
            With pOIDFldEdit
                .Name = "FID"
                .Type = esriFieldTypeOID
            End With
            Set pNewFieldsEdit.Field(j) = pOIDFld
        Else
            Set pClone = pTemplateFC.Fields.Field(j)
            Set pCloneField = pClone.Clone
            Set pNewFieldsEdit.Field(j) = pCloneField
        End If
    Next
    
    pFieldChecker.Validate pNewFields, pErrorEnum, pOutputFields
    
    'Create the output featureclass
    Dim pUID As IUID
    Set pUID = pTemplateFC.clsid
    Dim shapeFieldName As String
    shapeFieldName = "Shape"
    
    Dim pFeatureWorkspace As IFeatureWorkspace
    Set pFeatureWorkspace = pShapeWorkspace
    Set pNewFeatClass = pFeatureWorkspace.CreateFeatureClass(strFcName, pOutputFields, pUID, Nothing, esriFTSimple, shapeFieldName, "")
    
    'Create Metadata for Output Feature Class
    'Setup metadata for input feature class
    Set pDsInput = pTemplateFC
    Set pNameInput = pDsInput.FullName
    Set pMetadataInput = pNameInput
    Set pPsInput = pMetadataInput.Metadata
    'Setup metadata for output feature class
    Set pDsOutput = pNewFeatClass
    Set pNameOutput = pDsOutput.FullName
    Set pMetadataOutput = pNameOutput
    pMetadataOutput.Synchronize esriMSAAlways, 0 'Create metadata XML file for output feature class
    pMetadataOutput.Metadata = pPsInput 'Copy input metadata records to output metadata records
    pMetadataOutput.Synchronize esriMSAAlways, 0 'Update the metadata XML file
    
    'Finish assigning variables and exit the function
    Set NewShapeFile = pNewFeatClass
    
    Exit Function
  
ErrorHandler:
    HandleError True, c_strModuleName & ".NewShapeFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub LabelFeatures(pFLayerPly As IFeatureLayer)
'------------------------------------------------------------------------------
'Title: LabelFeatures
'Date: 20060214
'Version: 1.1, 20060718
'Abstract:  This procedure labels the AUM, BG 24K, BG 100K polygons, and the
'           oil and gas well locations.
'------------------------------------------------------------------------------
'James McDonald
'GIMS Specialist
'Ohio Division of Geological Survey
'2045 Morse Road
'Columbus, OH  43229-6693
'Ph. (614) 265-6601
'E-mail: jim.mcdonald@dnr.state.oh.us
'------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    Dim pGeoFLayer As IGeoFeatureLayer
    Dim pAnnoProps As IAnnotateLayerPropertiesCollection
    Dim pAnnoLayerProps As IAnnotateLayerProperties
    Dim pLabelEngine As ILabelEngineLayerProperties
    Dim pBasicOpLyrProp As IBasicOverposterLayerProperties
    Dim pTextSym As ITextSymbol
    Dim pFTextSym As IFormattedTextSymbol
    Dim pFont As stdole.IFontDisp
    Dim pWs As IWorkspace
    Dim pFWs As IFeatureWorkspace
    Dim pFc As IFeatureClass
    Dim pDs As IDataset
    Dim pPntProps As IPointPlacementPriorities
    
    Set pGeoFLayer = pFLayerPly
    
    If (Me.ProjectType = odnrGeology) Then
        'Set the Font
        Set pFont = New stdole.StdFont
        pFont.Name = "Arial"
        pFont.Bold = True
        
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        pTextSym.Size = 9
        
        Set pAnnoProps = pGeoFLayer.AnnotationProperties
        pAnnoProps.Clear
        
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "[UNIT_CODE]"
        Set pLabelEngine.Symbol = pTextSym
        pLabelEngine.IsExpressionSimple = True
        
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp

        'Set the AnnotateLayerProperties and the apply to the layer
        Set pAnnoLayerProps = pLabelEngine
        Dim strPolySize As String
        
        'Set the area of the polygon to be labeled
        If (Me.QuadScale = odnr24K) Then
            strPolySize = "750000"
        ElseIf (Me.QuadScale = odnr100K) Then
            strPolySize = "30000000"
        End If
        
        'Create the WHERE clause for the IAnnotateLayerProperties interface
        Set pFc = pFLayerPly.FeatureClass
        Set pDs = pFc
        Set pWs = pDs.Workspace
        If (pWs.Type = esriRemoteDatabaseWorkspace) Then
            pAnnoLayerProps.WhereClause = "SHAPE.AREA >= " & strPolySize
        ElseIf (pWs.Type = esriLocalDatabaseWorkspace) Then
            pAnnoLayerProps.WhereClause = "[SHAPE_AREA] >= " & strPolySize
        ElseIf (pWs.Type = esriFileSystemWorkspace) Then
            pAnnoLayerProps.WhereClause = """SHAPE_AREA"" >= " & strPolySize
        End If
        pAnnoLayerProps.DisplayAnnotation = True

        pAnnoProps.Add pAnnoLayerProps
        pGeoFLayer.AnnotationProperties = pAnnoProps
        pGeoFLayer.DisplayAnnotation = True
    ElseIf (Me.ProjectType = odnrAUM) Then
        'Set the Font
        Set pFont = New stdole.StdFont
        pFont.Name = "Arial"
        pFont.Bold = True
        
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        
        Set pAnnoProps = pGeoFLayer.AnnotationProperties
        pAnnoProps.Clear
        
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "[MINE_CODE]"
        pLabelEngine.IsExpressionSimple = True
        
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp

        'Create the WHERE clause for the IAnnotateLayerProperties interface
        Dim strFieldArea As String
        Dim strWhere As String
        Set pFc = pFLayerPly.FeatureClass
        Set pDs = pFc
        Set pWs = pDs.Workspace
        If (pWs.Type = esriRemoteDatabaseWorkspace) Then
            strFieldArea = "SHAPE.AREA "
        ElseIf (pWs.Type = esriLocalDatabaseWorkspace) Then
            strFieldArea = "[SHAPE_AREA] "
        ElseIf (pWs.Type = esriFileSystemWorkspace) Then
            strFieldArea = """SHAPE_AREA"" "
        End If
        
        'Small labels
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        pTextSym.Size = 7
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "[MINE_CODE]"
        pLabelEngine.IsExpressionSimple = True
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        pBasicOpLyrProp.FeatureWeight = esriLowWeight
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp
        strWhere = strFieldArea & "<= 25000000"
        Set pLabelEngine.Symbol = pTextSym
        Set pAnnoLayerProps = pLabelEngine
        pAnnoLayerProps.DisplayAnnotation = True
        pAnnoLayerProps.WhereClause = strWhere
        pAnnoLayerProps.Class = "Small"
        pAnnoProps.Add pAnnoLayerProps
        
        'Medium labels
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        pTextSym.Size = 10
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "[MINE_CODE]"
        pLabelEngine.IsExpressionSimple = True
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        pBasicOpLyrProp.FeatureWeight = esriLowWeight
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp
        strWhere = strFieldArea & "> 25000001 AND " & strFieldArea & "<= 150000000"
        Set pLabelEngine.Symbol = pTextSym
        Set pAnnoLayerProps = pLabelEngine
        pAnnoLayerProps.DisplayAnnotation = True
        pAnnoLayerProps.WhereClause = strWhere
        pAnnoLayerProps.Class = "Medium"
        pAnnoProps.Add pAnnoLayerProps
        
        'Large labels
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Font = pFont
        pTextSym.Size = 13
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "[MINE_CODE]"
        pLabelEngine.IsExpressionSimple = True
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        pBasicOpLyrProp.FeatureWeight = esriLowWeight
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp
        strWhere = strFieldArea & "> 150000001"
        Set pLabelEngine.Symbol = pTextSym
        Set pAnnoLayerProps = pLabelEngine
        pAnnoLayerProps.DisplayAnnotation = True
        pAnnoLayerProps.WhereClause = strWhere
        pAnnoLayerProps.Class = "Large"
        pAnnoProps.Add pAnnoLayerProps
        
        pGeoFLayer.AnnotationProperties = pAnnoProps
        pGeoFLayer.DisplayAnnotation = True
    ElseIf (Me.ProjectType = odnrOGWells) Then
        'Set the Font
        Set pFont = New stdole.StdFont
        pFont.Name = "Arial"
        pFont.Bold = False
        pFont.Size = 6#
        
        'Create a new text symbol
        Set pTextSym = New TextSymbol
        pTextSym.Size = 6#
        pTextSym.Font = pFont
        
        Set pAnnoProps = pGeoFLayer.AnnotationProperties
        pAnnoProps.Clear
        
        'Create the label engine.
        Set pLabelEngine = New LabelEngineLayerProperties
        pLabelEngine.Expression = "Mid([API_NO],6,5)"
        pLabelEngine.IsExpressionSimple = True
        
        'Create a PointPlacementProperties object
        Set pPntProps = New PointPlacementPriorities
        pPntProps.AboveCenter = 1
        pPntProps.AboveLeft = 2
        pPntProps.AboveRight = 2
        pPntProps.BelowCenter = 3
        pPntProps.BelowLeft = 2
        pPntProps.BelowRight = 2
        pPntProps.CenterLeft = 2
        pPntProps.CenterRight = 2
        
        'Create a BasicOverposterLayerProperties object
        Set pBasicOpLyrProp = New BasicOverposterLayerProperties
        pBasicOpLyrProp.BufferRatio = 0
        pBasicOpLyrProp.FeatureType = esriOverposterPoint
        pBasicOpLyrProp.FeatureWeight = esriHighWeight
        pBasicOpLyrProp.LabelWeight = esriHighWeight
        pBasicOpLyrProp.NumLabelsOption = esriOneLabelPerShape
        pBasicOpLyrProp.PointPlacementMethod = esriAroundPoint
        pBasicOpLyrProp.PointPlacementOnTop = False
        pBasicOpLyrProp.PointPlacementPriorities = pPntProps
        
        'Create a OverposterLayerProperties2 object
        Dim pOLP2 As IOverposterLayerProperties2
        Set pOLP2 = pBasicOpLyrProp
        pOLP2.TagUnplaced = True
        
        'Finish creating the label engine
        Set pLabelEngine.BasicOverposterLayerProperties = pBasicOpLyrProp
        Set pLabelEngine.Symbol = pTextSym
        Set pAnnoLayerProps = pLabelEngine
        pAnnoLayerProps.DisplayAnnotation = True
        pAnnoProps.Add pAnnoLayerProps
        
        pGeoFLayer.AnnotationProperties = pAnnoProps
        pGeoFLayer.DisplayAnnotation = True
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LabelFeatures " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Public Property Get OGWellsLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrOGWells) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("OGWELLS")
            Set OGWellsLayer = pStateLayer.ESRILayer
        Else
            Set OGWellsLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".OGWellsLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get DirDrilledWellboreLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrOGWells) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Directional drilled wellbore")
            Set DirDrilledWellboreLayer = pStateLayer.ESRILayer
        Else
            Set DirDrilledWellboreLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".OGWellsLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get OLSLayer() As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pStateLayer As ODNRStateLayer
        m_pStateLayers.ActiveMap = odnrGeologyMap
        If (Me.ProjectType = odnrOGWells) Then
            Set pStateLayer = m_pStateLayers.GetLayerByName("Original_Land_Subdivisions")
            Set OLSLayer = pStateLayer.ESRILayer
        Else
            Set OLSLayer = Nothing
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".OLSLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function ChangeMapFrameGeometry(pMapFrame As IMapFrame) As IGeometry
    On Error GoTo ErrorHandler

    Dim pElem As IElement
    Dim pFGeom As IGeometry
    Dim pMGeom As IGeometry
    Dim pFPoint As IPoint
    Dim pMPoint As IPoint
    Dim pFPoly As IPolygon2
    Dim pMPoly As IPolygon2
    Dim pFArea As IArea
    Dim pMArea As IArea
    Dim pFRing() As IRing
    Dim pMRing() As IRing
    Dim pTRing As IRing
    Dim pFCentPnt As IPoint
    Dim pMCentPnt As IPoint
    Dim dblFCentroidX As Double
    Dim dblFCentroidY As Double
    Dim dblMCentroidX As Double
    Dim dblMCentroidY As Double
    Dim pFPntColl As IPointCollection
    Dim pMPntColl As IPointCollection
    Dim pEnumVert As IEnumVertex
    Dim dblXvert As Double
    Dim dblYvert As Double
    Dim intFPntCnt As Integer
    Dim dblScale As Double
    Dim lOutPart As Long, lOutVertex As Long
    Dim pNewPoint As IPoint
    Dim pMNewPoly As IPointCollection
    
    'Delete this after debug
    Dim pFSO As FileSystemObject
    Dim pTxtStrm As TextStream
    Set pFSO = New FileSystemObject
    Set pTxtStrm = pFSO.CreateTextFile("D:\shp_export\ChangeMapFrameGeometry.txt", True, False)
    pTxtStrm.WriteLine "ChangeMapFrameGeometry Coordinates: " & Date
            
    'Get the centroid of the feature class and the map frame
    Set pFGeom = m_pQuads.FocusQuad.QuadBoundary
    Set pFArea = pFGeom
    Set pFCentPnt = pFArea.Centroid
    dblFCentroidX = pFCentPnt.X
    dblFCentroidY = pFCentPnt.Y
    
    Set pElem = pMapFrame
    Set pMGeom = pElem.Geometry
    Set pMArea = pMGeom
    Set pMCentPnt = pMArea.Centroid
    dblMCentroidX = pMCentPnt.X
    dblMCentroidY = pMCentPnt.Y
    
    'Cycle through the feature class geometry, divide by the scale, and place
    'the new geometry into the map frame
    Set pFPoly = pFGeom
    Set pMNewPoly = New Polygon
    Set pFPntColl = pFGeom
    intFPntCnt = pFPntColl.PointCount
    Set pEnumVert = pFPntColl.EnumVertices
    pEnumVert.Next pFPoint, lOutPart, lOutVertex
    Set pNewPoint = New Point
    Do While Not pFPoint Is Nothing
        Select Case Me.QuadScale
            Case Is = odnr24K
                dblScale = 24000 / 12
            Case Is = odnr100K
                dblScale = 100000 / 12
            Case Is = odnr250K
                dblScale = 250000 / 12
            Case Is = odnr62K
                dblScale = 62500 / 12
            Case Is = odnr500K
                dblScale = 500000 / 12
        End Select
        dblXvert = ((pFPoint.X - dblFCentroidX) / dblScale) + dblMCentroidX
        dblYvert = ((pFPoint.Y - dblFCentroidY) / dblScale) + dblMCentroidY
        pNewPoint.PutCoords dblXvert, dblYvert
        pMNewPoly.AddPoint pNewPoint
        'Delete later, 20060525, Jim McDonald
        pTxtStrm.WriteLine "Scale = " & dblScale
        pTxtStrm.WriteLine "X_SPS83 = " & pFPoint.X & " | " & "Y_SPS83 = " & pFPoint.Y
        pTxtStrm.WriteLine "dblXVert = " & dblXvert & " | " & "dblYVert = " & dblYvert
        pTxtStrm.WriteBlankLines 1
        'End of delete later, 20060525, Jim McDonald
        pEnumVert.Next pFPoint, lOutPart, lOutVertex
    Loop
    Set ChangeMapFrameGeometry = pMNewPoly
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ChangeMapFrameGeometry " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub LoadQuadOGWellLayers()
'******************************************************************************
'
'   Program:    LoadQuadOGWellLayers
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       July 18, 2006
'   Purpose:    Creates the individual quadrangle Oil and Gas Well layer.
'   Called from: ODNRProject.ShowQuadLayers
'
'******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    Dim pQuadlyr As IFeatureLayer
    Dim pQFc As IFeatureClass
    Dim strDefQry As String
    Dim pQFCur As IFeatureCursor
    Dim pOGFCur As IFeatureCursor
    Dim pInterTopoOp2 As ITopologicalOperator2
    Dim pQFeature As IFeature
    Dim pSF As ISpatialFilter
    Dim pQf As IQueryFilter
    Dim pOgWellLyr As ILayer
    Dim pOGFLyr As IFeatureLayer
    Dim pOGFc As IFeatureClass
    Dim pOGSelSet As ISelectionSet
    Dim pTable As ITable
    Dim pWs As IWorkspace
    Dim pDataSet As IDataset
    Dim pOGoldGFlyr As IGeoFeatureLayer
    Dim pOGnewGFlyr As IGeoFeatureLayer
    Dim pRenderer As IFeatureRenderer
    Dim pHotCon As IHotlinkContainer
    Dim pHotConNew As IHotlinkContainer
    Dim pDRGLyr As ILayer
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)

    ODNR_Common.ShowMessage "Loading Quad Oil and Gas Layer...", False
    
    'Create a feature cursor for the focus quad
    Set pQuadlyr = Me.QuadFeatureLayer
    Set pQFc = pQuadlyr.FeatureClass
    
    'This section gets the 24K or 100K Quad attribute field name
    If (Me.QuadScale = odnr24K) Then
        Set pStateLayer = m_pStateLayers.GetLayerByName("Quad24K")
        strField = pStateLayer.Quad24KFieldName
    ElseIf (Me.QuadScale = odnr100K) Then
        Set pStateLayer = m_pStateLayers.GetLayerByName("Quad100K")
        strField = pStateLayer.Quad100KFieldName
    Else
        Exit Sub
    End If
    
    If (m_pQuads.QuadCount > 0) Then
        'Create the Where clause for the Query Filter
        'If the focus quad is not selected, then create the list of quad names.
        If (m_pQuads.FocusQuad Is Nothing) Then
            m_pQuads.Reset
            Set pQuad = m_pQuads.NextQuad
            Do While Not pQuad Is Nothing
                If (strDefQry = "") Then
                    strDefQry = strField & " = '" & pQuad.QuadName & "'"
                Else
                    strDefQry = strDefQry & " Or " & strField & " = '" & pQuad.QuadName & "'"
                End If
                Set pQuad = m_pQuads.NextQuad
            Loop
        'Use the focus quad.
        Else
            Set pQuad = m_pQuads.FocusQuad
            strDefQry = strField & " = '" & pQuad.QuadName & "'"
        End If
        
        Set pQf = New QueryFilter
        pQf.WhereClause = strDefQry
        
        Set pQFCur = pQFc.Search(pQf, False)
        Set pQFeature = pQFCur.NextFeature
        Set pInterTopoOp2 = pQFeature.Shape
        
        Set pSF = New SpatialFilter
        pSF.GeometryField = pQFc.shapeFieldName
        Set pSF.Geometry = pInterTopoOp2
        pSF.SpatialRel = esriSpatialRelContains
            
        'Display selected oil and gas features
        Set pOgWellLyr = Me.OGWellsLayer
        Set pOGFLyr = pOgWellLyr
        
        Set pOGFc = pOGFLyr.FeatureClass
        Set pDataSet = pOGFc
        Set pWs = pDataSet.Workspace
        Set pTable = pOGFc
    '    pOGSelSet.Search pSF, False, pOGFCur
        Set pOGSelSet = pTable.Select(pSF, esriSelectionTypeHybrid, esriSelectionOptionNormal, pWs)
        
        Dim pOGFeatureSel As IFeatureSelection
        Set pOGFeatureSel = pOGFLyr
        Set pOGFeatureSel.SelectionSet = pOGSelSet
        
        'Create a FeatureLayerDefinition using a selection set
        Dim pFlyrDef As IFeatureLayerDefinition
        Dim pFlyrNew As IFeatureLayer
        Set pFlyrDef = pOGFLyr
        Set pFlyrNew = pFlyrDef.CreateSelectionLayer("Quad Oil & Gas Wells", True, vbNullString, vbNullString)
        pOGFeatureSel.Clear
        
        'Copy the renderer for the oil and gas wells
        Set pOGoldGFlyr = pOGFLyr
        Set pOGnewGFlyr = pFlyrNew
        Set pRenderer = pOGoldGFlyr.Renderer
        Set pOGnewGFlyr.Renderer = pRenderer
        
        'Copy the Hotlinks from the statewide OG Wells layer to the Quad OG Wells layer
        Set pHotCon = pOGFLyr
        Set pHotConNew = pFlyrNew
        pHotConNew.HotlinkField = pHotCon.HotlinkField
        pHotConNew.HotlinkType = pHotCon.HotlinkType
        
        Call LabelFeatures(pFlyrNew)
        ODNR_Common.CollapseLegend pFlyrNew
        pMap.AddLayer pFlyrNew
        
        'Update the Table of Contents and the Active View
        pMxDoc.UpdateContents
        pMxDoc.ActiveView.Refresh
        
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadOGWellLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadAUMMinesOutline()
'******************************************************************************
'
'   Program:    LoadAUMMinesOutline
'   Author:     Jim McDonald
'   Company:    Ohio Division of Geological Survey
'               2045 Morse Road
'               Columbus, OH 43229-6693
'               (614) 265-6601
'   Date:       June 2, 2009
'   Purpose:    Creates an outline of the AUM Mines for cartographic purposes.
'   Called from: ODNRProject.ShowQuadLayers
'
'******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pLayer As ILayer
    Dim pFlyr As IFeatureLayer
    Dim pFlyrNew As IFeatureLayer
    Dim pGeoFlyrNew As IGeoFeatureLayer
    Dim pFc As IFeatureClass

    Dim pSRenderer As ISimpleRenderer
    Dim pSimpleFillSymbol As IFillSymbol
    Dim pLineSymbol As ISimpleLineSymbol
    Dim pLineColor As IColor
    Dim pColor As IColor
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)
    
    Set pFlyr = AUMMinesLayer
    Set pFc = pFlyr.FeatureClass
    
    'Create new feature layer
'    Set pLayer = New Layer
    Set pFlyrNew = New FeatureLayer
    Set pFlyrNew.FeatureClass = pFc
    
    pFlyrNew.Name = "Underground Mine - Outline"
    pFlyrNew.Visible = True
    
    Set pLayer = pFlyrNew
            
    'Set up the outline color, width, and style
    Set pLineColor = New RgbColor
    pLineColor.RGB = vbBlack
'    pLineColor.RGB = RGB(104, 104, 104) '60% Grey
    Set pLineSymbol = New SimpleLineSymbol
    pLineSymbol.Color = pLineColor
'    pLineSymbol.Width = 1#
    pLineSymbol.Width = 0.5
    pLineSymbol.Style = esriSLSSolid
    
    'Set the Polygon color to NULL Color
    Set pColor = New RgbColor
    pColor.NullColor = True
    
    'Finish setting up the Polygon symbol, i.e. color and outline
    Set pSimpleFillSymbol = New SimpleFillSymbol
    pSimpleFillSymbol.Color = pColor
    pSimpleFillSymbol.Outline = pLineSymbol
    
    'Create the Simple Renderer
    Set pSRenderer = New SimpleRenderer
    Set pSRenderer.Symbol = pSimpleFillSymbol
    
    Set pGeoFlyrNew = pFlyrNew
    Set pGeoFlyrNew.Renderer = pSRenderer
    Set pLayer = pGeoFlyrNew
       
    ODNR_Common.CollapseLegend pFlyrNew
    pMap.AddLayer pLayer
    pMap.MoveLayer pLayer, 3
    
    'Update the Table of Contents and the Active View
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadAUMMinesOutline " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


