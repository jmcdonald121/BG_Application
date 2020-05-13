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

    If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap
        Dim pQuadLayer As ODNRQuadLayer, pLayer As ILayer
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
        If (Me.ProjectType = odnrGeology) Then LoadGeologyLayer
        pMxDoc.ActiveView.Refresh
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
        Dim pODNRMapEl As ODNRMapElement, vKey As Variant
        Dim pColorBlack As IRgbColor
        
        If Not (m_pLayoutElements Is Nothing) Then
            Set pMxDoc = m_pApp.Document
            Set pPageLayout = pMxDoc.PageLayout
            Set pGraCont = pPageLayout
            pGraCont.Reset
            Set pColorBlack = New RgbColor
            pColorBlack.RGB = RGB(0, 0, 0)
            For Each vKey In m_pLayoutElements.Keys
                Set pODNRMapEl = m_pLayoutElements.Item(vKey)
                pODNRMapEl.SetTextToKey
                If (m_pQuads.QuadCount > 1) Then Set pODNRMapEl.SymbolColor = pColorBlack
                pODNRMapEl.UpdateGraphicsContainer
                m_pLayoutElements.Remove vKey
                Set pODNRMapEl = Nothing
            Next
        End If
        'Empty the list once everything has been reset and then make the list nothing
        Set m_pLayoutElements = Nothing
        m_blnDoResetLayout = False
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
        Dim blnSingleQuad As Boolean, blnListCounties As Boolean
        Dim pCountyList As Collection, lngIdx As Long
        Dim pFWs As IFeatureWorkspace, pWsFact As IWorkspaceFactory, pTypesTbl As ITable
        Dim pQf As IQueryFilter, pTypesRow As IRow, pCursor As ICursor
        Dim pQuadLayer As ODNRQuadLayer, pBedrockLayer As ODNRBedrockLayer
        Dim pGlyr As IGroupLayer, pClyr As ICompositeLayer
        Dim lngLyrIdx As Long, pLayer As ILayer, strBedrockType As String
        Dim dblWidth As Double, dblHeight As Double
        
        Dim strText As String, strMapType As String
        Dim strTitle As String, strResYear As String
        Dim strBedrock As String, strQuadName As String 'remember to preface with the word Map
        Dim strLatN As String, strLatS As String
        Dim strLongW As String, strLongE As String
        Dim strAuthor As String, strCounties As String
        Dim strInterval As String
        
        Dim pODNRMapEl As ODNRMapElement, pColor As IRgbColor
        
        Set pMxDoc = m_pApp.Document
        Set pMap = pMxDoc.FocusMap
        Set pActiveView = pMap
        
        blnSingleQuad = False
        If Not (m_pQuads.FocusQuad Is Nothing) Then
            blnSingleQuad = True
        End If
        If (blnSingleQuad) Then
            Set pRow = m_pQuads.QuadLayoutTextRecord
            If (pRow Is Nothing) Then
                MsgBox "Unable to locate title record to customize layout text."
                Exit Sub
            End If
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
                    If (pRow.Fields.FindField("QUAD") > -1) Then strBedrock = "BEDROCK STRUCTURE OF THE " _
                    & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
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
                If (pRow.Fields.FindField("REV_DATE") > -1) Then strResYear = FieldValue(pRow, pRow.Fields.FindField("REV_DATE"))
                If (pRow.Fields.FindField("QUAD") > -1) Then strQuadName = "Map " & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
                If (Me.ProjectType = odnrGeology) Then
                    If (pRow.Fields.FindField("QUAD") > -1) Then strBedrock = "BEDROCK GEOLOGY OF THE " _
                    & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
                ElseIf (Me.ProjectType = odnrTopography) Then
                    If (pRow.Fields.FindField("QUAD") > -1) Then strBedrock = "BEDROCK TOPOGRAPHY OF THE " _
                    & FieldValue(pRow, pRow.Fields.FindField("QUAD"))
                End If
            End If
        End If
    
        Me.QuadFeatureLayer.Visible = True
        
        'go to layout
        Set pPageLayout = pMxDoc.PageLayout
        Set pMxDoc.ActiveView = pPageLayout
        Set pMxDoc = m_pApp.Document
        'pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
        
        Set pActiveView = pMxDoc.PageLayout
        Set pGraCont = pPageLayout
        pGraCont.Reset
        
        Set m_pLayoutElements = New Dictionary
        
        blnListCounties = False
        If (Me.QuadScale = odnr24K) Then
            Set pCountyList = m_pQuads.CountyList
            If Not (pCountyList Is Nothing) Then
                If (pCountyList.Count >= 1) Then
                    If (pCountyList.Count = 1) Then
                        strCounties = pCountyList.Item(1) & " COUNTY, OHIO"
                    ElseIf (pCountyList.Count > 1) Then
                        For lngIdx = 1 To pCountyList.Count
                            If (lngIdx = 1) Then
                                strCounties = pCountyList.Item(lngIdx)
                            ElseIf (lngIdx = pCountyList.Count) Then
                                strCounties = strCounties & " and " & pCountyList.Item(lngIdx) & " COUNTIES, OHIO"
                            Else
                                strCounties = strCounties & ", " & pCountyList.Item(lngIdx)
                            End If
                        Next
                    End If
                    blnListCounties = True
                End If
            End If
        End If
        
        Set pColor = New RgbColor
        pColor.RGB = RGB(255, 255, 255)
        Set pElem = pGraCont.Next

        Do While Not pElem Is Nothing
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
                Set pODNRMapEl = New ODNRMapElement
                pODNRMapEl.LoadElement pElem, strText, pGraCont
                Select Case strText
                    Case "Northing1"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLatN
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Easting1"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLongW
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Northing2"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLatN
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Easting2"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLongE
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Northing3"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLatS
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Easting3"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLongW
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Northing4"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLatS
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Easting4"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strLongE
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Mapquadrangle"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strQuadName
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "quadname"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strQuadName
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Revised 02/2002"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strResYear
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "Author"
                        pODNRMapEl.ElementText = strAuthor
                    Case "Title Line 1"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strTitle
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                    Case "COUNTY"
                        If blnListCounties Then
                            pODNRMapEl.ElementText = strCounties
                        Else
                            pODNRMapEl.ElementText = "OHIO"
                        End If
                    Case "BEDROCK"
                        If (blnSingleQuad) Then
                            If (Me.QuadScale = odnr24K) Then
                                pODNRMapEl.ElementText = UCase$(strTitle)
                            ElseIf (Me.QuadScale = odnr100K) Then
                                pODNRMapEl.ElementText = strBedrock
                            End If
                        Else
                            If (Me.ProjectType = odnrGeology) Then
                                pODNRMapEl.ElementText = "BEDROCK GEOLOGY"
                            ElseIf (Me.ProjectType = odnrTopography) Then
                                pODNRMapEl.ElementText = "BEDROCK TOPOGRAPHY"
                            ElseIf (Me.ProjectType = odnrBedrockStructure) Then
                                pODNRMapEl.ElementText = "BEDROCK STRUCTURE"
                            End If
                        End If
                    Case "Contour Interval"
                        If (blnSingleQuad) Then
                            pODNRMapEl.ElementText = strInterval
                        Else
                            Set pODNRMapEl.SymbolColor = pColor
                        End If
                End Select
                    
                pODNRMapEl.UpdateGraphicsContainer
                m_pLayoutElements.Add Key:=strText, Item:=pODNRMapEl
            End If
            Set pElem = pGraCont.Next
        Loop
        '**********
        'try legend thingy here
        
        'Call fixLegend(strProjName, m_pApp)
        
        '***********
        
        Set pActiveView = Me.ProjectMap(odnrGeologyMap)
        If (Me.QuadScale = odnr24K) Then
            pActiveView.FocusMap.MapScale = 24000
        ElseIf (Me.QuadScale = odnr100K) Then
            pActiveView.FocusMap.MapScale = 100000
        End If
        m_blnInDataView = False
        m_blnDoResetLayout = True
        Set pActiveView = pMxDoc.PageLayout
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

Private Sub LoadGeologyLayer()
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
    Dim pMapSurround As IMapSurround, pLegend As ILegend
    Dim pLegendItem As ILegendItem, lngMapSIdx As Long
    Dim pMxDoc As IMxDocument, pMap As IMap
    Dim pQuad As ODNRQuad, strField As String, pStateLayer As ODNRStateLayer
    Dim pEnv As IEnvelope, pLegendFormat As ILegendFormat, pLegendLayout As ILegendLayout
    
    Set pMxDoc = m_pApp.Document
    Set pMap = Me.ProjectMap(odnrGeologyMap)
    
    Set pStateLayer = m_pStateLayers.GetLayerByName(Me.GeologyUnitsPolygonLayer.Name)
    If Not (pStateLayer Is Nothing) Then
        If (Me.QuadScale = odnr24K) Then
            strField = pStateLayer.Quad24KFieldName
        ElseIf (Me.QuadScale = odnr100K) Then
            strField = pStateLayer.Quad100KFieldName
        End If
    Else
        Exit Sub
    End If
    ODNR_Common.ShowMessage "Loading Geology Layer...", False
    If (m_pQuads.QuadCount > 0) Then
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
        Else
            Set pQuad = m_pQuads.FocusQuad
            strDefQry = strField & " = '" & pQuad.QuadName & "'"
        End If
        Set pNewFlyr = New FeatureLayer
        Set pGFlyr = Me.GeologyUnitsPolygonLayer
        Set pUniqueValueRend = pGFlyr.Renderer
        Set pNewFlyr.FeatureClass = pGFlyr.FeatureClass
        Set pFlyrDef = pNewFlyr
        pFlyrDef.DefinitionExpression = strDefQry
        Set pDs = pGFlyr.FeatureClass
        Set pFWs = pDs.Workspace
        Set pQDef = pFWs.CreateQueryDef
        With pQDef
            .Tables = pDs.Name
            .SubFields = "DISTINCT(UNIT_CODE)"
            .WhereClause = strDefQry
            Set pCursor = .Evaluate
        End With
        Set pRow = pCursor.NextRow
        Set pUniqueUnitCodes = New Dictionary
        strDefQry = ""
        Do While Not pRow Is Nothing
            If (IsNull(pRow.Value(0)) = False) Then
                If (Trim(pRow.Value(0)) <> "") Then
                    If (strDefQry = "") Then
                        strDefQry = "[UNIT_CODE] in ('" & pRow.Value(0) & "'"
                    Else
                        strDefQry = strDefQry & ",'" & pRow.Value(0) & "'"
                    End If
                    pUniqueUnitCodes.Add Key:=pRow.Value(0), Item:=""
                End If
            End If
            Set pRow = pCursor.NextRow
        Loop
        strDefQry = strDefQry & ")"
        
        lngCSDAGEFieldIdx = pGFlyr.FeatureClass.FindField("CSD_AGE")
        lngSTRATCODEFieldIdx = pGFlyr.FeatureClass.FindField("STRATCODE")
        lngNAMEFieldIdx = pGFlyr.FeatureClass.FindField("NAME")
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
        Set pGFlyr = pNewFlyr
        Set pGFlyr.Renderer = pNewUniqueValueRend
        Set pRendererPropPage = New UniqueValuePropertyPage
        pGFlyr.RendererPropertyPageClassID = pRendererPropPage.ClassID
        
        For lngMapSIdx = 0 To pMap.MapSurroundCount - 1
            Set pMapSurround = pMap.MapSurround(lngMapSIdx)
            If (TypeOf pMapSurround Is ILegend) Then
                Set pLegend = pMapSurround
                Exit For
            End If
        Next
        Set pEnv = New Envelope
        pEnv.PutCoords 2, 33 - (2 + 21 * (pNewUniqueValueRend.ValueCount / 71)), 4, 33
        If Not (pLegend Is Nothing) Then
            Set pLegendFormat = pLegend.Format
            Set pLegendLayout = pLegend.Format
            pLegendLayout.ScaleGraphicsOnResize = False
            pLegendFormat.DefaultPatchWidth = 50
            pLegendFormat.DefaultPatchHeight = 12
            pLegendFormat.ShowTitle = True
            pLegend.Title = "Legend"
            pLegendFormat.TitleSymbol.Font.Size = 24
            If (pLegend.ItemCount = 0) Then
                Set pLegendItem = New VerticalLegendItem
                pLegend.AddItem pLegendItem
            End If
            Set pLegendItem = pLegend.Item(0)
            Set pLayer = pGFlyr
            Set pLegendItem.Layer = pLayer
            pLegendItem.LayerNameSymbol.Font.Size = 14
            pLegendItem.HeadingSymbol.Font.Size = 12
'            pLegend.ClearItems
            pLegend.AutoAdd = False
            pLegend.AutoVisibility = False
'            pLegend.AddItem pLegendItem
'            pLegend.FitToBounds pMxdoc.ActiveView.ScreenDisplay, pEnv, True
            pLegend.Refresh
        End If
        
        pGFlyr.Visible = True
        pGFlyr.Name = "Quad Geology Units Ply"
        Set pFLayer = pGFlyr
        ODNR_Common.CollapseLegend pFLayer
        pMap.AddLayer pGFlyr
        pMap.MoveLayer pGFlyr, pMap.LayerCount - 1
        pMxDoc.UpdateContents
    End If
    ODNR_Common.ShowMessage "", False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadGeologyLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

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

