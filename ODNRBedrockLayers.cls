VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRBedrockLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRBedrockLayers
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that serves individual ODNRBedrockLayers.
'
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pParent As ODNRProject
Private m_pBedrockLayers As Dictionary
Private m_lngLayerIndex As Long
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRBedrockLayers"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    m_lngLayerIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pBedrockLayers = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get LayerCount(strMapName As String) As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then LayerCount = m_pBedrockLayers.Count
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function NextLayer() As ODNRBedrockLayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (m_lngLayerIndex < m_pBedrockLayers.Count) Then
            Dim vKey As Variant
            vKey = m_pBedrockLayers.Keys(m_lngLayerIndex)
            Set NextLayer = m_pBedrockLayers.Item(vKey)
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
    
    If (m_blnIsloaded) Then m_lngLayerIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Reset " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub Refresh()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim vKey As Variant, pBedrockLayer As ODNRBedrockLayer
        For Each vKey In m_pBedrockLayers.Keys
            Set pBedrockLayer = m_pBedrockLayers.Item(vKey)
            m_pBedrockLayers.Remove vKey
            Set pBedrockLayer = Nothing
        Next
        LoadBedrockLayers
        m_lngLayerIndex = 0
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Refresh " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetLayerByName(strName As String) As ODNRBedrockLayer
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) Then
        If (m_pBedrockLayers.Exists(strName)) Then
            Set GetLayerByName = m_pBedrockLayers.Item(strName)
        Else
            Set GetLayerByName = Nothing
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".GetLayerByName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Property Get Parent() As ODNRProject
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Set Parent = m_pParent

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub Load(pApp As esriFramework.IApplication, pParent As ODNRProject)
    On Error GoTo ErrorHandler
    
    If Not (pParent Is Nothing) Then
        Set m_pApp = pApp
        Set m_pParent = pParent
        m_blnIsloaded = True
        LoadBedrockLayers
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadBedrockLayers()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pGlyr As IGroupLayer, pClyr As ICompositeLayer
        Dim pBedrockGlyrs As Dictionary, pBedrockFcExts As Dictionary
        Dim pBedrockLyr As ODNRBedrockLayer, pQuad As ODNRQuad
        Dim pWsFact As IWorkspaceFactory
        Dim pFWs As IFeatureWorkspace, pFWsManage As IFeatureWorkspaceManage
        Dim pTypesByQuadTbl As ITable
        Dim strQryDef As String, pRow As IRow, pCursor As ICursor, pQf As IQueryFilter
        Dim strType As String, pTypeList As Dictionary, lngTypeFieldIdx As Long
        Dim vKey As Variant, pKey As Variant
        Dim pTypeFc As IFeatureClass, pTypeFlyr As IFeatureLayer
        Dim pAnnoTypeFc As IFeatureClass, pAnnoTypeFlyr As IFeatureLayer
        Dim pFlyrDef As IFeatureLayerDefinition, strFlyrName As String, blnSetVisible As Boolean
        Dim pGFlyr As IGeoFeatureLayer, blnCanExport As Boolean
        Dim pDatapointsFieldList As Dictionary, strFieldName As String
        Dim strDPQryDef As String
        
        Set m_pBedrockLayers = New Dictionary
        If (m_pParent.QuadScale = odnr24K) Then
            'Handle bedrock structure layers here

            ODNR_Common.ShowMessage "Loading BS Contours...", False
            Set pBedrockGlyrs = New Dictionary
            Set pBedrockFcExts = New Dictionary
            pBedrockGlyrs.Add Key:="BS Contours Anno", Item:=FindLayer("BS Contours Anno")
            pBedrockFcExts.Add Key:="BS Contours Anno", Item:="_Contour_Anno"
            pBedrockGlyrs.Add Key:="BS Contours", Item:=FindLayer("BS Contours")
            pBedrockFcExts.Add Key:="BS Contours", Item:="_Contour"
            pBedrockGlyrs.Add Key:="BS Faults Anno", Item:=FindLayer("BS Faults Anno")
            pBedrockFcExts.Add Key:="BS Faults Anno", Item:="_Fault_Anno"
            pBedrockGlyrs.Add Key:="BS Faults", Item:=FindLayer("BS Faults")
            pBedrockFcExts.Add Key:="BS Faults", Item:="_Fault"
            pBedrockGlyrs.Add Key:="BS Contour Terminators", Item:=FindLayer("BS Contour Terminators")
            pBedrockFcExts.Add Key:="BS Contour Terminators", Item:="_Term"
            pBedrockGlyrs.Add Key:="BS Datapoints", Item:=FindLayer("BS Datapoints")
            pBedrockFcExts.Add Key:="BS Datapoints", Item:="Datapoints"
            
            For Each vKey In pBedrockGlyrs.Keys
                Set pGlyr = pBedrockGlyrs.Item(vKey)
                If (pGlyr Is Nothing) Then
                    Set pGlyr = New GroupLayer
                    pGlyr.Name = CStr(vKey)
                    pBedrockGlyrs.Item(vKey) = pGlyr
                Else
                    ClearGroupLayer pGlyr
                End If
            Next
            
            'Added section, to be used in the future once BS is loaded into SDE. 20060106, Jim McDonald
'            If (g_strSDEPGB = "SDE") Then
'                pFWs = ODNRGeologyWorkspace
'            ElseIf (g_strSDEPGB = "PGB") Then
'                Set pWsFact = New AccessWorkspaceFactory
'                Set pFWs = pWsFact.OpenFromFile(g_strBedrockDB_Path, 0)
'            End If
            Set pWsFact = New AccessWorkspaceFactory
            Set pFWs = pWsFact.OpenFromFile(g_strBedrockDB_Path, 0)
            If Not (pFWs Is Nothing) Then
                Set pFWsManage = pFWs
                ODNR_Common.ShowMessage "Loading BS Contours (searching for bedrock types)...", False
                Set pTypesByQuadTbl = pFWs.OpenTable("QuadUnitCodes")
                lngTypeFieldIdx = pTypesByQuadTbl.FindField("UNIT_CODE")
                Set pQuad = m_pParent.Quads.FocusQuad
                If Not (pQuad Is Nothing) Then
                    strQryDef = "[QUAD_ID] = '" & pQuad.ODNRQuadId & "'"
                    strDPQryDef = "[QUADCODE] = '" & pQuad.ODNRQuadId & "'"
                ElseIf (m_pParent.Quads.QuadCount > 0) Then
                    m_pParent.Quads.Reset
                    Set pQuad = m_pParent.Quads.NextQuad
                    Do While Not pQuad Is Nothing
                        If (strQryDef = "") Then
                            strQryDef = "[QUAD_ID] in ('" & pQuad.ODNRQuadId & "'"
                            strDPQryDef = "[QUADCODE] in ('" & pQuad.ODNRQuadId & "'"
                        Else
                            strQryDef = strQryDef & ",'" & pQuad.ODNRQuadId & "'"
                            strDPQryDef = strDPQryDef & ",'" & pQuad.ODNRQuadId & "'"
                        End If
                        Set pQuad = m_pParent.Quads.NextQuad
                        If (pQuad Is Nothing) Then
                            strQryDef = strQryDef & ")"
                            strDPQryDef = strDPQryDef & ")"
                        End If
                    Loop
                Else
                    strQryDef = ""
                    strDPQryDef = ""
                End If
                
                Set pTypeList = New Dictionary
                
                Set pDatapointsFieldList = New Dictionary
                
                'Note: This checks to see if BS_Datapoints is an object in the GDB
                If (pFWsManage.IsRegisteredAsObjectClass("BS_Datapoints")) Then
                    Set pTypeFc = pFWs.OpenFeatureClass("BS_Datapoints")
                End If
                
                If (strQryDef <> "") Then
                    Set pQf = New QueryFilter
                    pQf.WhereClause = strQryDef
                    Set pCursor = pTypesByQuadTbl.Search(pQf, False)
                    Set pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        strType = pRow.Value(lngTypeFieldIdx)
                        
                        If Not (pTypeFc Is Nothing) Then
                            If (pTypeFc.FindField("E" & strType) > -1) Then
                                If (Not pDatapointsFieldList.Exists(strType)) Then
                                    pDatapointsFieldList.Add Key:=strType, Item:="E" & strType
                                End If
                            End If
                        End If
                        
                        If (pTypeList.Exists(strType) = False) Then
                            pTypeList.Add Key:=strType, Item:=""
                        End If
                        Set pRow = pCursor.NextRow
                    Loop
                End If
                
                For Each pKey In pBedrockGlyrs.Keys
                    ODNR_Common.ShowMessage "Loading BS Contours (creating layer for " & CStr(pKey) & ")...", False
                    Set pGlyr = pBedrockGlyrs.Item(pKey)
                    blnCanExport = True
                    If (Right(CStr(pKey), 4) = "Anno") Then blnCanExport = False
                    blnSetVisible = True
                    If (pTypeList.Count > 0) Then
                        For Each vKey In pTypeList.Keys
                            strFlyrName = CStr(vKey) & pBedrockFcExts.Item(pKey)
                            If (pFWsManage.IsRegisteredAsObjectClass(strFlyrName)) Then
                                If (Right(strFlyrName, 4) = "Anno") Then
                                    Set pAnnoTypeFc = pFWs.OpenFeatureClass(strFlyrName)
                                    Set pQf = New QueryFilter
                                    pQf.WhereClause = strQryDef
                                    If (pAnnoTypeFc.FeatureCount(pQf) > 0) Then
                                        Set pAnnoTypeFlyr = New FDOGraphicsLayer
                                        Set pAnnoTypeFlyr.FeatureClass = pAnnoTypeFc
                                        pAnnoTypeFlyr.Name = strFlyrName
                                        Set pFlyrDef = pAnnoTypeFlyr
                                        pFlyrDef.DefinitionExpression = strQryDef
                                        pAnnoTypeFlyr.Visible = blnSetVisible
                                        pGlyr.Add pAnnoTypeFlyr
                                        Set pAnnoTypeFlyr = Nothing
                                    End If
                                Else
                                    Set pTypeFc = pFWs.OpenFeatureClass(strFlyrName)
                                    Set pQf = New QueryFilter
                                    pQf.WhereClause = strQryDef
                                    If (pTypeFc.FeatureCount(pQf) > 0) Then
                                        Set pTypeFlyr = New FeatureLayer
                                        Set pTypeFlyr.FeatureClass = pTypeFc
                                        pTypeFlyr.Name = strFlyrName
                                        Set pGFlyr = pTypeFlyr
                                        If (pTypeFc.ShapeType = esriGeometryPoint) Then
                                            Set pGFlyr.Renderer = TermRenderer
                                        Else
                                            Set pGFlyr.Renderer = ContourLineRenderer
                                        End If
                                        Set pFlyrDef = pTypeFlyr
                                        pFlyrDef.DefinitionExpression = strQryDef
                                        pTypeFlyr.Visible = blnSetVisible
                                        pGlyr.Add pTypeFlyr
                                        Set pTypeFlyr = Nothing
                                    End If
                                End If
                            Else
                                'This section is used if object class in the GDB is missing.
                                'This will be the case because it can not find the correct BS datapoints
                                'the correct BS contour layers.
                                strFlyrName = pBedrockFcExts.Item(pKey)
                                
                                'Maybe change the objectclass from "BS_" & strFlyyName to "BS_Datapoints" to make it easier for me.
                                If (pFWsManage.IsRegisteredAsObjectClass("BS_" & strFlyrName)) Then
                                    If (pDatapointsFieldList.Exists(vKey)) Then
                                        strFieldName = pDatapointsFieldList.Item(vKey)
                                        
                                        'This is interesting, the pTypeFc object is being reused.
                                        Set pTypeFc = pFWs.OpenFeatureClass("BS_" & strFlyrName)
'                                        MsgBox strFieldName & vbCrLf & strFlyrName
                                        Set pQf = New QueryFilter
                                        pQf.WhereClause = strDPQryDef & " And ([TYPE] >=1 and [TYPE] <=8) and [" & strFieldName & "] Is Not Null"
                                        If (pTypeFc.FeatureCount(pQf) > 0) Then
                                            Set pTypeFlyr = New FeatureLayer
                                            Set pTypeFlyr.FeatureClass = pTypeFc
                                            pTypeFlyr.Name = CStr(vKey) & "_" & strFlyrName
                                            Set pGFlyr = pTypeFlyr
'                                            This is my version of a BS Datapoints renderer
'                                            Set pGFlyr.Renderer = BSDatapointRenderer
                                            Set pGFlyr.Renderer = DatapointsRenderer
                                            Set pFlyrDef = pTypeFlyr
                                            pFlyrDef.DefinitionExpression = pQf.WhereClause
                                            pTypeFlyr.Visible = blnSetVisible
                                            CreateDatapointLabels pTypeFlyr, strFieldName
                                            pGFlyr.DisplayAnnotation = True
                                            pGlyr.Add pTypeFlyr
                                            Set pTypeFlyr = Nothing
                                        End If
                                    End If
                                End If
                            End If
                            blnSetVisible = False
                        Next
                    End If
                    Set pBedrockLyr = New ODNRBedrockLayer
                    pGlyr.Expanded = False
                    'Turn off layers for BG24K maps
                    If (m_pParent.ProjectType = odnrGeology) Then
                        pGlyr.Visible = False
                    Else 'Turn on layers for BS24K maps
                        pGlyr.Visible = True
                    End If
                    pBedrockLyr.LoadFromLayer m_pApp, Me, pGlyr, esriDTFeatureClass, odnrGeologyMap, blnCanExport
                    m_pBedrockLayers.Add Key:=pGlyr.Name, Item:=pBedrockLyr
                Next
            End If
            ODNR_Common.ShowMessage "", False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadBedrockLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function FindLayer(strName As String) As ILayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, pLayer As ILayer, lngIdx As Long
        Set pMxDoc = m_pApp.Document
        Set pMap = m_pParent.ProjectMap(odnrGeologyMap)
        If Not (pMap Is Nothing) Then
            For lngIdx = 0 To pMap.LayerCount - 1
                If (TypeOf pMap.Layer(lngIdx) Is IGroupLayer) Then
                    If (strName = pMap.Layer(lngIdx).Name) Then
                        Set pLayer = pMap.Layer(lngIdx)
                        Exit For
                    End If
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

Private Sub ClearGroupLayer(pGlyr As IGroupLayer)
    On Error GoTo ErrorHandler
    
    If (Not pGlyr Is Nothing) Then
        Dim pLyr As ILayer, pClyr As ICompositeLayer
        Set pClyr = pGlyr
        Do While pClyr.Count > 0
            Set pLyr = pClyr.Layer(0)
            pGlyr.Delete pLyr
        Loop
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ClearGroupLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function ContourLineRenderer() As ISimpleRenderer
    On Error GoTo ErrorHandler

    Dim pLineSym As ISimpleLineSymbol, pColor As IRgbColor, pSimpleRend As ISimpleRenderer
    Set pLineSym = New SimpleLineSymbol
    Set pColor = New RgbColor
    pColor.RGB = RGB(0, 0, 0)
    pLineSym.Color = pColor
    pLineSym.Width = 1
    pLineSym.Style = esriSLSSolid
    Set pSimpleRend = New SimpleRenderer
    Set pSimpleRend.Symbol = pLineSym
    Set ContourLineRenderer = pSimpleRend
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ContourLineRenderer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function TermRenderer() As IUniqueValueRenderer
    On Error GoTo ErrorHandler

    Dim pUniqRender As IUniqueValueRenderer
    Dim pSym As ISymbol, pMLyrMarkSym As IMultiLayerMarkerSymbol
    Dim pSym1 As ISymbol, pSym2 As ISymbol
    Dim pCharMarkSym As ICharacterMarkerSymbol, pMarkSym As IMarkerSymbol
    Dim pQMarkSym As IMarkerSymbol, pDotMarkSym As IMarkerSymbol, pClone As IClone, pClone2 As IClone
    Dim lngQMarkId As Long, lngDotMarkId As Long
    Dim pStyleGallery As IStyleGallery, pEnumStyle As IEnumStyleGalleryItem
    Dim pStyleItem As IStyleGalleryItem, pMarkerSym As IMarkerSymbol
    
    Set pStyleGallery = New StyleGallery
    Set pEnumStyle = pStyleGallery.Items("Marker Symbols", "ESRI.style", "")
    Set pStyleItem = pEnumStyle.Next
    Do While Not pStyleItem Is Nothing
        If (pStyleItem.Name = "Question") Then
            Set pQMarkSym = pStyleItem.Item
            lngQMarkId = pStyleItem.ID
        ElseIf (pStyleItem.Name = "Circle 1") Then
            Set pDotMarkSym = pStyleItem.Item
            lngDotMarkId = pStyleItem.ID
        End If
        If (lngQMarkId > 0) And (lngDotMarkId > 0) Then
            Exit Do
        End If
        Set pStyleItem = pEnumStyle.Next
    Loop
    Set pUniqRender = New UniqueValueRenderer
    Set pMLyrMarkSym = New MultiLayerMarkerSymbol
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.CharacterIndex = 103
    pCharMarkSym.Size = 8
    pCharMarkSym.XOffset = 5
    pCharMarkSym.YOffset = 5
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.CharacterIndex = 33
    pCharMarkSym.Size = 8
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pSym = pMLyrMarkSym
    pUniqRender.FieldCount = 1
    pUniqRender.Field(0) = "uncertain"
    pUniqRender.FieldType(0) = False
    
    pUniqRender.AddValue 1, "Uncertain", pSym
    pUniqRender.Label(1) = "Yes"
    
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.CharacterIndex = 33
    pCharMarkSym.Size = 8
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    
    pUniqRender.AddValue 0, "Uncertain", pSym
    pUniqRender.Label(0) = "No"
    Set TermRenderer = pUniqRender
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".TermRenderer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function BSDatapointRenderer() As IUniqueValueRenderer
    On Error GoTo ErrorHandler

    Dim pUniqRender As IUniqueValueRenderer
    Dim pMSym(200) As IMarkerSymbol
    Dim strMSym(200) As String
    Dim pSym As ISymbol, pMLyrMarkSym As IMultiLayerMarkerSymbol
    Dim pSym1 As ISymbol, pSym2 As ISymbol
    Dim pCharMarkSym As ICharacterMarkerSymbol, pMarkSym As IMarkerSymbol
    Dim pQMarkSym As IMarkerSymbol, pDotMarkSym As IMarkerSymbol, pClone As IClone, pClone2 As IClone
    Dim lngQMarkId As Long, lngDotMarkId As Long
    Dim pStyleGallery As IStyleGallery, pEnumStyle As IEnumStyleGalleryItem
    Dim pStyleItem As IStyleGalleryItem, pMarkerSym As IMarkerSymbol
    Dim intArrayCount As Integer
    
    Set pStyleGallery = New StyleGallery 'Note: I might have to change this to using the MXD style gallery.  See DoRenderer proceedure for the example.
    Set pEnumStyle = pStyleGallery.Items("Marker Symbols", "S:\Basemaps\ArcGIS_styles\BS_24K.style", "")
    
    intArrayCount = 0
    Set pStyleItem = pEnumStyle.Next
    Do While Not pStyleItem Is Nothing
        Set pMSym(intArrayCount) = pStyleItem.Item
        strMSym(intArrayCount) = pStyleItem.Name
'        MsgBox "pStyleItem Name = " & pStyleItem.Name
'        If (pStyleItem.Name = "Question") Then
'            Set pQMarkSym = pStyleItem.Item
'            lngQMarkId = pStyleItem.ID
'        ElseIf (pStyleItem.Name = "Circle 1") Then
'            Set pDotMarkSym = pStyleItem.Item
'            lngDotMarkId = pStyleItem.ID
'        End If
'        If (lngQMarkId > 0) And (lngDotMarkId > 0) Then
'            Exit Do
'        End If
        Set pStyleItem = pEnumStyle.Next
        intArrayCount = intArrayCount + 1
    Loop
    
    Set pUniqRender = New UniqueValueRenderer
'    Set pMLyrMarkSym = New MultiLayerMarkerSymbol
'    Set pCharMarkSym = New CharacterMarkerSymbol
'    pCharMarkSym.CharacterIndex = 103
'    pCharMarkSym.Size = 8
'    pCharMarkSym.XOffset = 5
'    pCharMarkSym.YOffset = 5
'    pMLyrMarkSym.AddLayer pCharMarkSym
'
'    Set pCharMarkSym = New CharacterMarkerSymbol
'    pCharMarkSym.CharacterIndex = 33
'    pCharMarkSym.Size = 8
'    pMLyrMarkSym.AddLayer pCharMarkSym
'    Set pSym = pMLyrMarkSym
'    pUniqRender.FieldCount = 1
'    pUniqRender.Field(0) = "uncertain"
'    pUniqRender.FieldType(0) = False
'
'    pUniqRender.AddValue 1, "Uncertain", pSym
'    pUniqRender.Label(1) = "Yes"
'
'    Set pCharMarkSym = New CharacterMarkerSymbol
'    pCharMarkSym.CharacterIndex = 33
'    pCharMarkSym.Size = 8
'    Set pMarkSym = pCharMarkSym
'    Set pSym = pMarkSym
'
'    pUniqRender.AddValue 0, "Uncertain", pSym
'    pUniqRender.Label(0) = "No"
    pUniqRender.FieldCount = 1
    pUniqRender.Field(0) = "TYPE"

    Dim j As Integer
    
    For j = 0 To intArrayCount
        Select Case strMSym(j)
        Case 1
            pUniqRender.AddValue 1, "TYPE", pMSym(j)
            pUniqRender.Label(1) = "Geophysical Log"
'            MsgBox pUniqRender.Label(1)
        Case 2
            pUniqRender.AddValue 2, "TYPE", pMSym(j)
            pUniqRender.Label(2) = "Described Well Cuttings"
'            MsgBox pUniqRender.Label(2)
        Case 3
            pUniqRender.AddValue 3, "TYPE", pMSym(j)
            pUniqRender.Label(3) = "Driller's Log From Completion Card"
'            MsgBox pUniqRender.Label(3)
        Case 4
            pUniqRender.AddValue 4, "TYPE", pMSym(j)
            pUniqRender.Label(4) = "Diamond Drill Hole (ODGS Core)"
'            MsgBox pUniqRender.Label(4)
        Case 5
            pUniqRender.AddValue 5, "TYPE", pMSym(j)
            pUniqRender.Label(5) = "Water Well Information"
'            MsgBox pUniqRender.Label(5)
        Case 6
            pUniqRender.AddValue 6, "TYPE", pMSym(j)
            pUniqRender.Label(6) = "ODGS Open File Description (Measured Section)"
'            MsgBox pUniqRender.Label(6)
        Case 7
            pUniqRender.AddValue 7, "TYPE", pMSym(j)
            pUniqRender.Label(7) = "Field Location Descriptions"
'            MsgBox pUniqRender.Label(7)
        End Select
    Next
    Set BSDatapointRenderer = pUniqRender
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".BSDatapointRenderer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function DatapointsLabelTextSymbol(blnUp As Boolean) As ITextSymbol
    On Error GoTo ErrorHandler
    
    Dim pTextSym As ITextSymbol, pSTextSym As ISimpleTextSymbol
    Dim pRGBColor As IRgbColor, pFontDisp As IFontDisp
    Set pTextSym = New TextSymbol
    Set pRGBColor = New RgbColor
    pRGBColor.Red = 0
    pRGBColor.Blue = 0
    pRGBColor.Green = 0
    pTextSym.Color = pRGBColor
    pTextSym.HorizontalAlignment = esriTHACenter
    pTextSym.VerticalAlignment = esriTVACenter
    pTextSym.Size = 8
    Set pFontDisp = New StdFont
    pFontDisp.Name = "Arial"
    pFontDisp.Size = 8
    pFontDisp.Bold = False
    pTextSym.Font = pFontDisp
    Set pSTextSym = pTextSym
    If (blnUp) Then
        pSTextSym.YOffset = 5
    Else
        pSTextSym.YOffset = -5
    End If
    Set DatapointsLabelTextSymbol = pSTextSym
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".DatapointsLabelTextSymbol " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub CreateDatapointLabels(pFlyr As IFeatureLayer, strFieldName As String)
    On Error GoTo ErrorHandler

    Dim pMxDoc As IMxDocument, pMap As IMap
    Set pMxDoc = m_pApp.Document
    Set pMap = m_pParent.ProjectMap(odnrGeologyMap)

    Dim pGFlyr As IGeoFeatureLayer, pAnnoLyrPropsColl As IAnnotateLayerPropertiesCollection
    Dim pAnnoLyrProps As IAnnotateLayerProperties, pLblEngLyrProps As ILabelEngineLayerProperties
    Dim pGraCont As IGraphicsContainer, pBOPLyrProps As IBasicOverposterLayerProperties
    Dim pOverposterLyrProps2 As IOverposterLayerProperties2 'Added to place overlapping labels, 20060601, Jim McDonald
    Dim pPtPlacePriorities As IPointPlacementPriorities
    Dim pRemoveAnnoLyrProps As IAnnotateLayerProperties
    
    Set pGFlyr = pFlyr
    Set pAnnoLyrPropsColl = pGFlyr.AnnotationProperties
    Set pGraCont = pMap
    
    Set pLblEngLyrProps = New LabelEngineLayerProperties
    Set pBOPLyrProps = pLblEngLyrProps.BasicOverposterLayerProperties
    pBOPLyrProps.FeatureType = esriOverposterPoint
    pBOPLyrProps.FeatureWeight = esriLowWeight
    pBOPLyrProps.GenerateUnplacedLabels = True
    pBOPLyrProps.PointPlacementMethod = esriAroundPoint
    Set pOverposterLyrProps2 = pBOPLyrProps 'Added to place overlapping labels, 20060601, Jim McDonald
    pOverposterLyrProps2.TagUnplaced = False 'Added to place overlapping labels, 20060601, Jim McDonald
    Set pPtPlacePriorities = New PointPlacementPriorities
    pPtPlacePriorities.AboveCenter = 1
    pPtPlacePriorities.AboveLeft = 2
    pPtPlacePriorities.AboveRight = 2
    pPtPlacePriorities.BelowCenter = 0
    pPtPlacePriorities.BelowRight = 0
    pPtPlacePriorities.BelowLeft = 0
    pPtPlacePriorities.CenterLeft = 0
    pPtPlacePriorities.CenterRight = 0
    pBOPLyrProps.PointPlacementPriorities = pPtPlacePriorities
    pLblEngLyrProps.Expression = "[DATANO]"
    pLblEngLyrProps.IsExpressionSimple = True
    Set pLblEngLyrProps.Symbol = DatapointsLabelTextSymbol(True)
    Set pAnnoLyrProps = pLblEngLyrProps
    pAnnoLyrProps.AddUnplacedToGraphicsContainer = False
    pAnnoLyrProps.CreateUnplacedElements = False
    pAnnoLyrProps.FeatureLinked = False
    pAnnoLyrProps.UseOutput = True
    pAnnoLyrProps.Class = "STATEMAP ID"
    pAnnoLyrProps.DisplayAnnotation = True
    pAnnoLyrProps.LabelWhichFeatures = esriVisibleFeatures
    Set pAnnoLyrProps.GraphicsContainer = pGraCont
    pAnnoLyrPropsColl.Add pAnnoLyrProps
    
    Set pLblEngLyrProps = New LabelEngineLayerProperties
    Set pBOPLyrProps = pLblEngLyrProps.BasicOverposterLayerProperties
    pBOPLyrProps.FeatureType = esriOverposterPoint
    pBOPLyrProps.FeatureWeight = esriLowWeight
    pBOPLyrProps.GenerateUnplacedLabels = True
    pBOPLyrProps.PointPlacementMethod = esriAroundPoint
    Set pOverposterLyrProps2 = pBOPLyrProps 'Added to place overlapping labels, 20060601, Jim McDonald
    pOverposterLyrProps2.TagUnplaced = False 'Added to place overlapping labels, 20060601, Jim McDonald
    Set pPtPlacePriorities = New PointPlacementPriorities
    pPtPlacePriorities.AboveCenter = 0
    pPtPlacePriorities.AboveLeft = 0
    pPtPlacePriorities.AboveRight = 0
    pPtPlacePriorities.BelowCenter = 1
    pPtPlacePriorities.BelowRight = 2
    pPtPlacePriorities.BelowLeft = 2
    pPtPlacePriorities.CenterLeft = 0
    pPtPlacePriorities.CenterRight = 0
    pBOPLyrProps.PointPlacementPriorities = pPtPlacePriorities
    pLblEngLyrProps.Expression = "[" & strFieldName & "]"
    pLblEngLyrProps.IsExpressionSimple = True
    Set pLblEngLyrProps.Symbol = DatapointsLabelTextSymbol(False)
    Set pAnnoLyrProps = pLblEngLyrProps
    pAnnoLyrProps.AddUnplacedToGraphicsContainer = False
    pAnnoLyrProps.CreateUnplacedElements = False
    pAnnoLyrProps.FeatureLinked = False
    pAnnoLyrProps.UseOutput = True
    pAnnoLyrProps.Class = "ELEVATION"
    pAnnoLyrProps.DisplayAnnotation = True
    pAnnoLyrProps.LabelWhichFeatures = esriVisibleFeatures
    Set pAnnoLyrProps.GraphicsContainer = pGraCont
    pAnnoLyrPropsColl.Add pAnnoLyrProps
    
    pAnnoLyrPropsColl.QueryItem 0, pRemoveAnnoLyrProps, Nothing, Nothing
    pAnnoLyrPropsColl.Remove pRemoveAnnoLyrProps

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CreateDatapointLabels " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function DatapointsRenderer() As IUniqueValueRenderer
    On Error GoTo ErrorHandler

    Dim pUniqRender As IUniqueValueRenderer, pFontDisp As IFontDisp
    Dim pSym As ISymbol, pMLyrMarkSym As IMultiLayerMarkerSymbol
    Dim pCharMarkSym As ICharacterMarkerSymbol, pMarkSym As IMarkerSymbol
    
    Set pUniqRender = New UniqueValueRenderer
    'Type 1
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 33
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 1, "WELL DATA", pSym
    pUniqRender.Label(1) = "Geophyisical Log"
    
    'Type 2
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 63
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 2, "WELL DATA", pSym
    pUniqRender.Label(2) = "Described Well Cuttings"
    
    'Type 3
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 62
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 3, "WELL DATA", pSym
    pUniqRender.Label(3) = "Driller's Log from Completion Card"
    
    'Type 4
    Set pMLyrMarkSym = New MultiLayerMarkerSymbol
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 46
    pCharMarkSym.Size = 10
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pFontDisp = New StdFont
    pFontDisp.Name = "Arial"
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font = pFontDisp
    pCharMarkSym.CharacterIndex = 68
    pCharMarkSym.Size = 7
    pCharMarkSym.XOffset = 6
    pCharMarkSym.YOffset = 4
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font = pFontDisp
    pCharMarkSym.CharacterIndex = 68
    pCharMarkSym.Size = 7
    pCharMarkSym.XOffset = 11
    pCharMarkSym.YOffset = 4
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font = pFontDisp
    pCharMarkSym.CharacterIndex = 72
    pCharMarkSym.Size = 7
    pCharMarkSym.XOffset = 16
    pCharMarkSym.YOffset = 4
    pMLyrMarkSym.AddLayer pCharMarkSym
    Set pSym = pMLyrMarkSym
    pUniqRender.AddValue 4, "WELL DATA", pSym
    pUniqRender.Label(4) = "Diamond Drill Hole"
    
    'Type 5
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 40
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 5, "WELL DATA", pSym
    pUniqRender.Label(5) = "Water Well Informaiton"
    
    'Type 6
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 68
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 6, "OUTCROP DATA", pSym
    pUniqRender.Label(6) = "OGS Open File Description"
    
    'Type 7
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 69
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 7, "OUTCROP DATA", pSym
    pUniqRender.Label(7) = "Field Location Description"
    
    'Type 8
    Set pCharMarkSym = New CharacterMarkerSymbol
    pCharMarkSym.Font.Name = "ESRI Default Marker"
    pCharMarkSym.CharacterIndex = 68
    pCharMarkSym.Size = 10
    Set pMarkSym = pCharMarkSym
    Set pSym = pMarkSym
    pUniqRender.AddValue 8, "OUTCROP DATA", pSym
    pUniqRender.Label(8) = "OGS Open File Description"
    
    pUniqRender.FieldCount = 1
    pUniqRender.Field(0) = "TYPE"
    pUniqRender.FieldType(0) = False
    
    Set DatapointsRenderer = pUniqRender
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".DatapointsRenderer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


