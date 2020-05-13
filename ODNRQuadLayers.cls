VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRQuadLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRQuadLayers
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that serves individual ODNRQuadLayers.
'
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pParent As ODNRProject
Private m_pQuadLayers As Dictionary
Private m_lngLayerIndex As Long
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRQuadLayers"

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
    
    Set m_pQuadLayers = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get LayerCount(strMapName As String) As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then LayerCount = m_pQuadLayers.Count
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function NextLayer() As ODNRQuadLayer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (m_lngLayerIndex < m_pQuadLayers.Count) Then
            Dim vKey As Variant
            vKey = m_pQuadLayers.Keys(m_lngLayerIndex)
            Set NextLayer = m_pQuadLayers.Item(vKey)
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
        Dim vKey As Variant, pQuadLayer As ODNRQuadLayer
        For Each vKey In m_pQuadLayers.Keys
            Set pQuadLayer = m_pQuadLayers.Item(vKey)
            m_pQuadLayers.Remove vKey
            Set pQuadLayer = Nothing
        Next
        LoadQuadLayers
        m_lngLayerIndex = 0
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Refresh " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetLayerByName(strName As String) As ODNRQuadLayer
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) Then
        If (m_pQuadLayers.Exists(strName)) Then
            Set GetLayerByName = m_pQuadLayers.Item(strName)
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
        LoadQuadLayers
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadQuadLayers()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pGlyr As IGroupLayer
        Dim pLyr As ILayer, pRLyr As IRasterLayer, pRDs As IRasterDataset
        Dim pSuffixList As Collection
        Dim strSuffix As String, lngSuffixIdx As Long
        Dim pQuadlyr As ODNRQuadLayer, pQuad As ODNRQuad
        Dim strPath As String, strRepairPath As String, strName As String
        Dim pWsFact As IWorkspaceFactory, pRWs As IRasterWorkspace
        Dim pLEffects As ILayerEffects
        
        Set m_pQuadLayers = New Dictionary
        If (m_pParent.QuadScale = odnr24K) Then
            'Handle the DRGs here
            ODNR_Common.ShowMessage "Loading DRGs...", False
            Set pGlyr = FindLayer("DRGs")
            If (pGlyr Is Nothing) Then
                Set pGlyr = New GroupLayer
                pGlyr.Name = "DRGs"
            Else
                ClearGroupLayer pGlyr
            End If
            m_pParent.Quads.Reset
            Set pQuad = m_pParent.Quads.NextQuad
            Do While Not pQuad Is Nothing
                strPath = g_strDRG_Path & AddUnderscores(pQuad.QuadName) & ".lyr"
                strRepairPath = g_strDRGLOC_Path '& AddUnderscores(pQuad.QuadName) & ".tif"
'                MsgBox "strPath = " & strPath & vbCrLf & "strRepairPath = " & strRepairPath
                Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                If Not (pLyr Is Nothing) Then
                    ODNR_Common.CollapseLegend pLyr
                    'Added so as to set the transparency for the oil and gas well maps, Jim McDonald, 20060520
                    If (m_pParent.ProjectType = odnrOGWells Or m_pParent.ProjectType = odnrAUM) Then
                        Set pLEffects = pLyr
                        pLEffects.Transparency = 50
                    End If
                    pGlyr.Add pLyr
                End If
                Set pQuad = m_pParent.Quads.NextQuad
            Loop
            If Not (m_pParent.Quads.FocusQuad Is Nothing) Then
                m_pParent.Quads.ResetAdjacentQuads
                Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Do While Not pQuad Is Nothing
                    strPath = g_strDRG_Path & AddUnderscores(pQuad.QuadName) & ".lyr"
                    strRepairPath = g_strDRGLOC_Path '& AddUnderscores(pQuad.QuadName) & ".tif"
'                    MsgBox "strPath = " & strPath & vbCrLf & "strRepairPath = " & strRepairPath
                    Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                    If Not (pLyr Is Nothing) Then
                        ODNR_Common.CollapseLegend pLyr
                        'Added so as to set the transparency for the oil and gas well maps, Jim McDonald, 20060520
                        If (m_pParent.ProjectType = odnrOGWells Or m_pParent.ProjectType = odnrAUM) Then
                            Set pLEffects = pLyr
                            pLEffects.Transparency = 50
                        End If
                        pGlyr.Add pLyr
                    End If
                    Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Loop
            End If
            Set pQuadlyr = New ODNRQuadLayer
            pGlyr.Expanded = False
            If (m_pParent.ProjectType = odnrBedrockStructure) Then
                pGlyr.Visible = False
            Else
                pGlyr.Visible = True
            End If
            pQuadlyr.LoadFromLayer m_pApp, Me, pGlyr, esriDTRasterDataset, odnrZoomLevelQuad, odnrGeologyMap, True
            m_pQuadLayers.Add Key:=pGlyr.Name, Item:=pQuadlyr
            ODNR_Common.ShowMessage "", False
            
            'Handle the Bedrock Structure Scans here
            If (m_pParent.ProjectType = odnrBedrockStructure Or m_pParent.ProjectType = odnrGeology) Then
                ODNR_Common.ShowMessage "Loading Bedrock Structure Scans...", False
                Set pGlyr = FindLayer("Bedrock Structure Scans")
                If (pGlyr Is Nothing) Then
                    Set pGlyr = New GroupLayer
                    pGlyr.Name = "Bedrock Structure Scans"
                Else
                    ClearGroupLayer pGlyr
                End If
                Set pWsFact = New RasterWorkspaceFactory
                Set pRWs = pWsFact.OpenFromFile(g_strSCAN_Path, 0)
                m_pParent.Quads.Reset
                Set pQuad = m_pParent.Quads.NextQuad
                Do While Not pQuad Is Nothing
                    strPath = g_strSCAN_Path & AddUnderscores(pQuad.QuadName) & "*.tif"
                    strName = Dir$(strPath)
                    Do While strName <> ""
                        Set pRDs = pRWs.OpenRasterDataset(strName)
                        Set pRLyr = New RasterLayer
                        pRLyr.CreateFromDataset pRDs
                        If Not (pRLyr Is Nothing) Then
                            ODNR_Common.CollapseLegend pRLyr
                            pGlyr.Add pRLyr
                        End If
                        strName = Dir
                    Loop
                    Set pQuad = m_pParent.Quads.NextQuad
                Loop
                Set pQuadlyr = New ODNRQuadLayer
                pGlyr.Expanded = False
                pGlyr.Visible = False
                'Changed 20051006, James McDonald
                'Trying to figure out how not to export the Structure scans.  Have changed the CALL procedure from
                'TRUE to FALSE
                'pQuadLyr.LoadFromLayer m_pApp, Me, pGlyr, esriDTRasterDataset, odnrZoomLevelQuad, odnrGeologyMap, True
                pQuadlyr.LoadFromLayer m_pApp, Me, pGlyr, esriDTRasterDataset, odnrZoomLevelQuad, odnrGeologyMap, False
                m_pQuadLayers.Add Key:=pGlyr.Name, Item:=pQuadlyr
                ODNR_Common.ShowMessage "", False
            End If
            
        ElseIf (m_pParent.QuadScale = odnr100K) Then
            'Add the DRG's into the Map
            ODNR_Common.ShowMessage "Loading DRGs...", False
            Set pGlyr = FindLayer("DRGs")
            If (pGlyr Is Nothing) Then
                Set pGlyr = New GroupLayer
                pGlyr.Name = "DRGs"
            Else
                ClearGroupLayer pGlyr
            End If
            m_pParent.Quads.Reset
            Set pQuad = m_pParent.Quads.NextQuad
            Do While Not pQuad Is Nothing
                strPath = g_strDRG100KLyr_Path & AddUnderscores(pQuad.QuadName) & ".lyr"
                strRepairPath = g_strDRG100K_Path '& AddUnderscores(pQuad.QuadName) & ".tif"
'                MsgBox "strPath = " & strPath & vbCrLf & "strRepairPath = " & strRepairPath
                Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                If Not (pLyr Is Nothing) Then
                    ODNR_Common.CollapseLegend pLyr
                    pGlyr.Add pLyr
                End If
                Set pQuad = m_pParent.Quads.NextQuad
            Loop
            If Not (m_pParent.Quads.FocusQuad Is Nothing) Then
                m_pParent.Quads.ResetAdjacentQuads
                Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Do While Not pQuad Is Nothing
                    strPath = g_strDRG100KLyr_Path & AddUnderscores(pQuad.QuadName) & ".lyr"
                    strRepairPath = g_strDRG100K_Path '& AddUnderscores(pQuad.QuadName) & ".tif"
'                    MsgBox "strPath = " & strPath & vbCrLf & "strRepairPath = " & strRepairPath
                    Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                    If Not (pLyr Is Nothing) Then
                        ODNR_Common.CollapseLegend pLyr
                        pGlyr.Add pLyr
                    End If
                    Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Loop
            End If
            'Add in the BT Hillshade
            m_pParent.Quads.Reset
            Set pQuad = m_pParent.Quads.NextQuad
            Do While Not pQuad Is Nothing
                strPath = g_strDRG100KLyr_Path & "bt" & AddUnderscores(pQuad.QuadName) & "_hs.lyr"
                strRepairPath = g_strDRG100K_Path '& "bt" & AddUnderscores(pQuad.QuadName) & "_hs.tif"
                Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                If Not (pLyr Is Nothing) Then
                    ODNR_Common.CollapseLegend pLyr
                    pGlyr.Add pLyr
                End If
                Set pQuad = m_pParent.Quads.NextQuad
            Loop
            If Not (m_pParent.Quads.FocusQuad Is Nothing) Then
                m_pParent.Quads.ResetAdjacentQuads
                Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Do While Not pQuad Is Nothing
                    strPath = g_strDRG100KLyr_Path & "bt" & AddUnderscores(pQuad.QuadName) & "_hs.lyr"
                    strRepairPath = g_strDRG100K_Path '& "bt" & AddUnderscores(pQuad.QuadName) & "_hs.tif"
                    Set pLyr = ODNR_Common.LoadFromGxFile(strPath, ODNRGeologyWorkspace, strRepairPath)
                    If Not (pLyr Is Nothing) Then
                        ODNR_Common.CollapseLegend pLyr
                        pGlyr.Add pLyr
                    End If
                    Set pQuad = m_pParent.Quads.NextAdjacentQuad
                Loop
            End If
            'End of adding in the BT Hillshade
            Set pQuadlyr = New ODNRQuadLayer
            pGlyr.Expanded = False
            If (m_pParent.ProjectType = odnrBedrockStructure) Then
                pGlyr.Visible = False
            Else
                pGlyr.Visible = True
            End If
            pQuadlyr.LoadFromLayer m_pApp, Me, pGlyr, esriDTRasterDataset, odnrZoomLevelQuad, odnrGeologyMap, True
            m_pQuadLayers.Add Key:=pGlyr.Name, Item:=pQuadlyr
            ODNR_Common.ShowMessage "", False
            'End of adding the 100K DRGs
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuadLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
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

Function AddUnderscores(strQuadName As String) As String
'------------------------------------------------------------------------------
'Title: AddUnderscores
'Date: 20060107
'Version: 1.0
'Abstract:  This function will add underscores to the Quadrangle names
'------------------------------------------------------------------------------
'James McDonald
'GIMS Specialist
'Ohio Division of Geological Survey
'2045 Morse Road
'Columbus, OH  43229-6693
'Ph. (614) 265-6601
'Email: jim.mcdonald@dnr.state.oh.us
'------------------------------------------------------------------------------
On Error GoTo ErrorHandler
    Dim intLength As Integer
    Dim i As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim strNewQuadName As String
    
    intLength = Len(strQuadName)
    intStart = 1
    intEnd = intLength
    For i = 1 To intLength
        If (Mid(strQuadName, i, 1) = " ") Then
            intEnd = i
            strNewQuadName = strNewQuadName & Mid(strQuadName, intStart, intEnd - intStart) & "_"
            intStart = intEnd + 1
        ElseIf (Mid(strQuadName, i, 1) = ".") Then
            intEnd = i
            strNewQuadName = strNewQuadName & Mid(strQuadName, intStart, intEnd - intStart)
            intStart = intEnd + 1
        End If
    Next
    strNewQuadName = strNewQuadName & Mid(strQuadName, intStart, intLength - intStart + 1)
    
    AddUnderscores = strNewQuadName
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".AddUnderscores " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
