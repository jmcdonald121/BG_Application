VERSION 5.00
Begin VB.Form frmGeo1 
   Caption         =   "Select Quad"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboProject 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Change Project"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cboQuad 
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a Quadrangle"
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frmGeo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    frmToolBarControls
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Creates quad selection combo box on the tool bar
'               and populates it with either 24k or 100k quads
'   Called from:
'
'*****************************************
Option Explicit


Private m_pApp As esriCore.IApplication
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "D:\ODNR_Geology\VB_Projects\frmToolBarControls.frm"

Private Sub cboProject_Click()
  On Error GoTo ErrorHandler

    Dim pMxDoc As IMxDocument
    Dim pDoc As IDocument
    Dim Projname As String
    Set pMxDoc = m_pApp.Document
    Set pDoc = m_pApp.Document
    Projname = pDoc.Title
    If Not (gLayoutElements Is Nothing) Then
        gLayoutElements.RemoveAll
        Set gLayoutElements = Nothing
    End If
    Set gLayoutElements = New Dictionary
If LayPop Then 'reset the layout
    Call ResetLayout(Projname, m_pApp)
         g_strSelect_Quad = "x"
End If
         If K100_AdHoc Then
            K100_AdHoc = False
         End If
    
    
    Dim Tstring As String
    Tstring = cboProject.Text & "_Quad.mxd"
    
    If Tstring <> Projname Then
    ' If cboProject.Text = "BG100K" Then 'delete these when the BG projects are up
    '    MsgBox "Not yet implemented"
    '    Exit Sub
    ' End If
    ' If cboProject.Text = "BG24K" Then
    '    MsgBox "Not yet implemented"
     '   Exit Sub
     'End If
        m_pApp.OpenDocument g_strGeodata_Path & Tstring
            docChange = True
    End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "cboProject_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboQuad_Click()
  On Error GoTo ErrorHandler



    Dim pMxDoc As IMxDocument
    Dim pDoc As IDocument
    Dim pMap As IMap
    Dim pEnumFeature As IEnumFeature
    Dim pFeature As IFeature
    Dim pFeatureLayer As IFeatureLayer
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pMapCollection As IMaps
    Dim pMapNDX As Long
    Dim g As Long
    Dim k As Integer
    Dim mapname As String
    Dim Projname As String
    Dim Lstring As String
    Dim Got24 As Boolean
    Dim Got100 As Boolean
    Dim pField As IField
    Dim GotQuads(9) As String
    Dim arraycnt As Integer
    Dim lgotfield As Long
    Dim pquadname As String
    Dim BG As Boolean
    Dim holdmap As String
    BG = False
    
    Got24 = False
    Got100 = False
        
    
'***************** this routine sets the selectability to Quads only
    Set pMxDoc = m_pApp.Document
    Set pDoc = m_pApp.Document
    Projname = pDoc.Title
    
    If LayPop Then 'reset the layout
        Call ResetLayout(Projname, m_pApp)
    End If
         If K100_AdHoc Then
            K100_AdHoc = False
            K100_Quads(0) = ""
            K100_Quads(1) = ""
            K100_Quads(2) = ""
            K100_Quads(3) = ""
            K100_Quads(4) = ""
            Set K100_Geometry = Nothing
         End If
  
      
    Set pMapCollection = pMxDoc.Maps
    pMapNDX = pMapCollection.Count
    If pMapCollection.Count = 0 Then Exit Sub
    
      For g = 0 To pMapNDX - 1
        mapname = pMapCollection.Item(g).Name
        If mapname = "Geology Map24BG" Or mapname = "Geology Map100BG" Then
            BG = True
            holdmap = mapname
        End If

         If mapname = "Geology Map24BT" Or mapname = "Geology Map24BG" Then
            Got24 = True
            Set pMap = pMapCollection.Item(g)
            pMap.ClearSelection
            Set pEnumLayer = pMap.Layers
            pEnumLayer.Reset
            Set pLayer = pEnumLayer.Next
                    Do While Not pLayer Is Nothing
                          If Not TypeOf pLayer Is IRasterLayer Then
                      If pLayer.Name <> "Quad24K" Then
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = False
                        Else
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = True
                      End If
                          End If
                      Set pLayer = pEnumLayer.Next
                    Loop
         ElseIf mapname = "Geology Map100BT" Or mapname = "Geology Map100BG" Then
            Got100 = True
            Set pMap = pMapCollection.Item(g)
            pMap.ClearSelection
            Set pEnumLayer = pMap.Layers
            pEnumLayer.Reset
            Set pLayer = pEnumLayer.Next
                    Do While Not pLayer Is Nothing
                          If Not TypeOf pLayer Is IRasterLayer Then
                      If pLayer.Name <> "Quad100K" Then
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = False
                        Else
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = True
                      End If
                          End If
                      Set pLayer = pEnumLayer.Next
                    Loop
         End If
        Next g
'**************** now select something
Dim pActiveView As IActiveView
Set pActiveView = pMap

   Dim pVal As String
   pVal = frmToolBarControls.cboQuad.Text
   
   Dim pQueryFilter As IQueryFilter
   
   If Got24 Then
         Set pFeatureLayer = FindLayerByName(pMap, "Quad24K")
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "QUADNAME = '" & pVal & "'"
   End If
   If Got100 Then
         Set pFeatureLayer = FindLayerByName(pMap, "Quad100K")
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "NAME = '" & pVal & "'"
   End If
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureLayer.Search(pQueryFilter, False)

    Set pFeature = pFeatureCursor.NextFeature
    If pFeature Is Nothing Then
        MsgBox "Error locating Quadname record.", vbCritical, "Record not found"
        Exit Sub
    End If
    
    Dim Cfeature As IFeature
    Set Cfeature = pFeature  'this saves the Pfeature shape for passing to getcounties
    
    Dim SelCount As Long
    SelCount = pMap.SelectionCount
    If SelCount > 1 Then
        MsgBox "More than one Quad is selected - select only one quad and proceed.", vbCritical, "Procedure Error"
        Exit Sub
    End If
                    
         Dim m_pMxApp As IMxApplication
         Set m_pMxApp = m_pApp
         Dim pDisplayTransform As IDisplayTransformation
         Set pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation
         Dim pEnvelope As IEnvelope
         Set pEnvelope = pDisplayTransform.VisibleBounds
         Set pEnvelope = pFeature.Shape.Envelope
         pEnvelope.Expand 1.01, 1.01, True
         pDisplayTransform.VisibleBounds = pEnvelope
                              
         If Got24 Then
            lgotfield = pFeatureCursor.FindField("QUADNAME")
          Else
            lgotfield = pFeatureCursor.FindField("NAME")
         End If
         If lgotfield = -1 Then
           MsgBox "Field QUADNAME or NAME not found.", vbCritical, "Field not Found"
           Exit Sub
         End If
        
         pquadname = pFeature.Value(lgotfield)
 

'***** select other quads
         pEnvelope.Expand 1.2, 1.2, True
         pFeature.Shape.Envelope.Expand 1.2, 1.2, True
         pMap.SelectByShape pFeature.Shape, m_pMxApp.SelectionEnvironment, False
'***** get quadnames
         Dim pEnumFeatureSetup As IEnumFeatureSetup
         Set pEnumFeature = pMap.FeatureSelection
         Set pEnumFeatureSetup = pEnumFeature
         pEnumFeatureSetup.AllFields = True
         
         
         arraycnt = 0
         
         If Got24 Then
            Set pLayer = FindLayerByName(pMap, "Quad24K")
         ElseIf Got100 Then
            Set pLayer = FindLayerByName(pMap, "Quad100K")
         Else
            MsgBox "Unrecognized Project name - cannot proceed.", vbCritical, "Unrecognized Project"
            Exit Sub
         End If
         
         Set pFeature = pEnumFeature.Next
         Set pFeatureLayer = pLayer
         Set pFeatureCursor = pFeatureLayer.Search(Nothing, False) 'this returns all records
         If Got24 Then
            lgotfield = pFeatureCursor.FindField("QUADNAME")
          Else
            lgotfield = pFeatureCursor.FindField("NAME")
         End If
            
         If lgotfield = -1 Then
           MsgBox "Field QUADNAME or NAME not found.", vbCritical, "Field not Found"
           Exit Sub
         End If
         
         Do While Not pFeature Is Nothing
            GotQuads(arraycnt) = pFeature.Value(lgotfield)
            arraycnt = arraycnt + 1
            Set pFeature = pEnumFeature.Next
         Loop
'***** call getcounty if BG
    If mapname = "Geology Map24BG" Then
        Call GetCounties(Cfeature)
    End If
         
'***** get DRG layers or Quad 100 layers from the arcinfo stuff they havent given us yet

        Dim pGxLayer As esriCore.IGxLayer
        Dim pGxFile As esriCore.IGxFile
        Dim TotString As String
        Dim fullpath As String
        Dim chkpath As String
'***** switchout the BG Units layer here

    'If BG Then
    '    fullpath = BGLegend_Path & pQuadname & ".lyr"

    '    chkpath = Dir(fullpath)

    '    If Not chkpath = "" Then
    '        Set pLayer = FindLayerByName(pMap, "BG Units Ply")
    '        If Not pLayer Is Nothing Then
    '            pMap.DeleteLayer pLayer
    '        End If

     '       Set pGxLayer = New esriCore.GxLayer
     '       Set pGxFile = pGxLayer
     '       pGxFile.Path = fullpath
     '       pGxFile.Open

     '       pMap.AddLayer pGxLayer.Layer
     '       Set pGxFile = Nothing
     '       Set pGxLayer = Nothing
     '       Call ChangeOrder(m_pApp, holdmap)
     '   Else
     '       MsgBox "Error from form based selectquad, cannot locate BG units layer file"

     '   End If
   ' End If
        
        
        
        
        If Got24 Then
         For k = 0 To arraycnt - 1
            Set pGxLayer = New esriCore.GxLayer
            Set pGxFile = pGxLayer
                 TotString = g_strDRG_Path & GotQuads(k) & ".lyr"
            If Dir(TotString) <> "" Then
                pGxFile.Path = g_strDRG_Path & GotQuads(k) & ".lyr"
                pGxFile.Open
                pMap.AddLayer pGxLayer.Layer
            Else
               MsgBox GotQuads(k) & ".lyr" & " for DRG Tif image not found.  Check for correct path and filename and try again.", vbCritical, "File not Found"
            End If
         Next k
        End If
        If Got100 Then
            fullpath = g_strBase_Path & pquadname & "_A.lyr"
            chkpath = Dir(fullpath)
            If Not chkpath = "" Then
                Set pGxLayer = New esriCore.GxLayer
                Set pGxFile = pGxLayer
                pGxFile.Path = g_strBase_Path & pquadname & "_A.lyr"
                pGxFile.Open
                pMap.AddLayer pGxLayer.Layer
            End If
        
            fullpath = g_strBase_Path & pquadname & "_X.lyr"
            chkpath = Dir(fullpath)
            If Not chkpath = "" Then
                Set pGxLayer = New esriCore.GxLayer
                Set pGxFile = pGxLayer
                pGxFile.Path = g_strBase_Path & pquadname & "_X.lyr"
                pGxFile.Open
                pMap.AddLayer pGxLayer.Layer
            End If
        
            fullpath = g_strBase_Path & pquadname & "_P.lyr"
            chkpath = Dir(fullpath)
            If Not chkpath = "" Then
                Set pGxLayer = New esriCore.GxLayer
                Set pGxFile = pGxLayer
                pGxFile.Path = g_strBase_Path & pquadname & "_P.lyr"
                pGxFile.Open
                pMap.AddLayer pGxLayer.Layer
            End If
        
            fullpath = g_strBase_Path & pquadname & "_L.lyr"
            chkpath = Dir(fullpath)
            If Not chkpath = "" Then
                Set pGxLayer = New esriCore.GxLayer
                Set pGxFile = pGxLayer
                pGxFile.Path = g_strBase_Path & pquadname & "_L.lyr"
                pGxFile.Open
                pMap.AddLayer pGxLayer.Layer
            End If
        End If
        
        
        Set pGxFile = Nothing
        Set pGxLayer = Nothing
        Set pMap = pMxDoc.FocusMap
        pMap.ClearSelection

        Set pMxDoc.ActiveView = pMap
        pActiveView.Refresh
         
' Thats it - Call the resequencer
     g_strSelect_Quad = pquadname
     If BG Then
        Call switchout(pquadname, pMap, Got24, m_pApp)
     End If
        If Got24 Then
     Call addrstr(pquadname)
        End If
     Call CloseLayers(pMxDoc)
    If BG Then
        Call ChangeOrder(m_pApp, holdmap)
     Else
        Call ChangeOrder(m_pApp, mapname)
    End If
     Call OverviewMapQuad(Projname, pquadname, m_pApp)
    If BG Then
        Call DoRender(m_pApp)
    End If
     

'End If






  Exit Sub
ErrorHandler:
  HandleError True, "cboQuad_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cboQuad_GotFocus()
 If docChange Then
    Call PopCbo(m_pApp)
    docChange = False
 End If
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler


 Call PopCbo(m_pApp)

   frmToolBarControls.cboProject.AddItem "BT24K"
   frmToolBarControls.cboProject.AddItem "BT100K"
   frmToolBarControls.cboProject.AddItem "BG24K"
   frmToolBarControls.cboProject.AddItem "BG100K"






  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Terminate()
  On Error GoTo ErrorHandler




     Set m_pApp = Nothing





  Exit Sub
ErrorHandler:
  HandleError True, "Form_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Public Property Let Application(ByVal vNewVal As IApplication)
  On Error GoTo ErrorHandler





  Set m_pApp = vNewVal
  








  Exit Property
ErrorHandler:
  HandleError True, "Application " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Sub addrstr(sFilename As String)
  On Error GoTo ErrorHandler


'********************************************************************
'Adding Raster File
'********************************************************************

'********************************************************************
'   Create RasterWorkSpaceFactory
'********************************************************************

    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pDoc As IMxDocument
    Set pDoc = m_pApp.Document
    Dim isPath As String
    Dim aVal
    Dim drgEnvelope As IEnvelope
    Dim spath As String
    Dim slyr As String
    Dim stype As String
    Dim fullfile As String
    Dim myName As String

'********************************************************************
'   Get RasterWorkspace
'   Check if the path found in the system database is valid
'********************************************************************

    Dim pRasWS As IRasterWorkspace
    Dim pRLyr As IRasterLayer
    Dim pRasDS As IRasterDataset
    
    spath = g_strSCAN_Path & sFilename & " *.tif"
    myName = Dir$(spath) ', vbDirectory)   ' Retrieve the first entry.
    Do While myName <> ""   ' Start the loop.
            Set pRasWS = pWSF.OpenFromFile(g_strSCAN_Path, 0)
            Set pRasDS = pRasWS.OpenRasterDataset(myName)
            Set pRLyr = New RasterLayer
            pRLyr.CreateFromDataset pRasDS
            pDoc.AddLayer pRLyr
                 pRLyr.Visible = False
   myName = Dir   ' Get next entry.
   Loop




  Exit Sub
ErrorHandler:
  HandleError True, "addrstr " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub GetCounties(Qfeature As IFeature)
    Dim pMxDoc As IMxDocument
    'Dim pDoc As IDocument
    Dim pMap As IMap
    Dim pEnumFeature As IEnumFeature
    Dim pFeature As IFeature
    Dim pFeatureLayer As IFeatureLayer
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pMapCollection As IMaps
    Dim pMapNDX As Long
    Dim g As Long
    Dim k As Integer
    Dim mapname As String
    'Dim Projname As String
    Dim Lstring As String
    Dim pField As IField
    Dim GotCounties(7) As String
    Dim arraycnt As Integer
    Dim lgotfield As Long
    Dim pFeatureCursor As IFeatureCursor
    'Dim fullpath As String
    
'***************** this routine sets the selectability to Counties only
'****** counties must be on to select

    Set pMxDoc = m_pApp.Document
    Dim m_pMxApp As IMxApplication
    Set m_pMxApp = m_pApp
    'Set pDoc = m_pApp.Document
    'Projname = pDoc.Title
          
    Set pMapCollection = pMxDoc.Maps
    pMapNDX = pMapCollection.Count
    Dim found_cty As Boolean
    found_cty = False
    
      For g = 0 To pMapNDX - 1
        mapname = pMapCollection.Item(g).Name
         If mapname = "Geology Map24BG" Then
            Set pMap = pMapCollection.Item(g)
            pMap.ClearSelection
            Set pEnumLayer = pMap.Layers
            pEnumLayer.Reset
            Set pLayer = pEnumLayer.Next
                    Do While Not pLayer Is Nothing
                          If Not TypeOf pLayer Is IRasterLayer Then
                            If pLayer.Name <> "County_Bndry" Then
                                Set pFeatureLayer = pLayer
                                pFeatureLayer.Selectable = False
                            Else
                                Set pFeatureLayer = pLayer
                                pFeatureLayer.Visible = True
                                pFeatureLayer.Selectable = True
                                found_cty = True
                            End If
                          End If
                      Set pLayer = pEnumLayer.Next
                    Loop
         End If
        Next g
    
    If Not found_cty Then
        g_strCounty_Names = "NONAME"
        Exit Sub
    End If
    
'**************** now select something
Dim pActiveView As IActiveView
Set pActiveView = pMap

         
         Dim pDisplayTransform As IDisplayTransformation
         Set pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation
         Dim pEnvelope As IEnvelope
         Set pEnvelope = pDisplayTransform.VisibleBounds
         Set pEnvelope = Qfeature.Shape.Envelope
         pDisplayTransform.VisibleBounds = pEnvelope
                          
'***** select other counties

         pMap.SelectByShape Qfeature.Shape, m_pMxApp.SelectionEnvironment, False
         
'***** get quadnames
    Dim pEnumFeatureSetup As IEnumFeatureSetup
   ' Dim pQuadname As String
    Set pEnumFeature = pMap.FeatureSelection
    Set pEnumFeatureSetup = pEnumFeature
    pEnumFeatureSetup.AllFields = True
         
         
         Set pEnumFeature = pMap.FeatureSelection
         Set pEnumFeatureSetup = pEnumFeature
         pEnumFeatureSetup.AllFields = True
                  
         arraycnt = 0
         
         Set pLayer = FindLayerByName(pMap, "County_Bndry")
         
         Set pFeature = pEnumFeature.Next
         Set pFeatureLayer = pLayer
         Set pFeatureCursor = pFeatureLayer.Search(Nothing, False) 'this returns all records
            
         lgotfield = pFeatureCursor.FindField("COUNTY")
            
         If lgotfield = -1 Then
           MsgBox "Field COUNTY not found - Cannot continue.", vbCritical, "Field not Found"
           g_strCounty_Names = "NONAME"
           Exit Sub
         End If
         
         For k = 0 To 6  'spike the array
            GotCounties(k) = "X"
         Next k
         
         Do While Not pFeature Is Nothing
            GotCounties(arraycnt) = pFeature.Value(lgotfield)
            arraycnt = arraycnt + 1
            Set pFeature = pEnumFeature.Next
         Loop
         
Dim Cstring As String
If GotCounties(0) <> "X" Then
    Cstring = GotCounties(0)
        For k = 1 To 6
            If GotCounties(k) <> "X" Then
               If GotCounties(k + 1) <> "X" Then
                  Cstring = Cstring & ", " & GotCounties(k)
                Else
                  Cstring = Cstring & " and " & GotCounties(k)
            End If
             Else
                Exit For
            End If
        Next k
  Else
    MsgBox "County Name array position 0 has no value - Check county table and rerun"
    g_strCounty_Names = "NONAME"
End If
If GotCounties(1) = "X" Then
    g_strCounty_Names = UCase$(Cstring) & " COUNTY, OHIO"
  Else
    g_strCounty_Names = UCase$(Cstring) & " COUNTIES, OHIO"
End If
        

        
End Sub


