VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Select100_Tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
'   Program:    Select100_Tool
'   Author:     Jeffrey M Laird
'   Date:       March 25, 2002
'   Purpose:    Creates the Select100K Quad Area Tool on the Toolbar.
'               Enables the user to pick a quad sized area from the view
'   Called from:
'
'*****************************************

Option Explicit

Implements ICommand
Implements ITool

Private m_pApp As esriFramework.IApplication
Private m_pBitmap1 As IPictureDisp
Private m_pMxDoc As IMxDocument
Private m_pActiveView As IActiveView
Private m_pScreenDisplay As IScreenDisplay
Private m_pHitElem As IElement
Private m_pDispFeed As IDisplayFeedback
' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "Select100_Tool"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set m_pBitmap1 = ODNR_Common.ButtonImage("Sel100")
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pApp = Nothing
    Set m_pBitmap1 = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    
    '**********************
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnMouseUp " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_Select()
    On Error GoTo ErrorHandler
    
    'Get a reference to the ActiveView and ScreenDisplay
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Select " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Enabled = ControlEnabled("Select100_Tool")
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Enabled " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Checked = False
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Checked " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler
    
    ICommand_Name = "ODNR_Geo1.Select100_Tool"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler
    
    ICommand_Caption = "Select 30 X 60 Area"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ICommand_Tooltip = "Select 30 X 60 Area"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Tooltip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ICommand_Message = "Select 100K Area, Double click to generate new selection Poly."
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Message " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_HelpFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_HelpContextID " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    On Error GoTo ErrorHandler
    
    ICommand_Bitmap = m_pBitmap1
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Bitmap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
    On Error GoTo ErrorHandler
    
    ICommand_Category = "ODNRGeo_MiscTools"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Category " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
    On Error GoTo ErrorHandler
    
    Set m_pApp = hook
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnCreate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
    On Error GoTo ErrorHandler
    
    'this checks to see if a select polygon exists, if not it adds one in the middle
    Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
    Dim pQf As IQueryFilter, pF As IFeature, pFCursor As IFeatureCursor
    Dim pArea As IArea, pCtrPt As IPoint
    
    Set m_pActiveView = gODNRProject.ProjectMap(odnrGeologyMap)
    Set m_pScreenDisplay = m_pActiveView.ScreenDisplay
    ' Check that we have some selected elements
    If (m_pHitElem Is Nothing) Then
        Set pFlyr = gODNRProject.QuadFeatureLayer
        pFlyr.ShowTips = False
        Set pFc = pFlyr.FeatureClass
        Set pQf = New QueryFilter
        pQf.WhereClause = "NAME = 'Newark'"
        Set pFCursor = pFc.Search(pQf, False)
        Set pF = pFCursor.NextFeature
        
        Set m_pHitElem = New PolygonElement
        m_pHitElem.Geometry = pF.Shape
    '        m_pGraCont.AddElement m_pHitElem, 0
    
        ' Create a MovePolygonFeedback object and set its display property (to the ActiveView's ScreenDisplay)
        Set m_pDispFeed = New MovePolygonFeedback
        Set m_pDispFeed.Display = m_pScreenDisplay
        ' QI for the IMovePolygonFeedback interface
        Dim pMvPolyFeed As IMovePolygonFeedback
        Set pMvPolyFeed = m_pDispFeed
        'Start the feedback using the input (Polygon) geometry at the current mouse location
        Set pArea = m_pHitElem.Geometry
        Set pCtrPt = pArea.Centroid
        pMvPolyFeed.Start m_pHitElem.Geometry, pCtrPt
        m_pActiveView.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    On Error GoTo ErrorHandler
    
    'ITool_Cursor = m_pCursor1
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Cursor " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    
    '
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnMouseDown " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    
    If Not m_pDispFeed Is Nothing Then
        Dim pPt As IPoint
        ' Get the current mouse location in Map Units and move the feedback
        Set pPt = m_pScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
        m_pDispFeed.MoveTo pPt
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnMouseMove " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnDblClick()
    On Error GoTo ErrorHandler
    
    Dim pPolygon As IPolygon
    Dim pMvPolyFeed As IMovePolygonFeedback
    Dim pUID As New UID
    Dim pCmdItem As ICommandItem
    If Not (m_pHitElem Is Nothing) Then
        Set pMvPolyFeed = m_pDispFeed
        Set pPolygon = pMvPolyFeed.Stop
        Set m_pDispFeed = Nothing
        gODNRProject.Quads.AddQuadsByGeometry pPolygon.Envelope
        Set gODNRProject.Quads.ExtentEnvelope = pPolygon.Envelope
        gODNRProject.Quads.NoFocusQuad
        gODNRProject.ZoomToQuadsExtent
        gODNRProject.ShowQuadLayers
        'Zoom to quads here
        ' Use the GUID of the Save command
        pUID.Value = "{C22579D1-BC17-11D0-8667-0000F8751720}" 'Select_Elements tool
        Set pCmdItem = m_pApp.Document.CommandBars.Find(pUID)
        pCmdItem.Execute
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnDblClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnKeyDown " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyUp(ByVal keyCode As Long, ByVal Shift As Long)
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnKeyUp " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnContextMenu " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 
Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)
    On Error GoTo ErrorHandler
    
    If Not (m_pDispFeed Is Nothing) Then
        m_pDispFeed.Refresh hDC
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Refresh " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Function ITool_Deactivate() As Boolean
    On Error GoTo ErrorHandler
    
    If Not (m_pHitElem Is Nothing) Then
        'm_pGraCont.DeleteElement m_pHitElem
        Set m_pHitElem = Nothing
        'Set m_pGraCont = Nothing
        m_pActiveView.Refresh
        gODNRProject.QuadFeatureLayer.ShowTips = True
        Set m_pActiveView = Nothing
    End If
    ITool_Deactivate = True
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Deactivate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

