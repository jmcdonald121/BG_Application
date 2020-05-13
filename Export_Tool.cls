VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Export_Tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
'   Program:    Export_Tool
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Exports certain layers in the selected Quad.
'               Based on project name and type
'   Called from:
'
'*****************************************


Option Explicit


Implements esriSystemUI.ICommand
Implements esriSystemUI.ITool
Private m_pApp As esriFramework.IApplication
Private m_pBitmap1 As IPictureDisp
Private m_pCursor1 As IPictureDisp

' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "Export_Tool"

' Variables used by the Error handler function - DO NOT REMOVE
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set m_pCursor1 = LoadResPicture("CrossR", 2)
    Set m_pBitmap1 = ODNR_Common.ButtonImage("DatabaseExport")
    
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

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Enabled = ControlEnabled("Export_Tool")
    
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
    
    ICommand_Name = "ODNR_Geo1.Export_Tool"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler
    
    ICommand_Caption = "Pick Quadrangle"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ICommand_Tooltip = "Export Data to Shapefile"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Tooltip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ICommand_Message = "Click inside a quadrangle polygon to export data to shapefile."
    
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
    
    Load frmExportQuads
    Set frmExportQuads.App = m_pApp
    SetWindowLong frmExportQuads.hwnd, GWL_HWNDPARENT, m_pApp.hwnd
    frmExportQuads.Show vbModeless
    If (gODNRProject.Quads.QuadCount > 0) Then
        Dim pQuadList As Collection, pQuad As ODNRQuad
        Set pQuadList = New Collection
        If Not (gODNRProject.Quads.FocusQuad Is Nothing) Then
            Set pQuad = gODNRProject.Quads.FocusQuad
            gODNRProject.Quads.FocusQuad.HighlightQuad
            pQuadList.Add Item:=pQuad.QuadId
        Else
            gODNRProject.Quads.Reset
            Set pQuad = gODNRProject.Quads.NextQuad
            gODNRProject.Quads.HighlightQuads
            Do While Not pQuad Is Nothing
                pQuadList.Add Item:=pQuad.QuadId
                Set pQuad = gODNRProject.Quads.NextQuad
            Loop
        End If
        frmExportQuads.SelectedQuads pQuadList, 1
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    On Error GoTo ErrorHandler
    
    ITool_Cursor = m_pCursor1
    
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
    
    '
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnMouseMove " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    
    If Button = 1 Then
        Dim pPoint As IPoint, pMxDoc As IMxDocument
        Dim pQuadList As Collection
        Set pMxDoc = m_pApp.Document
        Set pPoint = pMxDoc.ActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
        Set pQuadList = gODNRProject.Quads.FindQuadIdsByGeometry(pPoint)
        frmExportQuads.SelectedQuads pQuadList, CInt(Shift)
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_OnMouseUp " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnDblClick()
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
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
    
    ' TODO: Add your implementation here
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Refresh " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Function ITool_Deactivate() As Boolean
    On Error GoTo ErrorHandler
    
    If Not (frmExportQuads Is Nothing) Then
        frmExportQuads.Hide
        Unload frmExportQuads
        Set frmExportQuads = Nothing
    End If
    ITool_Deactivate = True
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ITool_Deactivate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

