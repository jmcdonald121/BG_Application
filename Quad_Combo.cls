VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Quad_Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
'   Program:    Quad_Combo
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Creates the Quadrangle selection combo on frmToolBarControls
'   Called from:
'
'*****************************************

Option Explicit

Implements ICommand
Implements IToolControl
Dim m_pApp As esriFramework.IApplication      'ArcMap application
' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "Quad_Combo"

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler

    Set m_pApp = Nothing

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    On Error GoTo ErrorHandler

    'ToDo
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Bitmap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler

    ICommand_Caption = "Select a Quadrangle"

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Category() As String
    On Error GoTo ErrorHandler
    
    ICommand_Category = "ODNRGeo_MiscTools"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Category " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Checked() As Boolean
    On Error GoTo ErrorHandler
    
    'Todo
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Checked " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Enabled = ControlEnabled("Quad_Combo")
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Enabled " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpContextID() As Long
    On Error GoTo ErrorHandler
    
    'ToDo
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_HelpContextID " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpFile() As String
    On Error GoTo ErrorHandler

    'ToDo

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_HelpFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ICommand_Message = "Select a Quadrangle"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Message " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler
    
    ICommand_Name = "ODNR_Geo1.Quad_Combo"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub ICommand_OnClick()
    On Error GoTo ErrorHandler
    
    'ToDo
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    On Error GoTo ErrorHandler
    
    Set m_pApp = hook
      
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnCreate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ICommand_Tooltip = "Select a Quadrangle"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Tooltip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get IToolControl_hWnd() As esriSystem.OLE_HANDLE
    On Error GoTo ErrorHandler
    
    'pass back Imagecombo1 window handle when queried by ArcMap
    IToolControl_hWnd = frmToolbarControls.cboQuad.hwnd
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolControl_hWnd " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Function IToolControl_OnDrop(ByVal barType As esriSystemUI.esriCmdBarType) As Boolean
    On Error GoTo ErrorHandler
    
    'only drop onto ArcMap ToolBar
    If barType = esriCmdBarTypeToolbar Then
        IToolControl_OnDrop = True
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolControl_OnDrop " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub IToolControl_OnFocus(ByVal complete As esriSystemUI.ICompletionNotify)
     On Error GoTo ErrorHandler
    
    'ToDo
    
     Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolControl_OnFocus " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
