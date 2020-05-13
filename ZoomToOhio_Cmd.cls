VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZoomToOhio_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************
'
'   Program:    ZoomToOhio_CMD
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates the zoom to Ohio Button on the Toolbar.
'               Changes zoom level to extent of Ohio
'   Called from:
'
'*****************************************
Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pBitmap1 As IPictureDisp   'Bitmaps for the commands

Implements ICommand
' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "ZoomToOhio_Cmd"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set m_pBitmap1 = ODNR_Common.ButtonImage("Ohio")
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pApp = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Enabled = ControlEnabled("ZoomToOhio_Cmd")
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Enabled " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
    On Error GoTo ErrorHandler
    
    'TODO: Add your implementation here
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Checked " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler
    
    ICommand_Name = "OBDR_Geo1.ZoomToOhio_Cmd"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler
    
    ICommand_Caption = "Zoom to Ohio"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ICommand_Tooltip = "Zoom to Ohio"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Tooltip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ICommand_Message = "Zoom to Ohio"
    
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
    
    If Not (gODNRProject Is Nothing) Then
        gODNRProject.GoDataView
        gODNRProject.ResetLayout
        gODNRProject.ResetDataView
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

