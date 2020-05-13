VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Select_DataDir_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
'   Program:    Select_DataDir_Cmd
'   Author:     Gregory Palovchik, Taratec Corporation
'   Date:       June 4, 2004
'   Purpose:    Opens a form to select the data directory for the
'               tools
'   Called from:
'
'*****************************************
Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pBitmap1 As IPictureDisp   'Bitmaps for the commands
Private m_blnOpeningDoc As Boolean
Private WithEvents m_pDocEvents As DocumentEvents
Attribute m_pDocEvents.VB_VarHelpID = -1
Private WithEvents m_pViewEvents As Map
Attribute m_pViewEvents.VB_VarHelpID = -1

Implements ICommand
' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "Select_DataDir_Cmd"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler

    Set m_pBitmap1 = ODNR_Common.ButtonImage("Fileopen")
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler
    
    If Not (gODNRProject Is Nothing) Then
        Set gODNRProject = Nothing
    End If
    Set m_pApp = Nothing

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ICommand_Enabled = ODNR_Common.ControlEnabled("Select_DataDir_Cmd")
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Enabled " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Checked " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler
    
    ICommand_Name = "ODNR_Geo1.Select_DataDir_Cmd"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler
    
    ICommand_Caption = "ODNR Project Directories"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ICommand_Tooltip = "ODNR Project Directories"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_Tooltip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ICommand_Message = "Click to open the data directory form"
    
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
    
    Dim pMxDoc As IMxDocument
    Set m_pApp = hook
    ODNR_Common.HookApplication m_pApp
    LoadProjectsDatabase
    Set pMxDoc = m_pApp.Document
    Set m_pDocEvents = pMxDoc
    Set m_pViewEvents = pMxDoc.FocusMap
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnCreate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
    On Error GoTo ErrorHandler
    
    Load frmDirectoryPath
    Set frmDirectoryPath.ESRIApplication = m_pApp
    frmDirectoryPath.Show vbModal
  
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ICommand_OnClick " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function m_pDocEvents_BeforeCloseDocument() As Boolean
    On Error GoTo ErrorHandler
    
    g_blnMapsChanging = True
    m_blnOpeningDoc = True
    If Not (gODNRProject Is Nothing) Then
        Dim pMxDoc As IMxDocument, pDocDirty As IDocumentDirty2
        Set pMxDoc = m_pApp.Document
        Set pDocDirty = pMxDoc
        pDocDirty.SetClean
    End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".m_pDocEvents_BeforeCloseDocument " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub m_pDocEvents_CloseDocument()
    On Error GoTo ErrorHandler
    
    If Not (gODNRProject Is Nothing) Then
        gODNRProjectDb.CloseProject
    End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".m_pDocEvents_CloseDocument " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pDocEvents_OpenDocument()
    On Error GoTo ErrorHandler
    
    Dim pMxDoc As IMxDocument
    Set pMxDoc = m_pApp.Document
    Set m_pViewEvents = pMxDoc.PageLayout
    m_blnOpeningDoc = False
    If Not (gODNRProjectDb Is Nothing) Then
        gODNRProjectDb.LoadProject
    End If
    ODNR_Common.UpdateProjectCombo
 
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".m_pDocEvents_OpenDocument " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pViewEvents_AfterDraw(ByVal Display As esriDisplay.IDisplay, ByVal phase As esriCarto.esriViewDrawPhase)
    On Error GoTo ErrorHandler
    
    If (m_blnOpeningDoc = False) And (phase = esriViewGeography) Then
        g_blnMapsChanging = False
    End If
 
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".m_pDocEvents_OpenDocument " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

