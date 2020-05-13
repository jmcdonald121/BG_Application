VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRPackager
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates a zip file of the information in the
'               specified directory.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pZipUp As ZipUp.CZipUp
Private m_blnIsloaded As Boolean

Private WithEvents m_pZipEvents As ZipUp.CZipUp
Attribute m_pZipEvents.VB_VarHelpID = -1

Const c_strModuleName As String = "ODNRPackager"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set m_pZipUp = New ZipUp.CZipUp
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pZipUp = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function Package(strExportDir As String, strZipFileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strCheckPath As String, strFilePath As String
    m_pZipUp.Quiet = True
    m_pZipUp.RootDir = strExportDir

    m_pZipUp.ZipfileName = strZipFileName

    m_pZipUp.AddFile "*.*"
 
    Dim lngVal As Long
    lngVal = m_pZipUp.Zip
    m_pZipUp.Reset
    m_pZipUp.ClearFiles
    If (lngVal = 0) Then
        Package = True
    Else
        Package = False
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".Package " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub m_pZipEvents_FileComplete(ByVal CompleteText As String)
'
End Sub

Private Sub m_pZipEvents_Message(ByVal MessageText As String)
'
End Sub
