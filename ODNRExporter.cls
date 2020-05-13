VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_blnIsLoaded As Boolean

Const c_strModuleName As String = "ODNRExporter"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub Load(strRootDir As String, strZipFileName As String, blnDeleteFiles As Boolean)

End Sub


Private Function PackageExportFiles(strExportDir As String, strZipName As String, pFileList As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pZipUp As ZipUp.CZipUp
    Set pZipUp = New ZipUp.CZipUp
    pZipUp.Quiet = True
    pZipUp.RootDir = strExportDir
    pZipUp.TempDir = strExportDir
    pZipUp.ZipfileName = strZipName
    
    Dim vKey As Variant
    For Each vKey In pFileList
        MsgBox "Dir " & strExportDir & "  File " & CStr(vKey)
        pZipUp.AddFile CStr(vKey)
    Next
    Dim lngVal As Long
    lngVal = pZipUp.Zip
    pZipUp.ClearFiles
    pZipUp.Reset
    Set pZipUp = Nothing
    If (lngVal = 0) Then
        PackageExportFiles = True
    Else
        PackageExportFiles = False
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".PackageExportFiles " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


