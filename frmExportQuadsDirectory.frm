VERSION 5.00
Begin VB.Form frmExportQuadsDirectory 
   Caption         =   "Export Quads Target Directory"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2430
      TabIndex        =   5
      Top             =   4140
      Width           =   1005
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Top             =   4140
      Width           =   1005
   End
   Begin VB.DirListBox ExportDirListBox 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   30
      TabIndex        =   1
      Top             =   870
      Width           =   3405
   End
   Begin VB.DriveListBox DriveSelectionCombo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   3400
   End
   Begin VB.Label InstructionsLabel 
      Caption         =   "Look in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Please identify the export directory below."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   630
      Width           =   3255
   End
End
Attribute VB_Name = "frmExportQuadsDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    frmExportQuadsDirectory
'   Author:     Gregory Palovchik
'               Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229 ext. 202
'   Date:       June 28, 2004
'   Purpose:    Provide a form for selecting an export directory
'
'   Called from: frmExportQuads
'
'*****************************************

Private pFSO As FileSystemObject
Private m_strExport_Path As String
Private m_strCurrentDrive As String

' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "frmExportQuadsDirectory"

Public Property Let ExportPath(RHS As String)
'Set the exportpath from outside the form
    On Error GoTo ErrorHandler

    If (RHS <> "") Then
        If (Right(RHS, 1) = "\") Then RHS = Left(RHS, Len(RHS) - 1)
        If (pFSO.FolderExists(RHS)) Then
            m_strExport_Path = RHS
            strCurrentDrive = pFSO.GetFolder(m_strExport_Path).Drive.Path
            ExportDirListBox.Path = m_strExport_Path
            ExportDirListBox.Refresh
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportPath " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub CancelButton_Click()
'Close and unload the form
    On Error GoTo ErrorHandler

    Me.Hide
    Unload Me
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CancelButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
'Load the form and set the current drive to C:
    On Error GoTo ErrorHandler

    Set pFSO = New FileSystemObject
    strCurrentDrive = "C:"
    m_strExport_Path = strCurrentDrive
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub DriveSelectionCombo_Change()
'Change the current drive
    On Error GoTo ErrorHandler

    Dim strDrive As String, blnFailed As Boolean
    blnFailed = True
    strDrive = Left(DriveSelectionCombo.Drive, 2)
    MsgBox strCurrentDrive
    If (pFSO.DriveExists(strDrive)) Then
    MsgBox strCurrentDrive
        If (pFSO.Drives(strDrive).IsReady) Then
        MsgBox strCurrentDrive
            strCurrentDrive = Strings.UCase(strDrive)
            DataDirListBox.Path = strCurrentDrive
            DataDirListBox.Refresh
            blnFailed = False
        End If
    End If
    If (blnFailed) Then
        MsgBox "Drive unavailable.", vbInformation, "Information"
        MsgBox strCurrentDrive
        DriveSelectionCombo.Drive = strCurrentDrive
        DriveSelectionCombo.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".DriveSelectionCombo_Change " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub OKButton_Click()
'Set the export path in the form frmExportQuads to the selected directory
'Close the form and unload
    On Error GoTo ErrorHandler

    frmExportQuads.ExportPath = ExportDirListBox.Path
    CancelButton_Click
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".OKButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

