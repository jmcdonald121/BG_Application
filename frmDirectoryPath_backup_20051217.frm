VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDirectoryPath 
   Caption         =   "ODNR Geology Extension"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   Icon            =   "frmDirectoryPath_backup_20051217.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton FindBedrockDatabaseButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6570
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox ODNRBedrockDatabaseTextBox 
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
      Left            =   2160
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   420
      Width           =   4335
   End
   Begin VB.TextBox ODNRProjectsDatabaseTextBox 
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
      Left            =   2160
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   780
      Width           =   4335
   End
   Begin VB.CommandButton FindProjectsDatabaseButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6570
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":00AE
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin MSComDlg.CommonDialog FindGeoDatabaseDialog 
      Left            =   3780
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "GeoDatabase (*.mdb)|*.mdb|"
      InitDir         =   "C:"
   End
   Begin VB.CommandButton FindGeoDatabaseButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6570
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":0150
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox ODNRGeoDatabaseTextBox 
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
      Left            =   2160
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   60
      Width           =   4335
   End
   Begin VB.TextBox ScansDirTextBox 
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
      Left            =   3960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3690
      Width           =   2955
   End
   Begin VB.CommandButton ScansDirButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":01F2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3690
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox DRGLOCDirTextBox 
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
      Left            =   3960
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2550
      Width           =   2955
   End
   Begin VB.CommandButton DRGLOCDirButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":02F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2550
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox ExDirTextBox 
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
      Left            =   3960
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3120
      Width           =   2955
   End
   Begin VB.CommandButton ExDirButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":03F6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox DRGDirTextBox 
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
      Left            =   3960
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1980
      Width           =   2955
   End
   Begin VB.CommandButton DRGDirButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":04F8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox Base100DirTextBox 
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
      Left            =   3960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1410
      Width           =   2955
   End
   Begin VB.CommandButton Base100DirButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3540
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath_backup_20051217.frx":05FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1410
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.DriveListBox DriveSelectionCombo 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4950
      Width           =   3400
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5970
      TabIndex        =   13
      Top             =   4890
      Width           =   975
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4890
      TabIndex        =   12
      Top             =   4890
      Width           =   975
   End
   Begin VB.DirListBox DataDirListBox 
      Height          =   3240
      Left            =   60
      TabIndex        =   0
      Top             =   1410
      Width           =   3405
   End
   Begin VB.Label Label3 
      Caption         =   "ODNR Structure Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   28
      Top             =   450
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "ODNR Setup Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   26
      Top             =   810
      Width           =   1995
   End
   Begin VB.Label GeologyDBLabel 
      Caption         =   "ODNR Geology Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   90
      Width           =   1995
   End
   Begin VB.Label ScansDirLabel 
      Caption         =   "Scans Directory:"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   3450
      Width           =   1395
   End
   Begin VB.Label DRGLOCDirLabel 
      Caption         =   "DRG Data Directory:"
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
      Left            =   3960
      TabIndex        =   19
      Top             =   2310
      Width           =   1575
   End
   Begin VB.Label ExDirLabel 
      Caption         =   "Export Directory:"
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
      Left            =   3960
      TabIndex        =   18
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label DRGLabel 
      Caption         =   "DRG Layers Directory (lyr files):"
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
      Left            =   3960
      TabIndex        =   17
      Top             =   1740
      Width           =   2355
   End
   Begin VB.Label Base100Label 
      Caption         =   "Base100 Layers Directory (lyr files):"
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
      Left            =   3960
      TabIndex        =   16
      Top             =   1170
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Please identify the project directories below."
      Height          =   225
      Left            =   60
      TabIndex        =   15
      Top             =   1170
      Width           =   3255
   End
   Begin VB.Label InstructionsLabel 
      Caption         =   "Look in:"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   4710
      Width           =   735
   End
End
Attribute VB_Name = "frmDirectoryPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'
'   Program:    frmDirectoryPath
'   Author:     Gregory Palovchik
'               Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229 ext. 202
'   Date:       July 18, 2004
'   Purpose:    Provide a form for setting the directory paths
'               This module saves the settings in the registry at the following
'               location: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ArcView\ODNR_Geology
'
'   Called from: Select_DataDir_Cmd
'
'*****************************************
Option Explicit

Private strCurrentDrive As String
Private pFSO As FileSystemObject
Private m_pApp As esriFramework.IApplication


' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "frmDirectoryPath"

Public Property Set ESRIApplication(pApp As esriFramework.IApplication)
'Hook the application
    On Error GoTo ErrorHandler
    
    If (Not pApp Is Nothing) Then
        Set m_pApp = pApp
    End If

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRIApplication " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub DriveSelectionCombo_Change()
'Change the selected drive
    On Error GoTo ErrorHandler
    
    Dim strDrive As String, blnFailed As Boolean
    blnFailed = True
    strDrive = Left(DriveSelectionCombo.Drive, 2)
    If (pFSO.DriveExists(strDrive)) Then
        If (pFSO.Drives(strDrive).IsReady) Then
            strCurrentDrive = Strings.UCase(strDrive)
            DataDirListBox.Path = strCurrentDrive
            DataDirListBox.Refresh
            blnFailed = False
        End If
    End If
    If (blnFailed) Then
        MsgBox "Drive unavailable.", vbInformation, "Information"
        DriveSelectionCombo.Drive = strCurrentDrive
        DriveSelectionCombo.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".DriveSelectionCombo_Change " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub FindGeoDatabaseButton_Click()
'Locate the ODNR Geology database
    On Error GoTo ErrorHandler
    
    FindGeoDatabaseDialog.DialogTitle = "Locate ODNR Geology Database"
    FindGeoDatabaseDialog.ShowOpen
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pFWs As IFeatureWorkspace, pFC As IFeatureClass
    Dim blnWsError As Boolean
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pFWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pFC = pFWs.OpenFeatureClass("Quad24K")
        If (pFC Is Nothing) Then
            blnWsError = True
        Else
            ODNRGeoDatabaseTextBox.Text = strFileName
        End If
    Else
        blnWsError = True
    End If
    If (blnWsError) Then MsgBox "Not a valid ODNR Geology Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".FindGeoDatabaseButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub FindBedrockDatabaseButton_Click()
'Locate the ODNR Bedrock Structures database
    On Error GoTo ErrorHandler
    
    FindGeoDatabaseDialog.DialogTitle = "Locate ODNR Bedrock Database"
    FindGeoDatabaseDialog.ShowOpen
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pFWs As IFeatureWorkspace, pTbl As ITable
    Dim blnWsError As Boolean
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pFWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pTbl = pFWs.OpenTable("QuadUnitCodes")
        If (pTbl Is Nothing) Then
            blnWsError = True
        Else
            ODNRBedrockDatabaseTextBox.Text = strFileName
        End If
    Else
        blnWsError = True
    End If
    If (blnWsError) Then MsgBox "Not a valid ODNR Bedrock Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".FindGeoDatabaseButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub FindProjectsDatabaseButton_Click()
'Locate the ODNR project settings database
    On Error GoTo ErrorHandler
    
    FindGeoDatabaseDialog.DialogTitle = "Locate ODNR Projects Database"
    FindGeoDatabaseDialog.ShowOpen
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pWs As IWorkspace, pFC As IFeatureClass
    Dim blnWsError As Boolean, pProjectDb As ODNRProjectsDatabase
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pProjectDb = New ODNRProjectsDatabase
        If (pProjectDb.LoadDatabase(m_pApp, pWs)) Then
            ODNRProjectsDatabaseTextBox.Text = strFileName
        Else
            blnWsError = True
        End If
    Else
        blnWsError = True
    End If
    If (blnWsError) Then MsgBox "Not a valid ODNR Projects Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".FindGeoDatabaseButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
'Load form and set the default directory to the C: drive
    On Error GoTo ErrorHandler
    
    Set pFSO = New FileSystemObject
    DataDirListBox.Path = "C:"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Activate()
'Activate form and load the saved information into the appropriate text boxes
    On Error GoTo ErrorHandler
    
    DRGDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "DRGDirectory")
    'ProjectDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ProjectDirectory")
    ScansDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ScansDirectory")
    ExDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ExportDirectory")
    Base100DirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "Base100Directory")
    DRGLOCDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "DRGLOCDirectory")
    ODNRGeoDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "GeologyDatabasePath")
    ODNRBedrockDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "BedrockDatabasePath")
    ODNRProjectsDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ProjectsDatabasePath")
    
    strCurrentDrive = "C:"
    DriveSelectionCombo.Drive = "C:"
    DriveSelectionCombo_Change
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Activate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub OKButton_Click()
'Save the settings to the registry and close the form
    On Error GoTo ErrorHandler
    
    'SaveSetting "ArcView", "ODNR_Geology", "ProjectDirectory", ProjectDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "Base100Directory", Base100DirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "DRGDirectory", DRGDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "DRGLOCDirectory", DRGLOCDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ExportDirectory", ExDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ScansDirectory", ScansDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "GeologyDatabasePath", ODNRGeoDatabaseTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "BedrockDatabasePath", ODNRBedrockDatabaseTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ProjectsDatabasePath", ODNRProjectsDatabaseTextBox.Text
    LoadProjectsDatabase
    CancelButton_Click
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".OKButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CancelButton_Click()
'Close the form and unload
    On Error GoTo ErrorHandler
    
    Me.Hide
    Unload Me
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CancelButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Private Sub ProjectDirButton_Click()
'    On Error GoTo ErrorHandler
'
'    ProjectDirTextBox.Text = DataDirListBox.Path
'
'    Exit Sub
'ErrorHandler:
'    HandleError True, c_strModuleName & ".ProjectDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub

Private Sub Base100DirButton_Click()
'Copy the current path of the DataDirListBox to the Base100DirTextBox
    On Error GoTo ErrorHandler
    
    Base100DirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Base100DirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub DRGDirButton_Click()
'Copy the current path of the DataDirListBox to the DRGDirTextBox
    On Error GoTo ErrorHandler
    
    DRGDirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".DRGDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub DRGLOCDirButton_Click()
'Copy the current path of the DataDirListBox to the DRGLOCDirTextBox
    On Error GoTo ErrorHandler
    
    DRGLOCDirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".DRGLOCDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ExDirButton_Click()
'Copy the current path of the DataDirListBox to the ExDirTextBox
    On Error GoTo ErrorHandler
    
    ExDirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ExDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ScansDirButton_Click()
'Copy the current path of the DataDirListBox to the ScansDirTextBox
    On Error GoTo ErrorHandler
    
    ScansDirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ScansDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

