VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDirectoryPath 
   Caption         =   "Data sources for ODNR Geology Application"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   Icon            =   "frmDirectoryPath.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOGWellsDBPath 
      Height          =   300
      Left            =   6570
      TabIndex        =   52
      Top             =   2940
      Width           =   345
   End
   Begin VB.TextBox txtOGWellsDBPath 
      Height          =   315
      Left            =   2160
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2940
      Width           =   4335
   End
   Begin VB.CommandButton cmdOGWells 
      Caption         =   "Oil and Gas Wells"
      Height          =   735
      Left            =   4800
      TabIndex        =   49
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdDRG100K 
      Height          =   300
      Left            =   3660
      TabIndex        =   47
      Top             =   6600
      Width           =   345
   End
   Begin VB.TextBox txtDRG100K 
      Height          =   315
      Left            =   4080
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   6600
      Width           =   2955
   End
   Begin VB.ComboBox cboClip 
      Height          =   315
      Left            =   2160
      TabIndex        =   44
      Text            =   "Combo1"
      Top             =   3720
      Width           =   1955
   End
   Begin VB.TextBox txtAUMImages 
      Height          =   315
      Left            =   4080
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   7740
      Width           =   2955
   End
   Begin VB.CommandButton cmdAUMImages 
      Height          =   300
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7740
      Width           =   345
   End
   Begin VB.CommandButton cmdAumDBPath 
      Height          =   300
      Left            =   6570
      TabIndex        =   40
      Top             =   2580
      Width           =   345
   End
   Begin VB.TextBox txtAumDBPath 
      Height          =   315
      Left            =   2160
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2580
      Width           =   4335
   End
   Begin VB.CommandButton cmdBasemap 
      Caption         =   "Basemaps"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6120
      TabIndex        =   36
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdAUM 
      Caption         =   "AUM"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3480
      TabIndex        =   35
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdBG 
      Caption         =   "Geology BG/BT/BS"
      Height          =   735
      Left            =   2160
      TabIndex        =   34
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton optPgb 
      Caption         =   "Personnel GDB Data Sources"
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton optSde 
      Caption         =   "SDE Data Sources"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
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
      Picture         =   "frmDirectoryPath.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2220
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
      Top             =   2220
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
      Top             =   3300
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
      Picture         =   "frmDirectoryPath.frx":00AE
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3300
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin MSComDlg.CommonDialog FindGeoDatabaseDialog 
      Left            =   3900
      Top             =   8850
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
      Picture         =   "frmDirectoryPath.frx":0150
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1860
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
      Top             =   1860
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
      Left            =   4080
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   7170
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
      Left            =   3660
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath.frx":01F2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7170
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
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5460
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
      Left            =   3660
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath.frx":02F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5460
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
      Left            =   4080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   8310
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
      Left            =   3660
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath.frx":03F6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8310
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
      Left            =   4080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4890
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
      Left            =   3660
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath.frx":04F8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4890
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox txtDRG100KLyr 
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
      Left            =   4080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6030
      Width           =   2955
   End
   Begin VB.CommandButton cmdDRG100KLyr 
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
      Left            =   3660
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDirectoryPath.frx":05FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6030
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.DriveListBox DriveSelectionCombo 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   9030
      Width           =   3400
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6090
      TabIndex        =   13
      Top             =   8970
      Width           =   975
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5010
      TabIndex        =   12
      Top             =   8970
      Width           =   975
   End
   Begin VB.DirListBox DataDirListBox 
      Height          =   3690
      Left            =   180
      TabIndex        =   0
      Top             =   4890
      Width           =   3405
   End
   Begin VB.Label lblOGWellsDBPath 
      Caption         =   "ODNR Oil && Gas Database"
      Height          =   225
      Left            =   90
      TabIndex        =   50
      Top             =   2970
      Width           =   1995
   End
   Begin VB.Label lblDRG100K 
      Caption         =   "100K DRG Data Directory (TIF files):"
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
      Left            =   4080
      TabIndex        =   48
      Top             =   6360
      Width           =   2715
   End
   Begin VB.Label lblClipMethod 
      Caption         =   "Clipping Method"
      Height          =   225
      Left            =   90
      TabIndex        =   45
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label lblAUMImages 
      Caption         =   "AUM Images (TIF files):"
      Height          =   255
      Left            =   4080
      TabIndex        =   41
      Top             =   7500
      Width           =   2715
   End
   Begin VB.Label lblAumDBPath 
      Caption         =   "ODNR AUM Database"
      Height          =   225
      Left            =   90
      TabIndex        =   38
      Top             =   2610
      Width           =   1965
   End
   Begin VB.Label lblImagesExports 
      Caption         =   "Images, Layers, and Export Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   37
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label lblPgbDataSources 
      Caption         =   "Personnel GDB Data Sources"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblSdeDataSources 
      Caption         =   "SDE Data Sources"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "ODNR Structure Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   28
      Top             =   2250
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "ODNR Setup Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   26
      Top             =   3330
      Width           =   1995
   End
   Begin VB.Label GeologyDBLabel 
      Caption         =   "ODNR Geology Database:"
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   1890
      Width           =   1995
   End
   Begin VB.Label ScansDirLabel 
      Caption         =   "BS Scans Directory (TIF files):"
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
      Left            =   4080
      TabIndex        =   20
      Top             =   6960
      Width           =   2715
   End
   Begin VB.Label DRGLOCDirLabel 
      Caption         =   "24K DRG Data Directory (TIF files):"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   5220
      Width           =   2715
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
      Left            =   4080
      TabIndex        =   18
      Top             =   8040
      Width           =   2715
   End
   Begin VB.Label DRGLabel 
      Caption         =   "24K DRG Layers Directory (LYR files):"
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
      Left            =   4080
      TabIndex        =   17
      Top             =   4650
      Width           =   2715
   End
   Begin VB.Label lblDRG100KLyr 
      Caption         =   "100K Layers Directory (LYR files):"
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
      Left            =   4080
      TabIndex        =   16
      Top             =   5790
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Please identify the project directories below."
      Height          =   225
      Left            =   180
      TabIndex        =   15
      Top             =   4650
      Width           =   3255
   End
   Begin VB.Label InstructionsLabel 
      Caption         =   "Look in:"
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   8790
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

Private Sub cboClip_Click()
    On Error GoTo ErrorHandler
    
    MsgBox "Please change the ODNR Setup Database", vbExclamation, "Clipping Method Has Changed"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cboClip_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdAUM_Click()
    On Error GoTo ErrorHandler
    
    Load frmAUMSdeParameters
    frmAUMSdeParameters.Show vbModal
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdAUM_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdAumDBPath_Click()
'Locate the ODNR AUM database
    On Error GoTo ErrorHandler
    
    FindGeoDatabaseDialog.DialogTitle = "Locate ODNR AUM Database"
    FindGeoDatabaseDialog.ShowOpen
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pFWs As IFeatureWorkspace, pFc As IFeatureClass
    Dim blnWsError As Boolean
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pFWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pFc = pFWs.OpenFeatureClass("AUM_MINES")
        If (pFc Is Nothing) Then
            blnWsError = True
        Else
            txtAumDBPath.Text = strFileName
        End If
    Else
        blnWsError = True
    End If
    If (blnWsError) Then MsgBox "Not a valid ODNR AUM Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdAumDBPath_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdAUMImages_Click()
'Copy the current path of the DataDirListBox to the DRGLOCDirTextBox
    On Error GoTo ErrorHandler
    
    txtAUMImages.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdAUMImages_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdBasemap_Click()
    On Error GoTo ErrorHandler
    
    Load frmBasemapSdeParameters
    frmBasemapSdeParameters.Show vbModal
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdBasemap_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4


End Sub

Private Sub cmdBG_Click()
    On Error GoTo ErrorHandler
    
    Load frmSdeParameters
    frmSdeParameters.Show vbModal
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdBG_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdDRG100K_Click()
'Copy the current path of the DataDirListBox to the Base100DirTextBox
    On Error GoTo ErrorHandler
    
    txtDRG100K.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdDRG100K_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOGWells_Click()
    On Error GoTo ErrorHandler

    Load frmOGWellsSdeParameters
    frmOGWellsSdeParameters.Show vbModal
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdOGWells_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

Private Sub cmdOGWellsDBPath_Click()
'Locate the ODNR Oil & Gas Well Database
    On Error GoTo ErrorHandler
    
    FindGeoDatabaseDialog.DialogTitle = "Locate ODNR Oil & Gas Well Database"
    FindGeoDatabaseDialog.ShowOpen
    Dim strFileName As String, pWsFact As IWorkspaceFactory
    Dim pFWs As IFeatureWorkspace, pFc As IFeatureClass
    Dim blnWsError As Boolean
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pFWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pFc = pFWs.OpenFeatureClass("OG_WELLS")
        If (pFc Is Nothing) Then
            blnWsError = True
        Else
            txtOGWellsDBPath.Text = strFileName
        End If
    Else
        blnWsError = True
    End If
    If (blnWsError) Then MsgBox "Not a valid ODNR Oil & Gas Wells Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdOGWellsDBPath_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4

End Sub

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
    Dim pFWs As IFeatureWorkspace, pFc As IFeatureClass
    Dim blnWsError As Boolean
    Set pWsFact = New AccessWorkspaceFactory
    strFileName = FindGeoDatabaseDialog.FileName
    blnWsError = False
    If (pWsFact.IsWorkspace(strFileName)) Then
        Set pFWs = pWsFact.OpenFromFile(strFileName, 0)
On Error Resume Next
        Set pFc = pFWs.OpenFeatureClass("Quad24K")
        If (pFc Is Nothing) Then
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
    If (blnWsError) Then MsgBox "Not a valid ODNR Bedrock Structure Database.  Please select another database.", vbInformation, "ODNR Geology Extension"
    
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
    Dim pWs As IWorkspace, pFc As IFeatureClass
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
    
    Dim strSDEPGB As String
    
    'Add items to the Clipping Method CBO
    cboClip.AddItem "PreCut", 0
    cboClip.AddItem "Cut on the fly", 1
    
    'Read data from the registry
    DRGDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "DRGDirectory")
    ScansDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ScansDirectory")
    ExDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ExportDirectory")
    txtDRG100KLyr.Text = GetSetting("ArcView", "ODNR_Geology", "DRG100KLyrDirectory")
    txtDRG100K.Text = GetSetting("ArcView", "ODNR_Geology", "DRG100KDirectory")
    DRGLOCDirTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "DRGLOCDirectory")
    ODNRGeoDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "GeologyDatabasePath")
    ODNRBedrockDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "BedrockDatabasePath")
    ODNRProjectsDatabaseTextBox.Text = GetSetting("ArcView", "ODNR_Geology", "ProjectsDatabasePath")
    txtAumDBPath.Text = GetSetting("ArcView", "ODNR_Geology", "AUMDatabasePath")
    txtAUMImages.Text = GetSetting("ArcView", "ODNR_Geology", "AUMImagesPath")
    txtOGWellsDBPath.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsDatabasePath")

    cboClip.Text = GetSetting("ArcView", "ODNR_Geology", "ClipMethod")
    
    
    If (GetSetting("ArcView", "ODNR_Geology", "SDEPGB") = "SDE") Then
        optSde.Value = True
        optPgb.Value = False
        cmdBG.Enabled = True
        cmdAUM.Enabled = True
        cmdBasemap.Enabled = True
        cmdOGWells.Enabled = True
        FindGeoDatabaseButton.Enabled = False
        ODNRGeoDatabaseTextBox.Enabled = False
        GeologyDBLabel.Enabled = False
        lblAumDBPath.Enabled = False
        txtAumDBPath.Enabled = False
        cmdAumDBPath.Enabled = False
        lblOGWellsDBPath.Enabled = False
        txtOGWellsDBPath.Enabled = False
        cmdOGWellsDBPath.Enabled = False
    ElseIf (GetSetting("ArcView", "ODNR_Geology", "SDEPGB") = "PGB") Then
        optSde.Value = False
        optPgb.Value = True
        cmdBG.Enabled = False
        cmdAUM.Enabled = False
        cmdBasemap.Enabled = False
        cmdOGWells.Enabled = False
        FindGeoDatabaseButton.Enabled = True
        ODNRGeoDatabaseTextBox.Enabled = True
        GeologyDBLabel.Enabled = True
        lblAumDBPath.Enabled = True
        txtAumDBPath.Enabled = True
        cmdAumDBPath.Enabled = True
        lblOGWellsDBPath.Enabled = True
        txtOGWellsDBPath.Enabled = True
        cmdOGWellsDBPath.Enabled = True
    End If
    
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
    
    SaveSetting "ArcView", "ODNR_Geology", "DRG100KLyrDirectory", txtDRG100KLyr.Text
    SaveSetting "ArcView", "ODNR_Geology", "DRG100KDirectory", txtDRG100K.Text
    SaveSetting "ArcView", "ODNR_Geology", "DRGDirectory", DRGDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "DRGLOCDirectory", DRGLOCDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ExportDirectory", ExDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ScansDirectory", ScansDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "GeologyDatabasePath", ODNRGeoDatabaseTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "BedrockDatabasePath", ODNRBedrockDatabaseTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ProjectsDatabasePath", ODNRProjectsDatabaseTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "AUMDatabasePath", txtAumDBPath.Text
    SaveSetting "ArcView", "ODNR_Geology", "AUMImagesPath", txtAUMImages.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsDatabasePath", txtOGWellsDBPath.Text
    SaveSetting "ArcView", "ODNR_Geology", "ClipMethod", cboClip.Text
    If (optSde.Value = True) Then
        SaveSetting "ArcView", "ODNR_Geology", "SDEPGB", "SDE"
    ElseIf (optPgb.Value = True) Then
        SaveSetting "ArcView", "ODNR_Geology", "SDEPGB", "PGB"
    End If
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

Private Sub cmdDRG100KLyr_Click()
'Copy the current path of the DataDirListBox to the Base100DirTextBox
    On Error GoTo ErrorHandler
    
    txtDRG100KLyr.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cmdDRG100KLyr_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
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

Private Sub optPgb_Click()
    On Error GoTo ErrorHandler
    
    cmdBG.Enabled = False
    cmdAUM.Enabled = False
    cmdBasemap.Enabled = False
    cmdOGWells.Enabled = False
        
    ODNRGeoDatabaseTextBox.Enabled = True
    FindGeoDatabaseButton.Enabled = True
    GeologyDBLabel.Enabled = True
    lblAumDBPath.Enabled = True
    txtAumDBPath.Enabled = True
    cmdAumDBPath.Enabled = True
    lblOGWellsDBPath.Enabled = True
    txtOGWellsDBPath.Enabled = True
    cmdOGWellsDBPath.Enabled = True
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".optPgb_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub optSde_Click()
    On Error GoTo ErrorHandler
    
    cmdBG.Enabled = True
    cmdAUM.Enabled = True
    cmdBasemap.Enabled = True
    cmdOGWells.Enabled = True
    
    ODNRGeoDatabaseTextBox.Enabled = False
    FindGeoDatabaseButton.Enabled = False
    GeologyDBLabel.Enabled = False
    lblAumDBPath.Enabled = False
    txtAumDBPath.Enabled = False
    cmdAumDBPath.Enabled = False
    lblOGWellsDBPath.Enabled = False
    txtOGWellsDBPath.Enabled = False
    cmdOGWellsDBPath.Enabled = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".optSde_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ScansDirButton_Click()
'Copy the current path of the DataDirListBox to the ScansDirTextBox
    On Error GoTo ErrorHandler
    
    ScansDirTextBox.Text = DataDirListBox.Path
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ScansDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
