VERSION 5.00
Begin VB.Form frmOGWellsSdeParameters 
   Caption         =   "Oil & Gas Wells SDE Connection Parameters"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSdeVersion 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Text            =   "SDE Version"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtSdeInstance 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Text            =   "SDE Instance"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtSdePassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Password"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtSdeUserName 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   "User Name"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtSdeDatabaseName 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "Database Name"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtSdeServer 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   "Server Name"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblSdeVersion 
      Caption         =   "SDE VERSION"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblSdeInstance 
      Caption         =   "SDE INSTANCE"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblSdePassword 
      Caption         =   "SDE PASSWORD"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblSdeUser 
      Caption         =   "SDE USER NAME"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblSdeDatabase 
      Caption         =   "SDE DATABASE"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblSdeServer 
      Caption         =   "SDE SERVER:"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblSdeParameters 
      Caption         =   "Enter the SDE connection parameters"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOGWellsSdeParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Title: frmOGWellsSdeParameters.frm
'Date: 20051219
'Version: 1.0
'Abstract:  This form loads the OGWells SDE connection parameters into the registry
'at the following location: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ArcView\ODNR_Geology
'------------------------------------------------------------------------------
'James McDonald
'GIMS Specialist
'Ohio Division of Geological Survey
'2045 Morse Road
'Columbus, OH  43229-6693
'Ph. (614) 265-6601
'E-mail: jim.mcdonald@dnr.state.oh.us
'------------------------------------------------------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmOGWellsSdeParameters.frm"

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

'Unload the form
    
    Me.Hide
    Unload Me
    

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

'Save the data into the registry

    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEServer", txtSdeServer.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEDatabase", txtSdeDatabaseName.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEUser", txtSdeUserName.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEInstance", txtSdeInstance.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEPassword", txtSdePassword.Text
    SaveSetting "ArcView", "ODNR_Geology", "OGWellsSDEVersion", txtSdeVersion.Text
    cmdCancel_Click


  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorHandler

'Activate form and load the saved information into the appropriate text boxes
    
    txtSdeServer.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEServer")
    txtSdeDatabaseName.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEDatabase")
    txtSdeUserName.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEUser")
    txtSdeInstance.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEInstance")
    txtSdePassword.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEPassword")
    txtSdeVersion.Text = GetSetting("ArcView", "ODNR_Geology", "OGWellsSDEVersion")

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

    'Possibly Add something here

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

