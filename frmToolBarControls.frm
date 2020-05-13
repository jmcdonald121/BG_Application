VERSION 5.00
Begin VB.Form frmToolbarControls 
   Caption         =   "Select Quad"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboProject 
      Height          =   315
      ItemData        =   "frmToolBarControls.frx":0000
      Left            =   3000
      List            =   "frmToolBarControls.frx":0002
      TabIndex        =   1
      Text            =   "Open Project"
      Top             =   30
      Width           =   1575
   End
   Begin VB.ComboBox cboQuad 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a Quadrangle"
      Top             =   30
      Width           =   2895
   End
End
Attribute VB_Name = "frmToolbarControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    frmToolBarControls
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Creates quad selection combo box on the tool bar
'               and populates it with either 24k or 100k quads
'   Called from:
'
'*****************************************
Option Explicit

Private pQuadScaleList As Dictionary, pMapTypeList As Dictionary
Private m_pApp As esriFramework.IApplication

' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "frmToolBarControls"

Private Sub cboProject_Click()
On Error GoTo ErrorHandler

    Dim strText As String
    strText = cboProject.Text
    Me.Refresh
    gODNRProjectDb.OpenDocument strText
 
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cboProject_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboQuad_Click()
    On Error GoTo ErrorHandler

    Dim strText As String, pQuad As ODNRQuad
    strText = cboQuad.Text
    gODNRProject.Quads.AddQuadByName strText
    Set pQuad = gODNRProject.Quads.QuadByName(strText)
    gODNRProject.Quads.SetFocusQuad pQuad.QuadId
    Set gODNRProject.Quads.ExtentEnvelope = pQuad.QuadBoundary.Envelope
    gODNRProject.ZoomToQuadsExtent
    If (gODNRProject.ProjectType = odnrBedrockStructure) Or (gODNRProject.ProjectType = odnrGeology) Then
        gODNRProject.ShowBedrockLayers
    End If
    gODNRProject.ShowQuadLayers
    'Zoom to quad here
 
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".cboQuad_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

    cboProject.Text = "Open Project"
    cboQuad.Text = "Select a Quadrangle"
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Terminate()
  On Error GoTo ErrorHandler
     
     Set m_pApp = Nothing

  Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
