VERSION 5.00
Begin VB.Form frmBedrockStructures 
   Caption         =   "Bedrock Structures"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "frmBedrockStructures.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   1800
      Width           =   1005
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2820
      TabIndex        =   2
      Top             =   1800
      Width           =   1005
   End
   Begin VB.ListBox StructureList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   540
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Click the checkbox to make the layer visible."
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
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   3225
   End
   Begin VB.Label BedrockStructureLabel 
      Caption         =   "Bedrock Structures for "
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
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   3705
   End
End
Attribute VB_Name = "frmBedrockStructures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'
'  Program:     frmBedrockStructures
'  Author:      Gregory Palovchik
'               Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229 ext. 202
'  Date:        July 18, 2004
'  Purpose:     Provide a form for controlling the view of Bedrock structure
'               contours
'  Called from: Bedrock_Layers_Cmd
'
'*****************************************
Option Explicit

Private m_pApp As esriFramework.IApplication
' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "frmBedrockStructures"


Public Property Set App(RHS As esriFramework.IApplication)
'Hook application
    On Error GoTo ErrorHandler

    Set m_pApp = RHS
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".App " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub Form_Load()
'Load form, find bedrock structure contours layers in the map and populate
'the Structure listbox.
    On Error GoTo ErrorHandler

    Dim pBedrockLayer As ODNRBedrockLayer, pGlyr As IGroupLayer
    Dim pClyr As ICompositeLayer, pLyr As ILayer, lngLyrIdx As Long
    Dim strQuadName As String, blnSelected As Boolean, strType As String, intPos As Integer
    Set pBedrockLayer = gODNRProject.BedrockLayers.GetLayerByName("BS Contours")
    strQuadName = gODNRProject.Quads.FocusQuad.QuadName
    If Not (pBedrockLayer Is Nothing) Then
        blnSelected = False
        Set pGlyr = pBedrockLayer.ESRILayer
        Set pClyr = pGlyr
        For lngLyrIdx = 0 To pClyr.Count - 1
            Set pLyr = pClyr.Layer(lngLyrIdx)
            intPos = Strings.InStr(1, pLyr.Name, "_", vbTextCompare)
            strType = Strings.Left(pLyr.Name, intPos - 1)
            StructureList.AddItem strType
            If (pLyr.Visible) And (blnSelected = False) Then
                StructureList.Selected(lngLyrIdx) = True
                blnSelected = True
            End If
        Next
    End If
    BedrockStructureLabel.Caption = "BS Contours for " & strQuadName
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CancelButton_Click()
'Cancel and unload the form
    On Error GoTo ErrorHandler
    
    Me.Hide
    Unload Me

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CancelButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub OKButton_Click()
'Evaluate the user selection in the Structure Listbox and make the selected
'bedrock structure contours, faults, termintators, and anno visible
    On Error GoTo ErrorHandler

    Dim intIdx As Integer, strName As String
    For intIdx = 0 To StructureList.ListCount - 1
        If (StructureList.Selected(intIdx)) Then
            strName = StructureList.List(intIdx)
            Exit For
        End If
    Next

    Dim pBedrockLayer As ODNRBedrockLayer, pGlyr As IGroupLayer
    Dim pClyr As ICompositeLayer, pLyr As ILayer, lngLyrIdx As Long
    Dim pBedrockGlyrs As Dictionary, vKey As Variant, strLyrName As String
    
    Set pBedrockGlyrs = New Dictionary
'    pBedrockGlyrs.Add Key:="Bedrock Structures", Item:="_Contour"
'    pBedrockGlyrs.Add Key:="Bedrock Structures Anno", Item:="_Contour_Anno"
'    pBedrockGlyrs.Add Key:="Bedrock Faults", Item:="_Fault"
'    pBedrockGlyrs.Add Key:="Bedrock Faults Anno", Item:="_Fault_Anno"
'    pBedrockGlyrs.Add Key:="Bedrock Contour Terminators", Item:="_Term"
    pBedrockGlyrs.Add Key:="BS Contours", Item:="_Contour"
    pBedrockGlyrs.Add Key:="BS Contours Anno", Item:="_Contour_Anno"
    pBedrockGlyrs.Add Key:="BS Faults", Item:="_Fault"
    pBedrockGlyrs.Add Key:="BS Faults Anno", Item:="_Fault_Anno"
    pBedrockGlyrs.Add Key:="BS Contour Terminators", Item:="_Term"
    pBedrockGlyrs.Add Key:="BS Datapoints", Item:="_Datapoints"

    
    For Each vKey In pBedrockGlyrs.Keys
        Set pBedrockLayer = gODNRProject.BedrockLayers.GetLayerByName(CStr(vKey))
        strLyrName = strName & pBedrockGlyrs.Item(vKey)
        If Not (pBedrockLayer Is Nothing) Then
            Set pGlyr = pBedrockLayer.ESRILayer
            Set pClyr = pGlyr
            For lngLyrIdx = 0 To pClyr.Count - 1
                Set pLyr = pClyr.Layer(lngLyrIdx)
                If (pLyr.Name = strLyrName) Then
                    pLyr.Visible = True
                Else
                    pLyr.Visible = False
                End If
            Next
        End If
    Next
    
    Dim pMxDoc As IMxDocument, pContView As IContentsView
    Set pMxDoc = m_pApp.Document
    Set pContView = pMxDoc.CurrentContentsView
    pContView.Refresh Null
    pMxDoc.ActiveView.Refresh
    
    CancelButton_Click
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".OKButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub StructureList_ItemCheck(Item As Integer)
'Make sure that only the selected item is checked
    On Error GoTo ErrorHandler
    
    Dim intIdx As Integer
    For intIdx = 0 To StructureList.ListCount - 1
        If (Item <> intIdx) Then
            StructureList.Selected(intIdx) = False
        Else
            StructureList.Selected(intIdx) = True
        End If
    Next

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".StructureList_ItemCheck " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
