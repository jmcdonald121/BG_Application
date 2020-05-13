VERSION 5.00
Begin VB.Form frmExportQuads 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export Quads to Shapefiles"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExportDirButton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4020
      Picture         =   "frmExportQuads.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click to remove a selected quad"
      Top             =   1980
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox ExportDirTextBox 
      Height          =   285
      Left            =   870
      TabIndex        =   8
      Top             =   1980
      Width           =   3105
   End
   Begin VB.CommandButton RemoveQuadButton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4020
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmExportQuads.frx":00A2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click to remove a selected quad"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3300
      TabIndex        =   6
      Top             =   2310
      Width           =   1005
   End
   Begin VB.CheckBox SinglePackageCheckBox 
      Caption         =   "Export Quads together"
      Enabled         =   0   'False
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   2310
      Width           =   1995
   End
   Begin VB.ListBox SelectedQuadsListBox 
      Height          =   1620
      Left            =   2220
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   300
      Width           =   2055
   End
   Begin VB.CommandButton ExportButton 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Top             =   2310
      Width           =   1005
   End
   Begin VB.ListBox MapLayersListBox 
      Height          =   1635
      Left            =   30
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label StatusLabel 
      Height          =   225
      Left            =   2040
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label3 
      Caption         =   "Export to:"
      Height          =   225
      Left            =   30
      TabIndex        =   9
      Top             =   2010
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Selected quads:"
      Height          =   255
      Left            =   2220
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Check map layers to export."
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "frmExportQuads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    frmExportQuads
'   Author:     Gregory Palovchik
'               Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229 ext. 202
'   Date:       June 28, 2004
'   Purpose:    Provide a form for controlling the export of geology data
'
'   Called from: Export_Tool
'
'*****************************************
Option Explicit

Private m_strExport_Path As String
Private m_pApp As esriFramework.IApplication
'Private m_pQuadFc As IFeatureClass
'Private m_lngNameFieldIdx As Long
Private m_pSelectedQuads As Dictionary
Private m_pLayerList As Dictionary
Private m_pLayerTypes As Dictionary
Private m_pInitialQuadList As Dictionary
Private m_pFSO As FileSystemObject

' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "frmExportQuads"

Public Property Set App(RHS As esriFramework.IApplication)
'Hook the application
    On Error GoTo ErrorHandler

    Set m_pApp = RHS
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".App " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let ExportPath(RHS As String)
'Set the export path from outside the form.
    On Error GoTo ErrorHandler

    If (RHS <> "") Then
        If (m_pFSO.FolderExists(RHS)) Then
            m_strExport_Path = RHS
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportPath " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub SelectedQuads(pQuadList As Collection, intAction As Integer)
'Load selected quads from outside the form. See Export_Tool.ITool_OnMouseUp
    On Error GoTo ErrorHandler

    Dim pQuad As ODNRQuad, lngIdx As Long, lngId As Long, lngListIdx As Long
    If Not (pQuadList Is Nothing) Then
        If (pQuadList.Count > 0) Then
            If (intAction = 0) Or (intAction = 1) Then 'Add to selected
                For lngIdx = 1 To pQuadList.Count
                    lngId = pQuadList.Item(lngIdx)
                    If (m_pSelectedQuads.Exists(lngId) = False) Then
                        gODNRProject.Quads.AddQuadById lngId
                        Set pQuad = gODNRProject.Quads.QuadById(lngId)
                        SelectedQuadsListBox.AddItem pQuad.QuadName
                        SelectedQuadsListBox.ItemData(SelectedQuadsListBox.ListCount - 1) = lngId
                        m_pSelectedQuads.Add Key:=lngId, Item:=pQuad.QuadName
                    End If
                Next
            ElseIf (intAction = 2) Then
                For lngIdx = 1 To pQuadList.Count
                    lngId = pQuadList.Item(lngIdx)
                    RemoveQuad lngId
                Next
            End If
        End If
        gODNRProject.Quads.HighlightQuads
    End If
    UpdateControls
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".SelectedQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CancelButton_Click()
'Close the from and unselect the export tool button by selecting the pointer button
    On Error GoTo ErrorHandler

    Me.Hide
    gODNRProject.Quads.UnHighlightQuads
    Dim pUID As New UID
    Dim pCmdItem As ICommandItem
    pUID.Value = "{C22579D1-BC17-11D0-8667-0000F8751720}" 'Select_Elements tool
    Set pCmdItem = m_pApp.Document.CommandBars.Find(pUID)
    pCmdItem.Execute
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".CancelButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ExportButton_Click()
'Export the selected layers for the list of quads
    On Error GoTo ErrorHandler

    Dim strRootDir As String, strExportDir As String, strTempDir As String
    Dim pFileList As Collection, blnExportOk As Boolean, blnSinglePackage As Boolean
    Dim pSubDirExportList As Dictionary, blnExported As Boolean
    Dim pFolder As Folder, pSubFolder As Folder, pFile As File, intResp As Integer
    Dim pTxtStrm As TextStream, blnExportErrors As Boolean
    Dim strDate As String, strBaseName As String
    
    If (m_pFSO.FolderExists(ExportDirTextBox.Text) = False) Then
        MsgBox "Export directory not found.  Please try again.", vbInformation, "Export Quads"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Me.Height = 2970 + 275
    Me.StatusLabel.Top = 2660
    Me.StatusLabel.Left = 30
    Me.StatusLabel.Width = 4000
    Me.StatusLabel.Visible = True
    Me.StatusLabel.Caption = "Preparing to export layers..."
    Me.Refresh
    
    blnSinglePackage = SinglePackageCheckBox.Value
    blnExportErrors = False
    
    strDate = CStr(Year(Now())) & "-" & Right("0" & CStr(Month(Now())), 2) & "-" & Right("0" & CStr(Day(Now())), 2) & "_" & Right("0" & CStr(Hour(Now())), 2) & "-" & Right("0" & CStr(Minute(Now())), 2)
    strRootDir = ExportDirTextBox.Text & "\" & strDate
    m_pFSO.CreateFolder strRootDir
    
    Set pTxtStrm = m_pFSO.CreateTextFile(strRootDir & "\ExportLog.txt", True, False)
    pTxtStrm.WriteLine "ODNR Geology Export Log.  Date: " & Now()
    Me.StatusLabel.Caption = "Getting temp directory..."
    Me.Refresh
    strTempDir = TempDirectory(strRootDir)
    gODNRProject.Quads.UnHighlightQuads
    gODNRProject.Quads.RemoveAll
    If (m_pFSO.FolderExists(strRootDir) = False) Then
        MsgBox "The export directory " & strRootDir & " could not be found." & vbCrLf & "Please located the directory and try again.", vbInformation, "Export Quads"
        Exit Sub
    Else
        Dim pStateLayer As ODNRStateLayer, pQuadLayer As ODNRQuadLayer, pBedrockLayer As ODNRBedrockLayer
        Dim pQuad As ODNRQuad
        Dim lngIdx As Long, strName As String, vKey As Variant
        'if the export is a single package or only one quad then set the quads and
        'do the export.  If the export is by individual quads then set the quad one
        'by one and export the layers.
        For lngIdx = 0 To MapLayersListBox.ListCount - 1
            If (MapLayersListBox.Selected(CInt(lngIdx))) Then
                strName = m_pLayerList.Item(lngIdx)
                If (m_pLayerTypes.Item(lngIdx) = "STATE_WIDE") Then
                    Set pStateLayer = gODNRProject.StateLayers.GetLayerByName(strName)
                    Me.StatusLabel.Caption = "Exporting " & strName & "..."
                    Me.Refresh
                    pTxtStrm.WriteLine pStateLayer.Export2(strTempDir, pStateLayer.Name)
                    MapLayersListBox.Selected(CInt(lngIdx)) = False
                End If
            End If
        Next
        Set pSubDirExportList = New Dictionary
        If (blnSinglePackage) Or (m_pSelectedQuads.Count = 1) Then
            For Each vKey In m_pSelectedQuads.Keys
                gODNRProject.Quads.AddQuadById CLng(vKey)
            Next
            For lngIdx = 0 To MapLayersListBox.ListCount - 1
                If (MapLayersListBox.Selected(CInt(lngIdx))) Then
                    strName = m_pLayerList.Item(lngIdx)
                    Me.StatusLabel.Caption = "Exporting " & strName & "..."
                    Me.Refresh
                    If (m_pLayerTypes.Item(lngIdx) = "STATE_QUAD") Then
                        Set pStateLayer = gODNRProject.StateLayers.GetLayerByName(strName)
                        pStateLayer.LimitToQuads
                        pTxtStrm.WriteLine pStateLayer.Export2(strTempDir, pStateLayer.Name)
                    ElseIf (m_pLayerTypes.Item(lngIdx) = "BEDROCK") Then
                        gODNRProject.BedrockLayers.Refresh
                        Set pBedrockLayer = gODNRProject.BedrockLayers.GetLayerByName(strName)
                        pTxtStrm.WriteLine pBedrockLayer.Export(strTempDir)
                    ElseIf (m_pLayerTypes.Item(lngIdx) = "QUAD") Then
                        gODNRProject.QuadLayers.Refresh
                        Set pQuadLayer = gODNRProject.QuadLayers.GetLayerByName(strName)
                        pTxtStrm.WriteLine pQuadLayer.Export(strTempDir)
                    End If
                    MapLayersListBox.Selected(CInt(lngIdx)) = False
                End If
            Next
        ElseIf (m_pSelectedQuads.Count > 1) Then
            For Each vKey In m_pSelectedQuads.Keys
                gODNRProject.Quads.AddQuadById CLng(vKey)
                Set pQuad = gODNRProject.Quads.QuadById(CLng(vKey))
                strExportDir = strTempDir & "\" & pQuad.QuadName
                If (CheckDirectory(strExportDir)) Then
                    If (pSubDirExportList.Exists(strExportDir) = False) Then pSubDirExportList.Add Key:=strExportDir, Item:=Nothing
                    For lngIdx = 0 To MapLayersListBox.ListCount - 1
                        If (MapLayersListBox.Selected(CInt(lngIdx))) Then
                            strName = m_pLayerList.Item(lngIdx)
                            Me.StatusLabel.Caption = "Exporting " & strName & "..."
                            Me.Refresh
                            If (m_pLayerTypes.Item(lngIdx) = "STATE_QUAD") Then
                                Set pStateLayer = gODNRProject.StateLayers.GetLayerByName(strName)
                                pStateLayer.LimitToQuads
                                pTxtStrm.WriteLine pStateLayer.Export2(strExportDir, pStateLayer.Name)
                            ElseIf (m_pLayerTypes.Item(lngIdx) = "BEDROCK") Then
                                gODNRProject.BedrockLayers.Refresh
                                Set pBedrockLayer = gODNRProject.BedrockLayers.GetLayerByName(strName)
                                pTxtStrm.WriteLine pBedrockLayer.Export(strExportDir)
                            ElseIf (m_pLayerTypes.Item(lngIdx) = "QUAD") Then
                                gODNRProject.QuadLayers.Refresh
                                Set pQuadLayer = gODNRProject.QuadLayers.GetLayerByName(strName)
                                pTxtStrm.WriteLine pQuadLayer.Export(strExportDir)
                            End If
                        End If
                    Next
                    gODNRProject.Quads.RemoveAll
                End If
            Next
        End If
        
        Dim pPackager As ODNRPackager, strFileName As String
        Me.StatusLabel.Caption = "Packaging export files..."
        Me.Refresh
        pTxtStrm.WriteLine "Packaging export files..."
        Set pPackager = New ODNRPackager
        If (blnSinglePackage) Or (m_pSelectedQuads.Count = 1) Then
            pTxtStrm.WriteLine "Zipping files in " & strTempDir
            If (blnSinglePackage = False) Then
                blnExportOk = pPackager.Package(strTempDir, strRootDir & "\" & m_pSelectedQuads.Item(m_pSelectedQuads.Keys(0)) & ".zip")
                pTxtStrm.WriteLine vbTab & "to " & strRootDir & "\" & m_pSelectedQuads.Item(m_pSelectedQuads.Keys(0)) & ".zip"
            Else
                blnExportOk = pPackager.Package(strTempDir, strRootDir & "\" & strDate & ".zip")
                pTxtStrm.WriteLine vbTab & "to " & strRootDir & "\" & strDate & " .zip"
            End If
            If (blnExportOk) Then
                pTxtStrm.WriteLine vbTab & "Successfully zipped files in " & strTempDir
            Else
                pTxtStrm.WriteLine vbTab & "Error while zipping files in " & strTempDir
            End If
        Else
            For Each vKey In pSubDirExportList.Keys
                strExportDir = CStr(vKey)
                pTxtStrm.WriteLine "Zipping files in " & strExportDir
                strBaseName = m_pFSO.GetBaseName(strExportDir)
                blnExportOk = pPackager.Package(strExportDir, strRootDir & "\" & strBaseName & ".zip")
                blnExportOk = pPackager.Package(strTempDir, strRootDir & "\" & strBaseName & ".zip")
                If (blnExportOk) Then
                    pTxtStrm.WriteLine vbTab & "Successfully zipped files in " & CStr(vKey) & vbCrLf & " at " & strRootDir & "\" & strBaseName & ".zip"
                Else
                    pTxtStrm.WriteLine vbTab & "Error while zipping files in " & CStr(vKey) & vbCrLf & " at " & strRootDir & "\" & strBaseName & ".zip"
                End If
            Next
        End If
        Set pPackager = Nothing
        Me.StatusLabel.Caption = "Removing temporary folders and files..."
        Me.Refresh
        On Error Resume Next
        pTxtStrm.WriteLine "Removing temporary folders..."
        Set pFolder = m_pFSO.GetFolder(strTempDir)
        For Each pSubFolder In pFolder.SubFolders
            pSubFolder.Delete True
        Next
        pTxtStrm.WriteLine "Removing temporary files..."
        m_pFSO.DeleteFile strTempDir & "\*.*", True
        pTxtStrm.WriteLine "Export complete."
        pTxtStrm.Close
        Me.StatusLabel.Caption = "Finished."
        On Error GoTo ErrorHandler
        intResp = MsgBox("Finished exporting.  Would you like to view the export log?", vbYesNo, "Export Quads")
        If (intResp = 6) Then
            Shell "Notepad.exe " & strRootDir & "\ExportLog.txt", vbNormalFocus
        End If
    End If
    
    'Added Section 20051123, Jim McDonald
    For lngIdx = 0 To MapLayersListBox.ListCount - 1
        strName = m_pLayerList.Item(lngIdx)
        If (m_pLayerTypes.Item(lngIdx) = "STATE_QUAD") Then
            Set pStateLayer = gODNRProject.StateLayers.GetLayerByName(strName)
            pStateLayer.ShowAllFeatures
        End If
    Next
    'End of Added Section 20051123, Jim McDonald
    
    Screen.MousePointer = vbDefault
    strRootDir = ExportDirTextBox.Text
    SaveSetting "ArcView", "ODNR_Geology", "ExportDirectory", strRootDir
    g_strExport_Path = strRootDir & "\"
    CancelButton_Click
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ExportDirButton_Click()
'Change the export directory button
    On Error GoTo ErrorHandler

    Load frmExportQuadsDirectory
    frmExportQuadsDirectory.ExportPath = ExportDirTextBox.Text
    frmExportQuadsDirectory.Show vbModal, Me
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportDirButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub LoadMapLayers()
'Load the appropriate map layers to the MapLayersListBox for the project
    On Error GoTo ErrorHandler

    Dim pMap As IMap, pLayer As ILayer, lngLayerIdx As Long
    Set pMap = gODNRProject.ProjectMap(odnrGeologyMap)
    Dim pStateLayers As ODNRStateLayers, pStateLayer As ODNRStateLayer
    Dim pQuadLayers As ODNRQuadLayers, pQuadLayer As ODNRQuadLayer
    Dim pBedrockLayers As ODNRBedrockLayers, pBedrockLayer As ODNRBedrockLayer
    Set pStateLayers = gODNRProject.StateLayers
    Set pQuadLayers = gODNRProject.QuadLayers
    Set pBedrockLayers = gODNRProject.BedrockLayers
    Set m_pLayerList = New Dictionary
    Set m_pLayerTypes = New Dictionary
    lngLayerIdx = 0
    pStateLayers.ActiveMap = odnrGeologyMap
    pStateLayers.Reset
    Set pStateLayer = pStateLayers.NextLayer
    Do While Not pStateLayer Is Nothing
        m_pLayerList.Add Key:=lngLayerIdx, Item:=pStateLayer.Name
        If (((gODNRProject.QuadScale = odnr100K) And (pStateLayer.CanQueryBy100KQuad)) Or _
            ((gODNRProject.QuadScale = odnr24K) And (pStateLayer.CanQueryBy24KQuad))) Then
            m_pLayerTypes.Add Key:=lngLayerIdx, Item:="STATE_QUAD"
        Else
            m_pLayerTypes.Add Key:=lngLayerIdx, Item:="STATE_WIDE"
        End If
        MapLayersListBox.AddItem pStateLayer.Name
        MapLayersListBox.Selected(CInt(MapLayersListBox.ListCount - 1)) = True
        Set pStateLayer = pStateLayers.NextLayer
        lngLayerIdx = lngLayerIdx + 1
    Loop
    If (Not pBedrockLayers Is Nothing) Then
        pBedrockLayers.Reset
        Set pBedrockLayer = pBedrockLayers.NextLayer
        Do While Not pBedrockLayer Is Nothing
            If (pBedrockLayer.CanExport) Then
                m_pLayerList.Add Key:=lngLayerIdx, Item:=pBedrockLayer.Name
                m_pLayerTypes.Add Key:=lngLayerIdx, Item:="BEDROCK"
                MapLayersListBox.AddItem pBedrockLayer.Name
                MapLayersListBox.Selected(CInt(MapLayersListBox.ListCount - 1)) = True
                lngLayerIdx = lngLayerIdx + 1
            End If
            Set pBedrockLayer = pBedrockLayers.NextLayer
        Loop
    End If
    pQuadLayers.Reset
    Set pQuadLayer = pQuadLayers.NextLayer
    Do While Not pQuadLayer Is Nothing
        If (pQuadLayer.CanExport) Then
            m_pLayerList.Add Key:=lngLayerIdx, Item:=pQuadLayer.Name
            m_pLayerTypes.Add Key:=lngLayerIdx, Item:="QUAD"
            MapLayersListBox.AddItem pQuadLayer.Name
            MapLayersListBox.Selected(CInt(MapLayersListBox.ListCount - 1)) = True
            lngLayerIdx = lngLayerIdx + 1
        End If
        Set pQuadLayer = pQuadLayers.NextLayer
    Loop
    MapLayersListBox.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadMapLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Activate()
'Activate form and load the initial quad list.
    On Error GoTo ErrorHandler

    If Not (gODNRProject Is Nothing) Then
        ExportDirTextBox.Text = m_strExport_Path
        If Not (frmExportQuadsDirectory Is Nothing) Then
            Set frmExportQuadsDirectory = Nothing
        End If
        Me.Refresh
        If (gODNRProject.IsZoomedToQuadSelection) Then
            Dim pQuad As ODNRQuad, lngId As Long
            Set m_pInitialQuadList = New Dictionary
            If Not (gODNRProject.Quads.FocusQuad Is Nothing) Then
                Set pQuad = gODNRProject.Quads.FocusQuad
                m_pInitialQuadList.Add Key:=pQuad.QuadId, Item:=True
            Else
                gODNRProject.Quads.Reset
                Set pQuad = gODNRProject.Quads.NextQuad
                Do While Not pQuad Is Nothing
                    m_pInitialQuadList.Add Key:=pQuad.QuadId, Item:=False
                    Set pQuad = gODNRProject.Quads.NextQuad
                Loop
            End If
        End If
        If (m_pApp Is Nothing) Then MsgBox "App is nothing"
    Else
        Unload Me
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Activate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Initialize()
'Initialize form and private variables
    On Error GoTo ErrorHandler

    If Not (gODNRProject Is Nothing) Then
'        Dim pFlyr As IFeatureLayer
'        Set pFlyr = gODNRProject.QuadFeatureLayer
'        Set m_pQuadFc = pFlyr.FeatureClass
'        If (gODNRProject.QuadScale = odnr24K) Then
'            m_lngNameFieldIdx = m_pQuadFc.FindField("QUADNAME")
'        ElseIf (gODNRProject.QuadScale = odnr100K) Then
'            m_lngNameFieldIdx = m_pQuadFc.FindField("NAME")
'        End If
        Set m_pSelectedQuads = New Dictionary
        Set m_pFSO = New FileSystemObject
        m_strExport_Path = g_strExport_Path
        If (Right(m_strExport_Path, 1) = "\") Then m_strExport_Path = Left(m_strExport_Path, Len(m_strExport_Path) - 1)
        g_blnExportDialogOpen = True
        LoadMapLayers
    Else
        Unload Me
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Terminate
End Sub

Private Sub Form_Terminate()
'Terminate the form and unhighlight any selected quads.  If zoomed to quad selection
'reselect the focus quad
    On Error GoTo ErrorHandler

    If Not (gODNRProject Is Nothing) Then
        If (gODNRProject.IsZoomedToQuadSelection) And (Not m_pInitialQuadList Is Nothing) Then
            Dim pQuad As ODNRQuad, lngId As Long, vKey As Variant
            Dim pRemoveList As Collection, lngIdx As Long
            gODNRProject.Quads.RemoveAll
            Set pQuad = gODNRProject.Quads.NextQuad
            Set pRemoveList = New Collection
            Do While Not pQuad Is Nothing
                If (m_pInitialQuadList.Exists(pQuad.QuadId) = False) Then
                    pRemoveList.Add Item:=pQuad.QuadId
                Else
                    m_pInitialQuadList.Remove pQuad.QuadId
                End If
                Set pQuad = gODNRProject.Quads.NextQuad
            Loop
            If (pRemoveList.Count > 0) Then
                For lngIdx = 1 To pRemoveList.Count
                    lngId = pRemoveList.Item(lngIdx)
                    gODNRProject.Quads.RemoveQuadById lngId
                Next
            End If
            If (m_pInitialQuadList.Count > 0) Then
                For Each vKey In m_pInitialQuadList.Keys
                    gODNRProject.Quads.AddQuadById CLng(vKey)
                    If (CBool(m_pInitialQuadList.Item(vKey))) Then
                        gODNRProject.Quads.SetFocusQuad CLng(vKey)
                    End If
                Next
            End If
        Else
            gODNRProject.Quads.RemoveAll
        End If
    End If
    If Not (gODNRProject.Quads.FocusQuad Is Nothing) Then
        If (Not gODNRProject.BedrockLayers Is Nothing) Then
            gODNRProject.BedrockLayers.Refresh
        End If
        gODNRProject.QuadLayers.Refresh
    End If
    g_blnExportDialogOpen = False
    Set m_pFSO = Nothing
    Set m_pApp = Nothing
    ODNR_Common.SelectPointerTool
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Form_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub RemoveQuad(lngId As Long)
'Remove quad from selected quads list
    On Error GoTo ErrorHandler

    Dim intListIdx As Integer
    If (m_pSelectedQuads.Exists(lngId)) Then
        m_pSelectedQuads.Remove lngId
        For intListIdx = 0 To SelectedQuadsListBox.ListCount - 1
            If (SelectedQuadsListBox.ItemData(intListIdx) = lngId) Then
                gODNRProject.Quads.RemoveQuadById lngId
                SelectedQuadsListBox.RemoveItem intListIdx
                Exit For
            End If
        Next
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".RemoveQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub RemoveQuadButton_Click()
'Remove selected quad from list
    On Error GoTo ErrorHandler

    Dim intListIdx As Integer, lngId As Long, pList As Collection
    Set pList = New Collection
    For intListIdx = 0 To SelectedQuadsListBox.ListCount - 1
        If (SelectedQuadsListBox.Selected(intListIdx)) Then
            lngId = SelectedQuadsListBox.ItemData(intListIdx)
            pList.Add Item:=lngId
        End If
    Next
    If (pList.Count > 0) Then
        For intListIdx = 1 To pList.Count
            RemoveQuad pList.Item(intListIdx)
        Next
    End If
    gODNRProject.Quads.HighlightQuads
    UpdateControls
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".RemoveQuadButton_Click " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub UpdateControls()
'Refresh control states.  If one or more quads are selected enable export button.
'If two or more are selected enable the SinglePackageCheckBox.
    On Error GoTo ErrorHandler

    SinglePackageCheckBox.Enabled = False
    ExportButton.Enabled = False
    If (SelectedQuadsListBox.ListCount > 0) Then
        ExportButton.Enabled = True
        If (SelectedQuadsListBox.ListCount > 1) Then SinglePackageCheckBox.Enabled = True
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UpdateControls " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function CheckDirectory(strExportDir As String) As Boolean
'Check if the export directory is valid
    On Error GoTo ErrorHandler

    CheckDirectory = False
    Dim strParentFolder As String
    strParentFolder = m_pFSO.GetParentFolderName(strExportDir)
    If (m_pFSO.FolderExists(strParentFolder)) Then
        If (m_pFSO.FolderExists(strExportDir) = False) Then
            m_pFSO.CreateFolder strExportDir
        End If
        If (m_pFSO.FolderExists(strExportDir)) Then
            CheckDirectory = True
        End If
    End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".CheckDirectory " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'Private Sub ExportList(ByRef pExportList As Dictionary, strExportDir As String, strFileName As String)
'
'    On Error GoTo ErrorHandler

'    If Not (pExportList Is Nothing) Then
'        Dim pFileList As Collection
'        If (pExportList.Exists(strExportDir) = False) Then
'            Set pFileList = New Collection
'        Else
'            Set pFileList = pExportList.Item(strExportDir)
'            pExportList.Remove strExportDir
'        End If
'        pFileList.Add Item:=strFileName
'        pExportList.Add Key:=strExportDir, Item:=pFileList
'    End If
'
'    Exit Sub
'ErrorHandler:
'    HandleError True, c_strModuleName & ".ExportList " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub

Private Function TempDirectory(strExportDir As String) As String
'Get temp directory for exporting files
    On Error GoTo ErrorHandler

    Dim strSysTempDir As String, strTempDir
    Dim pFolder As Folder, blnValid As Boolean, intCount As Integer
    strSysTempDir = Environ$("TEMP")
    If (strSysTempDir = "") Then strSysTempDir = Environ$("TMP")
    If (strSysTempDir = "") Then
        If (m_pFSO.FolderExists("C:\Temp")) Then
            strSysTempDir = "C:\Temp"
        Else
            m_pFSO.CreateFolder "C:\Temp"
            If (m_pFSO.FolderExists("C:\Temp")) Then strSysTempDir = "C:\Temp"
        End If
    End If
    If (strSysTempDir = "") Then strSysTempDir = strExportDir
    blnValid = False
    intCount = 1
    strTempDir = strSysTempDir & "\Geology_Export"
    Do While Not blnValid
        If (m_pFSO.FolderExists(strTempDir)) Then
            Set pFolder = m_pFSO.GetFolder(strTempDir)
            If (pFolder.SubFolders.Count = 0) And (pFolder.Files.Count = 0) Then
                blnValid = True
            End If
        Else
            m_pFSO.CreateFolder strTempDir
            If (m_pFSO.FolderExists(strTempDir)) Then
                blnValid = True
            End If
        End If
        If (blnValid = False) Then
            strTempDir = strSysTempDir & "\Geology_Export" & Right("00" & CStr(intCount), 3)
            intCount = intCount + 1
            If (intCount > 999) Then Exit Do
        End If
    Loop

    TempDirectory = strTempDir
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".TempDirectory " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
