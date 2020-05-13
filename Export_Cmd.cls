VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Export_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************
'
'   Program:    EXPORT_CMD
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Creates the Export Button on the Toolbar.
'               Selects BT and/or BG data within the quad and
'               Exports it to shapefiles
'   Called from:
'
'*****************************************
Option Explicit

Private m_pApp As IApplication
Private m_pBitmap1 As IPictureDisp   'Bitmaps for the commands

Implements ICommand
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "D:\ODNR_Geology\VB_Projects\Export_Cmd.cls"

 
Private Sub Class_Initialize()
  On Error GoTo ErrorHandler



27:     Set m_pBitmap1 = LoadResPicture("EXPORT", 0)






  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler


43:     Set m_pApp = Nothing




  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler


    
58:     ICommand_Enabled = True
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler


    
    ' TODO: Add your implementation here
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


    
90:     ICommand_Name = "OBDR_Geo1.Export_Cmd"
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


    
106:     ICommand_Caption = "Export Geology to Shapefile"
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


    
122:     ICommand_Tooltip = "Export Geology to Shapefile"
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler


    
138:     ICommand_Message = "Export Geology to Shapefile"
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler


    
    ' TODO: Add your implementation here
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler


    
    ' TODO: Add your implementation here
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Bitmap() As esriCore.OLE_HANDLE
  On Error GoTo ErrorHandler


    
186:     ICommand_Bitmap = m_pBitmap1
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


    
202:     ICommand_Category = "ONDRGeo_MiscTools"
    




  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler


    
218:     Set m_pApp = hook
    




  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    Dim pMxDoc As IMxDocument
    Dim pDoc As IDocument
    Dim pMap As IMap
    Dim pEnumFeature As IEnumFeature
    Dim pFeature As IFeature
    Dim pFeatureLayer As IFeatureLayer
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pMapCollection As IMaps
    Dim pMapNDX As Long
    Dim g As Long
    Dim K As Integer
    Dim MapName As String
    Dim Projname As String
    Dim Lstring As String
    Dim Got24 As Boolean
    Dim Got100 As Boolean
    
    Got24 = False
    Got100 = False
    
    Set pMxDoc = m_pApp.Document
    Set pDoc = m_pApp.Document
    Projname = pDoc.Title
    
    Set pMap = pMxDoc.FocusMap

'Check to see if 24K or 100K
'BT only right now
    If (Projname = "BT24K_Quad.mxd") Then
    
     Set pMapCollection = pMxDoc.Maps
     pMapNDX = pMapCollection.Count
     If pMapCollection.Count = 0 Then Exit Sub
'set up the selects and select stuff
    
      For g = 0 To pMapNDX - 1
        MapName = pMapCollection.Item(g).Name
         If MapName = "Geology Map24BT" Then
            Got24 = True
            Set pMap = pMapCollection.Item(g)
            Set pMxDoc.ActiveView = pMap
            pMap.ClearSelection
            Set pEnumLayer = pMap.Layers
            pEnumLayer.Reset
            Set pLayer = pEnumLayer.Next
                    'set all to unselectable
                    Do While Not pLayer Is Nothing
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = False
                    Set pLayer = pEnumLayer.Next
                    Loop
                    pEnumLayer.Reset
                    Do While Not pLayer Is Nothing
                      If pLayer.Name <> "Quad24K" Then
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = True
                      End If
                      If pLayer.Name <> "Bedrock_Topography" Then
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = True
                      End If
                      If pLayer.Name <> "Boreholes" Then
                            Set pFeatureLayer = pLayer
                            pFeatureLayer.Selectable = True
                      End If
                'add other export layers here
                    Set pLayer = pEnumLayer.Next
                    Loop
          End If
        Next g

'Project - if 24 only BT and Boreholes - 100 BT Boreholes and basedata


'Dim pExportOperation As IExportOperation

'Dim pFeatureClassName As IFeatureClassName
'Dim pDataset As IDataset
'Set pDataset = pFeatureclass
'Set pFeatureClassName = pDataset.FullName

'Dim pQueryFilter As IQueryFilter
'Set pQueryFilter = New QueryFilter

'Dim pSelectionSet As ISelectionSet

'Dim pGeometryDef As IGeometryDef
'Dim hwnd As Long
    




  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 


