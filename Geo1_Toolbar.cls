VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEO1_Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
'   Program:    Geo1_Toolbar
'   Author:     Jeffrey M Laird
'   Date:       March 19, 2002
'   Purpose:    Creates the Geo1 Toolbar to support
'               phase 1 commands and tools
'
'   Called from:
'
'*****************************************


Option Explicit

Implements esriSystemUI.IToolBarDef

' Variables used by the Error handler function - DO NOT REMOVE
Const c_strModuleName As String = "GEO1_Toolbar"
 
Private Property Get IToolBarDef_ItemCount() As Long
    On Error GoTo ErrorHandler
    
    IToolBarDef_ItemCount = 11
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolBarDef_ItemCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
    On Error GoTo ErrorHandler
    
    Select Case pos
    Case 0
      itemDef.ID = "ODNR_Geo15.Select_DataDir_Cmd"
      itemDef.Group = False
    Case 1
      itemDef.ID = "ODNR_Geo15.Export_Image_Cmd"
      itemDef.Group = False
    Case 2
      itemDef.ID = "ODNR_Geo15.PickQuad_Tool"
      itemDef.Group = True
    Case 3
      itemDef.ID = "ODNR_Geo15.Select100_Tool"
      itemDef.Group = False
    Case 4
      itemDef.ID = "ODNR_Geo15.Quad_Combo"
      itemDef.Group = False
    Case 5
      itemDef.ID = "ODNR_Geo15.ZoomToOhio_Cmd"
      itemDef.Group = True
    Case 6
      itemDef.ID = "ODNR_Geo15.GoDataView_Cmd"
      itemDef.Group = False
    Case 7
      itemDef.ID = "ODNR_Geo15.GoLayout_Cmd"
      itemDef.Group = False
    Case 8
      itemDef.ID = "ODNR_Geo15.Export_Tool"
      itemDef.Group = False
    Case 9
      itemDef.ID = "ODNR_Geo15.Bedrock_Layers_Cmd"
      itemDef.Group = True
    Case 10
      itemDef.ID = "ODNR_Geo15.Switch_combo"
      itemDef.Group = True
    End Select
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolBarDef_GetItemInfo " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get IToolBarDef_Name() As String
    On Error GoTo ErrorHandler
    
    IToolBarDef_Name = "ODNR_Geology"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolBarDef_Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get IToolBarDef_Caption() As String
    On Error GoTo ErrorHandler
    
    IToolBarDef_Caption = "ODNR Geology Tools"
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IToolBarDef_Caption " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
