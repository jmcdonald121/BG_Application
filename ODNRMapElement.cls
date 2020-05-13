VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRMapElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRMapElement
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates a container to store default information
'               about about each text element on the layout.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pParent As IGraphicsContainer
Private m_pElement As IElement
Private m_strKey As String
Private m_blnIsLoaded As Boolean

Const c_strModuleName As String = "ODNRMapElement"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pParent = Nothing
    Set m_pElement = Nothing
    m_blnIsLoaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get ElementText() As String
    On Error GoTo ErrorHandler
    
    If m_blnIsLoaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = m_pElement
        ElementText = pTextEl.Text
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ElementText " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let ElementText(strText As String)
    On Error GoTo ErrorHandler
    
    If m_blnIsLoaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = m_pElement
        pTextEl.Text = strText
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ElementText " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get SymbolColor() As IRgbColor
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pTextEl As ITextElement, pTextSym As ITextSymbol
        Set pTextEl = m_pElement
        Set pTextSym = pTextEl.Symbol
        Set SymbolColor = pTextSym.Color
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".SymbolColor " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Set SymbolColor(pColor As IRgbColor)
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then
        Dim pTextEl As ITextElement, pTextSym As ITextSymbol
        Set pTextEl = m_pElement
        Set pTextSym = pTextEl.Symbol
        pTextSym.Color = pColor
        pTextEl.Symbol = pTextSym
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".SymbolColor " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get Parent() As IGraphicsContainer
    On Error GoTo ErrorHandler
    
    If (m_blnIsLoaded) Then Set Parent = m_pParent
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub SetTextToKey()
    On Error GoTo ErrorHandler
    
    If m_blnIsLoaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = m_pElement
        pTextEl.Text = m_strKey
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".SetTextToKey " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadElement(pElement As IElement, strKey As String, pParent As IGraphicsContainer)
    On Error GoTo ErrorHandler
    
    If (Not pElement Is Nothing) And (Not pParent Is Nothing) Then
        Set m_pElement = pElement
        Set m_pParent = pParent
        m_strKey = strKey
        m_blnIsLoaded = True
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadElement " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UpdateGraphicsContainer()
    On Error GoTo ErrorHandler
    
    If m_blnIsLoaded Then m_pParent.UpdateElement m_pElement
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UpdateGraphicsContainer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

