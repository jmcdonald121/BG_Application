VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRLayoutLegendElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRLayoutLegendElement
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
Private m_dblPatchHeight As Double
Private m_dblPatchWidth As Double
Private m_intLegendFontSize As Integer
Private m_strLegendFontName As String
Private m_pLegendItems As Collection
Private m_strKey As String
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRLayoutLegendElement"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    Set m_pLegendItems = New Collection
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_pParent = Nothing
    Set m_pElement = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get LegendTitleText() As String
    On Error GoTo ErrorHandler
    
    If m_blnIsloaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = GetTextElement
        LegendTitleText = pTextEl.Text
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ElementText " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let LegendTitleText(strText As String)
    On Error GoTo ErrorHandler
    
    If m_blnIsloaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = GetTextElement
        pTextEl.Text = strText
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ElementText " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get SymbolColor() As IRgbColor
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pFSEl As IFillShapeElement, pFillSym As IFillSymbol
        Set pFSEl = GetBorderElement
        Set pFillSym = pFSEl.Symbol
        Set SymbolColor = pFillSym.Color
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".SymbolColor " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Set SymbolColor(pColor As IRgbColor)
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pFSEl As IFillShapeElement, pFillSym As IFillSymbol
        Set pFSEl = GetBorderElement
        Set pFillSym = pFSEl.Symbol
        pFillSym.Color = pColor
        pFSEl.Symbol = pFillSym
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".SymbolColor " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get LegendElement() As IElement
    If (m_blnIsloaded) Then Set LegendElement = GetLegendElement
End Property

Public Sub AddLegendItem(pSym As ISymbol, strText As String)
    If (m_blnIsloaded) Then
        Dim pList As Collection
        Set pList = New Collection
        pList.Add Item:=pSym
        pList.Add Item:=strText
        m_pLegendItems.Add Item:=pList
    End If
End Sub

Public Sub BuildLegend()
    If (m_blnIsloaded) Then
        Dim pLegendEl As IElement, pLegendGroupEl As IGroupElement, pGroupEl As IGroupElement
        Dim pBorderEl As IElement, pBorderEnv As IEnvelope
        Dim pLegendTitleEl As IElement, pLegendTitlePt As IPoint, pLegendTitleEnv As IEnvelope, pLegendTitlePolygon As IPolygon
        Dim lngIdx As Long, pList As Collection
        Dim pSym As ISymbol, strText As String
        Dim pPatchEl As IElement, pTextEl As IElement
        Dim pPatchEnv As IEnvelope
        Dim pPageLayout As IPageLayout, pActiveView As IActiveView
        Dim dblXMax As Double, dblXMin As Double
        Dim dblYMax As Double, dblYMin As Double
        Dim dblXGap As Double, dblYGap As Double
        Dim dblPatchHeight As Double, dblHeight As Double
        Dim lngItemCount As Long
        Dim dblExcessHeight As Double, dblNeededHeight As Double
        Dim pPatchStartPt As IPoint, pTextStartPt As IPoint
        Set pBorderEl = GetBorderElement
        Set pBorderEnv = pBorderEl.Geometry.Envelope
        Set pLegendTitleEl = GetTextElement
'        Set pLegendTitlePt = pLegendTitleEl.Geometry
        Set pPageLayout = m_pParent
        Set pActiveView = pPageLayout
        Set pLegendTitlePolygon = New Polygon
        pLegendTitleEl.QueryOutline pActiveView.ScreenDisplay, pLegendTitlePolygon
        Set pLegendTitleEnv = pLegendTitlePolygon.Envelope
        dblYMax = pLegendTitleEnv.YMin
        dblYMin = pBorderEnv.YMin
        dblXMin = pBorderEnv.XMin
        dblXMax = pBorderEnv.XMax
        dblXGap = 0.3
        dblYGap = 0.2
        dblHeight = dblYMax - dblYMin
        lngItemCount = m_pLegendItems.Count
        dblPatchHeight = m_dblPatchHeight
        If (lngItemCount > 0) Then
            Set pLegendEl = New GroupElement
            Set pLegendGroupEl = pLegendEl
            dblNeededHeight = lngItemCount * dblPatchHeight + lngItemCount * dblYGap + dblYGap
            dblExcessHeight = dblNeededHeight - dblHeight
            If (dblExcessHeight > 0) Then
                dblPatchHeight = dblPatchHeight - dblExcessHeight / lngItemCount
            End If
            Set pPatchStartPt = New Point
            pPatchStartPt.X = dblXMin + dblXGap
            pPatchStartPt.Y = dblYMax - dblYGap
            Set pTextStartPt = New Point
            pTextStartPt.X = m_dblPatchWidth + dblXMin + dblXGap * 2
            pTextStartPt.Y = dblYMax - dblYGap - dblPatchHeight / 2
            For lngIdx = 1 To m_pLegendItems.Count
                Set pList = m_pLegendItems.Item(lngIdx)
                Set pSym = pList.Item(1)
                strText = pList.Item(2)
                Set pPatchEnv = New Envelope
                pPatchEnv.XMin = pPatchStartPt.X
                pPatchEnv.YMax = pPatchStartPt.Y
                pPatchEnv.XMax = pPatchStartPt.X + m_dblPatchWidth
                pPatchEnv.YMin = pPatchStartPt.Y - dblPatchHeight
                'Set pPatchEl = GetTextFrame(pPatchEnv)
                Set pPatchEl = BuildLegendPatch(pPatchEnv, pSym)
                pPatchEl.Activate pActiveView.ScreenDisplay
                Set pTextEl = BuildTextElement(pTextStartPt, strText, (dblXMax - (dblXMin + m_dblPatchWidth + dblXGap * 2)))
                pTextEl.Activate pActiveView.ScreenDisplay
                pLegendGroupEl.AddElement pPatchEl
                pLegendGroupEl.AddElement pTextEl
                pPatchStartPt.Y = pPatchStartPt.Y - dblPatchHeight - dblYGap
                pTextStartPt.Y = pTextStartPt.Y - dblPatchHeight - dblYGap
            Next
            Set pGroupEl = m_pElement
            pGroupEl.AddElement pLegendGroupEl
        End If
    End If
End Sub

Public Sub ClearLegend()
    If (m_blnIsloaded) Then
        Dim pLegendElement As IElement, pGroupElement As IGroupElement
        Set pLegendElement = GetLegendElement
        If (Not pLegendElement Is Nothing) Then
            Set pGroupElement = m_pElement
            pGroupElement.DeleteElement pLegendElement
        End If
    End If
End Sub

Public Property Get Parent() As IGraphicsContainer
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Set Parent = m_pParent
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub SetLegendTitleTextToKey()
    On Error GoTo ErrorHandler
    
    If m_blnIsloaded Then
        Dim pTextEl As ITextElement
        Set pTextEl = GetTextElement
        pTextEl.Text = m_strKey
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".SetTextToKey " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadElement(pElement As IElement, strKey As String, dblPatchHeight As Double, dblPatchWidth As Double, intFontSize As Integer, strFontName As String, pParent As IGraphicsContainer)
    On Error GoTo ErrorHandler
    
    If (Not pElement Is Nothing) And (Not pParent Is Nothing) Then
        Set m_pElement = pElement
        m_dblPatchHeight = dblPatchHeight
        m_dblPatchWidth = dblPatchWidth
        m_intLegendFontSize = intFontSize
        m_strLegendFontName = strFontName
        Set m_pParent = pParent
        m_strKey = strKey
        m_blnIsloaded = True
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadElement " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UpdateGraphicsContainer()
    On Error GoTo ErrorHandler
    
    If m_blnIsloaded Then m_pParent.UpdateElement m_pElement
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UpdateGraphicsContainer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function GetTextElement() As ITextElement
    Dim pGroupElement As IGroupElement
    Set pGroupElement = m_pElement
    If (TypeOf pGroupElement.Element(0) Is ITextElement) Then
        Set GetTextElement = pGroupElement.Element(0)
    ElseIf (TypeOf pGroupElement.Element(1) Is ITextElement) Then
        Set GetTextElement = pGroupElement.Element(1)
    End If
End Function

Private Function GetBorderElement() As IRectangleElement
    Dim pGroupElement As IGroupElement
    Set pGroupElement = m_pElement
    If (TypeOf pGroupElement.Element(0) Is IRectangleElement) Then
        Set GetBorderElement = pGroupElement.Element(0)
    ElseIf (TypeOf pGroupElement.Element(1) Is IRectangleElement) Then
        Set GetBorderElement = pGroupElement.Element(1)
    End If
End Function

Private Function GetLegendElement() As IGroupElement
    Dim pGroupElement As IGroupElement
    Set pGroupElement = m_pElement
    If (pGroupElement.ElementCount = 3) Then
        If (TypeOf pGroupElement.Element(0) Is IGroupElement) Then
            Set GetLegendElement = pGroupElement.Element(0)
        ElseIf (TypeOf pGroupElement.Element(1) Is IGroupElement) Then
            Set GetLegendElement = pGroupElement.Element(1)
        ElseIf (TypeOf pGroupElement.Element(2) Is IGroupElement) Then
            Set GetLegendElement = pGroupElement.Element(2)
        End If
    Else
        Set GetLegendElement = Nothing
    End If
End Function

Private Function BuildTextElement(pAnchorPt As IPoint, strText As String, dblMaxWidth As Double) As ITextElement
    Dim pEl As IElement, pCount As Long, pTextEl As ITextElement
    Dim pTextSym As IFormattedTextSymbol, pSimTextSym As ISimpleTextSymbol
    Dim pRGBColor As IRgbColor, pFontDisp As IFontDisp
    Dim varParsedText As Variant, vKey As Variant
    Dim strElementText As String, strElementTextWithNext As String
    Dim pPageLayout As IPageLayout, pActiveView As IActiveView, pEnv As IEnvelope
    Set pPageLayout = m_pParent
    Set pActiveView = pPageLayout
    Set pEnv = New Envelope
    Set pTextSym = New TextSymbol
    Set pTextEl = New TextElement
    Set pEl = pTextEl
    pEl.Geometry = pAnchorPt
    Set pRGBColor = New RgbColor
    pRGBColor.Red = 0
    pRGBColor.Blue = 0
    pRGBColor.Green = 0
    pTextSym.Color = pRGBColor
    pTextSym.HorizontalAlignment = esriTHALeft
    pTextSym.VerticalAlignment = esriTVACenter
    pTextSym.Size = m_intLegendFontSize
    Set pFontDisp = New StdFont
    pFontDisp.Name = m_strLegendFontName
    pFontDisp.Size = m_intLegendFontSize
    pFontDisp.Bold = False
    pTextSym.Font = pFontDisp
    pTextEl.Symbol = pTextSym
    varParsedText = Strings.Split(strText, " ")
    For Each vKey In varParsedText
        If (strElementText = "") Then
            strElementText = CStr(vKey)
        Else
            strElementText = pTextEl.Text
            pTextEl.Text = strElementText & " " & CStr(vKey)
            pEl.QueryBounds pActiveView.ScreenDisplay, pEnv
            If ((pEnv.XMax - pEnv.XMin) > dblMaxWidth) Then
                strElementText = strElementText & vbCrLf & CStr(vKey)
            Else
                strElementText = strElementText & " " & CStr(vKey)
            End If
        End If
        pTextEl.Text = strElementText
    Next
    Set BuildTextElement = pTextEl
End Function

Private Function BuildLegendPatch(pEnv As IEnvelope, pSym As ISymbol) As IElement
    Dim pSimFillSym As ISimpleFillSymbol, pFShapeEl As IFillShapeElement
    Dim pLineSym As ILineSymbol, pEl As IElement
    Dim pRGBColor As IRgbColor, pPolygon As IPolygon, pPtCol As IPointCollection
    Dim pFSym As IFillSymbol
    'Set pSimFillSym = New SimpleFillSymbol
    Set pFShapeEl = New RectangleElement
    Set pEl = pFShapeEl
    'pSimFillSym.Style = esriSFSSolid
    Set pLineSym = New SimpleLineSymbol
    Set pRGBColor = New RgbColor
    pRGBColor.Red = 0
    pRGBColor.Green = 0
    pRGBColor.Blue = 0
    pLineSym.Color = pRGBColor
    pLineSym.Width = 1
    'pSimFillSym.Outline = pLineSym
    Set pFSym = pSym
    pFShapeEl.Symbol = pFSym ' pSimFillSym
    Set pPtCol = New Polygon
    pPtCol.AddPoint pEnv.UpperLeft
    pPtCol.AddPoint pEnv.UpperRight
    pPtCol.AddPoint pEnv.LowerRight
    pPtCol.AddPoint pEnv.LowerLeft
    pPtCol.AddPoint pEnv.UpperLeft
    Set pPolygon = pPtCol
    pEl.Geometry = pPolygon
    Set BuildLegendPatch = pEl
End Function
