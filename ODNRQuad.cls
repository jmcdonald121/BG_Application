VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRQuad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRQuad
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object representing a specified quadrangle
'               in the quadrangle feature class.  This object is owned
'               by the ODNRQuads object
'   Called from:
'
'*****************************************

Option Explicit

Private m_pQuadFlyr As IFeatureLayer
Private m_pQuadF As IFeature
Private m_pQuadScale As ODNRQuadScale
Private m_pParent As ODNRQuads
Private m_blnFocusQuad As Boolean
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRQuad"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    m_blnFocusQuad = False
    m_pQuadScale = odnrScaleUnknown
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Me.UnHighlightQuad
    Set m_pQuadF = Nothing
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get QuadId() As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then QuadId = m_pQuadF.OID
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadId " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadName() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (m_pQuadScale = odnr24K) Then
            QuadName = m_pQuadF.Value(m_pQuadF.Fields.FindField("QUADNAME"))
        ElseIf (m_pQuadScale = odnr100K) Then
            QuadName = m_pQuadF.Value(m_pQuadF.Fields.FindField("NAME"))
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ODNRQuadId() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (m_pQuadScale = odnr24K) Then
            ODNRQuadId = m_pQuadF.Value(m_pQuadF.Fields.FindField("OGSCODE"))
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ODNRQuadId " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadBoundary() As IGeometry
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Set QuadBoundary = m_pQuadF.Shape
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadBoundary " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadScale() As ODNRQuadScale
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then QuadScale = m_pQuadScale
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadScale " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get FocusQuad() As Boolean
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then FocusQuad = m_blnFocusQuad
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".FocusQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let FocusQuad(RHS As Boolean)
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        m_blnFocusQuad = RHS
    End If
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".FocusQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get Parent() As ODNRQuads
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then Set Parent = m_pParent
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get IsHighlighted() As Boolean
    On Error GoTo ErrorHandler
    
    IsHighlighted = False
    If (m_blnIsloaded) Then
        Dim pFSEl As IFeatureSelection, pEnumIDs As IEnumIDs, lngId As Long
        Set pFSEl = m_pQuadFlyr
        Set pEnumIDs = pFSEl.SelectionSet.IDs
        lngId = pEnumIDs.Next
        Do While lngId > -1
            If (lngId = m_pQuadF.OID) Then
                IsHighlighted = True
                Exit Do
            End If
            lngId = pEnumIDs.Next
        Loop
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".IsHighlighted " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub HighlightQuad()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pFSEl As IFeatureSelection
        Set pFSEl = m_pQuadFlyr
        pFSEl.CombinationMethod = esriSelectionResultAdd
        pFSEl.Add m_pQuadF
        pFSEl.SelectionSet.MakePermanent
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".HighlightQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UnHighlightQuad()
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pFSEl As IFeatureSelection, lngOIDs(1) As Long
        lngOIDs(0) = m_pQuadF.OID
        Set pFSEl = m_pQuadFlyr
        pFSEl.SelectionSet.RemoveList 1, lngOIDs(0)
        'Added 20051217, Jim McDonald
'        MsgBox "m_blnIsloaded = " & m_blnIsloaded & vbCrLf & _
'        "lngOIDs(0) = " & lngOIDs(0) & vbCrLf & _
'        "m_pQuadFlyr.Name = " & m_pQuadFlyr.Name & vbCrLf & _
'        "pFSEl.SelectionSet.Count = " & pFSEl.SelectionSet.Count
        'End added section, 20051217, Jim McDonald
        'pFSEl.SelectionSet.MakePermanent
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UnHighlightQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub LoadQuad(pParent As ODNRQuads, ByVal lngId As Long, ByVal pQuadScale As ODNRQuadScale)
    On Error GoTo ErrorHandler
    If (0 < lngId) And (Not pParent Is Nothing) Then
        Dim pFc As IFeatureClass, pODNRProject As ODNRProject
        Set m_pParent = pParent
        m_pQuadScale = pQuadScale
        Set pODNRProject = m_pParent.Parent
        Set m_pQuadFlyr = pODNRProject.QuadFeatureLayer
        Set pFc = m_pQuadFlyr.FeatureClass
        Set m_pQuadF = pFc.GetFeature(lngId)
        m_blnIsloaded = True
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

