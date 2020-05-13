VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRQuads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRQuads
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that serves individual ODNRQuads.
'
'   Called from:
'
'*****************************************

Option Explicit

Private m_pFSO As FileSystemObject
Private m_pExtentEnv As IEnvelope

Private m_pApp As esriFramework.IApplication
Private m_pQuadList As Dictionary
Private m_pAdjacentQuadList As Dictionary
Private m_pQuadLayer As ILayer
Private m_pParent As ODNRProject
Private m_lngQuadIdx As Long
Private m_lngAdjacentQuadIdx As Long
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRQuads"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
35:     m_blnIsloaded = False
36:     m_lngQuadIdx = 0
37:     m_lngAdjacentQuadIdx = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
47:     Set m_pExtentEnv = Nothing
48:     Set m_pQuadList = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get QuadFeatureLayer() As ILayer
    On Error GoTo ErrorHandler
    
58:     If (m_blnIsloaded) Then Set QuadFeatureLayer = m_pQuadLayer
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadFeatureLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadCount() As Integer
    On Error GoTo ErrorHandler
    
68:     If (m_blnIsloaded) Then QuadCount = m_pQuadList.Count
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadScale() As ODNRQuadScale
    On Error GoTo ErrorHandler
    
78:     If (m_blnIsloaded) Then QuadScale = m_pParent.QuadScale
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadScale " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ExtentEnvelope() As IEnvelope
    On Error GoTo ErrorHandler
    
88:     If (m_blnIsloaded) Then Set ExtentEnvelope = m_pExtentEnv
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ExtentEnvelope " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Set ExtentEnvelope(RHS As IEnvelope)
    On Error GoTo ErrorHandler
    
98:     If (m_blnIsloaded) Then
99:         Set m_pExtentEnv = RHS
100:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ExtentEnvelope " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get FocusQuad() As ODNRQuad
    On Error GoTo ErrorHandler
    
110:     If (m_blnIsloaded) Then
111:         If Not (m_pQuadList Is Nothing) Then
            Dim vKey As Variant, pQuad As ODNRQuad
113:             For Each vKey In m_pQuadList.Keys
114:                 Set pQuad = m_pQuadList.Item(vKey)
115:                 If (pQuad.FocusQuad) Then
116:                     Set FocusQuad = pQuad
117:                     Exit For
118:                 End If
119:             Next
120:         End If
121:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".FocusQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get NextAdjacentQuad() As ODNRQuad
    On Error GoTo ErrorHandler
    
131:     If (m_blnIsloaded) Then
132:         If Not (m_pAdjacentQuadList Is Nothing) Then
133:             If Not (Me.FocusQuad Is Nothing) Then
134:                 If (m_lngAdjacentQuadIdx < m_pAdjacentQuadList.Count) Then
                    Dim vKey As Variant
136:                     vKey = m_pAdjacentQuadList.Keys(m_lngAdjacentQuadIdx)
137:                     Set NextAdjacentQuad = m_pAdjacentQuadList.Item(vKey)
138:                     m_lngAdjacentQuadIdx = m_lngAdjacentQuadIdx + 1
139:                 Else
140:                     Set NextAdjacentQuad = Nothing
141:                 End If
142:             End If
143:         End If
144:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".NextAdjacentQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub ResetAdjacentQuads()
    On Error GoTo ErrorHandler
    
154:     If (m_blnIsloaded) Then m_lngAdjacentQuadIdx = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ResetAdjacentQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub SetFocusQuad(lngQuadId As Long)
    On Error GoTo ErrorHandler
    
164:     If (m_blnIsloaded) Then
165:         If Not (m_pQuadList Is Nothing) Then
            Dim vKey As Variant, pQuad As ODNRQuad
            Dim pPolygon As IPolygon, pArea As IArea, pTrans2D As ITransform2D
            Dim pQuadList As Collection, pKey As Variant
169:             For Each vKey In m_pQuadList.Keys
170:                 Set pQuad = m_pQuadList.Item(vKey)
171:                 pQuad.FocusQuad = False
172:                 If (pQuad.QuadId = lngQuadId) Then
173:                     pQuad.FocusQuad = True
174:                     Set pPolygon = pQuad.QuadBoundary
175:                     Set pArea = pPolygon
176:                     Set pTrans2D = pPolygon
177:                     With pTrans2D
178:                         .Scale pArea.Centroid, 1.1, 1.1
179:                     End With
180:                     Set m_pAdjacentQuadList = New Dictionary
181:                     m_lngAdjacentQuadIdx = 0
182:                     Set pQuadList = Me.FindQuadIdsByGeometry(pPolygon)
183:                     For Each pKey In pQuadList
184:                         If (lngQuadId <> CLng(pKey)) Then
185:                             Set pQuad = New ODNRQuad
186:                             pQuad.LoadQuad Me, CLng(pKey), m_pParent.QuadScale
187:                             m_pAdjacentQuadList.Add Key:=CLng(pKey), Item:=pQuad
188:                         End If
189:                     Next
190:                 End If
191:             Next
192:         End If
193:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".SetFocusQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub NoFocusQuad()
    On Error GoTo ErrorHandler
    
203:     If (m_blnIsloaded) Then
204:         If Not (m_pQuadList Is Nothing) Then
            Dim vKey As Variant, pQuad As ODNRQuad
206:             For Each vKey In m_pQuadList.Keys
207:                 Set pQuad = m_pQuadList.Item(vKey)
208:                 pQuad.FocusQuad = False
209:             Next
210:             Set m_pAdjacentQuadList = Nothing
211:             m_lngAdjacentQuadIdx = 0
212:         End If
213:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".NoFocusQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function QuadById(lngQuadId As Long) As ODNRQuad
    On Error GoTo ErrorHandler
    
223:     If (m_blnIsloaded) Then
224:         If Not (m_pQuadList Is Nothing) Then
225:             If (m_pQuadList.Exists(lngQuadId)) Then
226:                 Set QuadById = m_pQuadList.Item(lngQuadId)
227:             Else
228:                 Set QuadById = Nothing
229:             End If
230:         End If
231:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadById " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function QuadByName(strQuadName As String) As ODNRQuad
    On Error GoTo ErrorHandler
    
241:     If (m_blnIsloaded) Then
242:         If Not (m_pQuadList Is Nothing) Then
            Dim vKey As Variant, pODNRQuad As ODNRQuad
244:             For Each vKey In m_pQuadList.Keys
245:                 Set pODNRQuad = m_pQuadList.Item(vKey)
246:                 If (pODNRQuad.QuadName = strQuadName) Then
247:                     Set QuadByName = pODNRQuad
248:                     Exit For
249:                 End If
250:             Next
251:         End If
252:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadByName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function NextQuad() As ODNRQuad
    On Error GoTo ErrorHandler
    
262:     If (m_blnIsloaded) Then
263:         If Not (m_pQuadList Is Nothing) Then
264:             If (m_lngQuadIdx < m_pQuadList.Count) Then
                Dim vKey As Variant
266:                 vKey = m_pQuadList.Keys(m_lngQuadIdx)
267:                 Set NextQuad = m_pQuadList.Item(vKey)
268:                 m_lngQuadIdx = m_lngQuadIdx + 1
269:             Else
270:                 Set NextQuad = Nothing
271:             End If
272:         End If
273:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".NextQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub Reset()
    On Error GoTo ErrorHandler
    
283:     If (m_blnIsloaded) Then m_lngQuadIdx = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Reset " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub AddQuadById(lngQuadId As Long)
    On Error GoTo ErrorHandler
    
293:     If (m_blnIsloaded) Then
294:         If (m_pQuadList Is Nothing) Then Set m_pQuadList = New Dictionary
295:         If (m_pQuadList.Exists(lngQuadId) = False) Then
            Dim pODNRQuad As ODNRQuad
297:             Set pODNRQuad = New ODNRQuad
298:             pODNRQuad.LoadQuad Me, lngQuadId, m_pParent.QuadScale
299:             m_pQuadList.Add Key:=lngQuadId, Item:=pODNRQuad
300:         End If
301:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".AddQuadById " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub AddQuadByName(strQuadName As String)
    On Error GoTo ErrorHandler
    
311:     If (m_blnIsloaded) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pQf As IQueryFilter
        Dim pF As IFeature, pFCursor As IFeatureCursor
        Dim lngFCount As Long, lngFId As Long
315:         Set pFlyr = m_pQuadLayer
316:         Set pFc = pFlyr.FeatureClass
317:         Set pQf = New QueryFilter
318:         If (m_pParent.QuadScale = odnr24K) Then
319:             pQf.WhereClause = "QUADNAME = '" & strQuadName & "'"
320:         ElseIf (m_pParent.QuadScale = odnr100K) Then
321:             pQf.WhereClause = "NAME = '" & strQuadName & "'"
322:         End If
323:         lngFCount = pFc.FeatureCount(pQf)

325:         If (lngFCount >= 1) Then
326:             Set pFCursor = pFc.Search(pQf, False)
327:             Set pF = pFCursor.NextFeature
328:             Do While Not pF Is Nothing
329:                 lngFId = pF.OID
330:                 Me.AddQuadById lngFId
331:                 Set pF = pFCursor.NextFeature
332:             Loop
333:         End If
334:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".AddQuadByName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub AddQuadsByGeometry(pG As IGeometry)
    On Error GoTo ErrorHandler
    
344:     If (m_blnIsloaded) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pSpF As ISpatialFilter
        Dim pF As IFeature, pFCursor As IFeatureCursor
        Dim lngFCount As Long, lngFId As Long
        
349:         Set pFlyr = m_pQuadLayer
350:         Set pFc = pFlyr.FeatureClass
351:         Set pSpF = New SpatialFilter
352:         Set pSpF.Geometry = pG
353:         pSpF.GeometryField = "Shape"
354:         pSpF.SpatialRel = esriSpatialRelIntersects
355:         lngFCount = pFc.FeatureCount(pSpF)

357:         If (lngFCount >= 1) Then
358:             Set pFCursor = pFc.Search(pSpF, False)
359:             Set pF = pFCursor.NextFeature
360:             Do While Not pF Is Nothing
361:                 lngFId = pF.OID
362:                 Me.AddQuadById lngFId
363:                 Set pF = pFCursor.NextFeature
364:             Loop
365:         End If
366:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".AddQuadsByGeometry " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function FindQuadIdsByGeometry(pG As IGeometry) As Collection
    On Error GoTo ErrorHandler

376:     If (m_blnIsloaded) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pSpF As ISpatialFilter
        Dim pF As IFeature, pFCursor As IFeatureCursor
        Dim lngFCount As Long, lngFId As Long, pQuadList As Collection
380:         Set pFlyr = m_pQuadLayer
381:         Set pFc = pFlyr.FeatureClass
382:         Set pSpF = New SpatialFilter
383:         Set pSpF.Geometry = pG
384:         pSpF.GeometryField = "Shape"
385:         pSpF.SpatialRel = esriSpatialRelIntersects
386:         lngFCount = pFc.FeatureCount(pSpF)
387:         Set pQuadList = New Collection
388:         If (lngFCount >= 1) Then
389:             Set pFCursor = pFc.Search(pSpF, False)
390:             Set pF = pFCursor.NextFeature
391:             Do While Not pF Is Nothing
392:                 lngFId = pF.OID
393:                 pQuadList.Add Item:=lngFId
394:                 Set pF = pFCursor.NextFeature
395:             Loop
396:             Set FindQuadIdsByGeometry = pQuadList
397:         End If
398:     End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".FindQuadIdsByGeometry " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub RemoveQuadById(lngQuadId As Long)
    On Error GoTo ErrorHandler
    
408:     If (m_blnIsloaded) Then
409:         If Not (m_pQuadList Is Nothing) Then
            Dim pMxDoc As IMxDocument, pQuad As ODNRQuad, blnRefresh As Boolean
411:             If (m_pQuadList.Exists(lngQuadId)) Then
412:                 Set pQuad = m_pQuadList.Item(lngQuadId)
413:                 If (pQuad.FocusQuad) Then
414:                     Set m_pAdjacentQuadList = Nothing
415:                     m_lngAdjacentQuadIdx = 0
416:                 End If
417:                 m_pQuadList.Remove lngQuadId
418:                 blnRefresh = pQuad.IsHighlighted
419:                 Set pMxDoc = m_pApp.Document
420:                 Set pQuad = Nothing
421:                 If (blnRefresh) Then pMxDoc.ActiveView.Refresh
422:             End If
423:         End If
424:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".RemoveQuadById " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub RemoveAll()
    On Error GoTo ErrorHandler
        
434:     If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument
        Dim pQuad As ODNRQuad, vKey As Variant, blnRefresh As Boolean
437:         For Each vKey In m_pQuadList.Keys
438:             Set pQuad = m_pQuadList.Item(vKey)
439:             If (blnRefresh = False) And (pQuad.IsHighlighted) Then blnRefresh = True
440:             m_pQuadList.Remove vKey
441:             Set pQuad = Nothing
442:         Next
443:         Set m_pAdjacentQuadList = Nothing
444:         m_lngAdjacentQuadIdx = 0
445:         Set pMxDoc = m_pApp.Document
446:         If (blnRefresh) Then pMxDoc.ActiveView.Refresh
447:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".RemoveAll " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub HighlightQuads()
    On Error GoTo ErrorHandler
    
457:     If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pQuad As ODNRQuad, vKey As Variant
        Dim pFlyr As IFeatureLayer, pFSEl As IFeatureSelection, pSelSet As ISelectionSet
        Dim pEnumIDs As IEnumIDs, lngId As Long, pSelList As Dictionary
461:         Set pMxDoc = m_pApp.Document
462:         Set pFlyr = m_pQuadLayer
463:         Set pFSEl = pFlyr
464:         Set pSelSet = pFSEl.SelectionSet
465:         Set pEnumIDs = pSelSet.IDs
466:         lngId = pEnumIDs.Next
467:         Set pSelList = New Dictionary
468:         Do While lngId > -1
469:             pSelList.Add Key:=lngId, Item:=""
470:             lngId = pEnumIDs.Next
471:         Loop
472:         If (m_pQuadList.Count > 0) Then
473:             For Each vKey In m_pQuadList.Keys
474:                 Set pQuad = m_pQuadList.Item(vKey)
475:                 If (pSelList.Exists(pQuad.QuadId) = False) Then
476:                     pSelSet.Add pQuad.QuadId
477:                 End If
478:             Next
479:         End If
480:         pSelSet.MakePermanent
481:         Set pFSEl.SelectionSet = pSelSet
482:         pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
483:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".HighlightQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub UnHighlightQuads()
    On Error GoTo ErrorHandler
    
493:     If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument
        Dim pFlyr As IFeatureLayer, pFSEl As IFeatureSelection
496:         Set pMxDoc = m_pApp.Document
497:         Set pFlyr = m_pQuadLayer
498:         Set pFSEl = pFlyr
499:         pFSEl.Clear
500:         pMxDoc.ActiveView.Refresh
501:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".UnHighlightQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get Parent() As ODNRProject
    On Error GoTo ErrorHandler
    
511:     If Not (m_pParent Is Nothing) Then Set Parent = m_pParent
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub Load(pApp As esriFramework.IApplication, pParent As ODNRProject)
    On Error GoTo ErrorHandler
    
521:     If (Not pApp Is Nothing) And (Not pParent Is Nothing) Then
522:         Set m_pApp = pApp
523:         Set m_pParent = pParent
524:         Set m_pQuadLayer = m_pParent.QuadFeatureLayer
525:         Set m_pQuadList = New Dictionary
526:         m_blnIsloaded = True
527:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get QuadNameList() As Collection
    On Error GoTo ErrorHandler
536:     If (m_blnIsloaded) Then
        Dim pFWs As IFeatureWorkspace, pFlyr As IFeatureLayer, pFc As IFeatureClass
        Dim pF As IFeature, pFCursor As IFeatureCursor
        Dim lngFieldIdx As Long, pQuadList As Collection
540:         Set pFlyr = m_pQuadLayer
541:         If (pFlyr Is Nothing) Then
542:             If (g_strSDEPGB = "PGB") Then
543:                 Set pFWs = ODNR_Common.ODNRGeologyWorkspace
544:             ElseIf (g_strSDEPGB = "SDE") Then
545:                 Set pFWs = ODNR_Common.ODNRBasemapWorkspace
546:             End If
547:             If (Me.QuadScale = odnr24K) Then
548:                 Set pFc = pFWs.OpenFeatureClass("Quad24K")
549:             ElseIf (Me.QuadScale = odnr100K) Then
550:                 Set pFc = pFWs.OpenFeatureClass("Quad100K")
551:             End If
552:         Else
553:             Set pFc = pFlyr.FeatureClass
554:         End If
555:         If (m_pParent.QuadScale = odnr24K) Then
556:             lngFieldIdx = pFc.FindField("QUADNAME")
557:         ElseIf (m_pParent.QuadScale = odnr100K) Then
558:             lngFieldIdx = pFc.FindField("NAME")
559:         End If
560:         Set pQuadList = New Collection
561:         Set pFCursor = pFc.Search(Nothing, False)
562:         Set pF = pFCursor.NextFeature
563:         Do While Not pF Is Nothing
564:             pQuadList.Add Item:=pF.Value(lngFieldIdx)
565:             Set pF = pFCursor.NextFeature
566:         Loop
567:         Set QuadNameList = pQuadList
568:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadNameList " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadLayoutTextRecord() As IRow
    On Error GoTo ErrorHandler
    
578:     If (m_blnIsloaded) And (Not Me.FocusQuad Is Nothing) Then
        Dim pFWs As IFeatureWorkspace
        Dim pTable As ITable, pCursor As ICursor, pRow As IRow
        Dim pQf As IQueryFilter, pQryStr As String, pQuad As ODNRQuad
        Dim strQuadName As String, intQuadCount As Integer, strType As String
        Dim pQuadLayer As ODNRQuadLayer, pBedrockLayer As ODNRBedrockLayer
        Dim pGlyr As IGroupLayer, pClyr As ICompositeLayer
        Dim lngLayerIdx As Long, pLayer As ILayer, strBedrockType As String
        
587:         Set pQuad = Me.FocusQuad
588:         If Not (pQuad Is Nothing) Then
589:             strQuadName = pQuad.QuadName
590:             Set pFWs = ODNR_Common.ODNRGeologyWorkspace
591:             Set pQf = New QueryFilter
592:             If (m_pParent.ProjectType = odnrGeology) Then
593:                 strType = "BG"
594:             ElseIf (m_pParent.ProjectType = odnrTopography) Then
595:                 strType = "BT"
596:             ElseIf (m_pParent.ProjectType = odnrBedrockStructure) Then
597:                 strType = "BS"
598:             ElseIf (m_pParent.ProjectType = odnrAUM) Then
599:                 strType = "AUM"
600:             End If
601:             If (Me.QuadScale = odnr24K) Then
602:                 If (Me.Parent.ProjectType = odnrBedrockStructure) Then
603:                     Set pBedrockLayer = Me.Parent.BedrockLayers.GetLayerByName("BS Contours")
604:                     Set pGlyr = pBedrockLayer.ESRILayer
605:                     Set pClyr = pGlyr
606:                     For lngLayerIdx = 0 To pClyr.Count - 1
607:                         Set pLayer = pClyr.Layer(lngLayerIdx)
608:                         If (pLayer.Visible) Then
609:                             strBedrockType = Left(pLayer.Name, Len(pLayer.Name) - 8)
610:                             Exit For
611:                         End If
612:                     Next
613:                     If (strBedrockType <> "") Then
614:                         Set pTable = pFWs.OpenTable("TITLE24K_NEW2")
615:                         pQryStr = "QUAD = '" & strQuadName & "' And PROJECTTYPE = '" & strType & "' And TYPECODE = '" & strBedrockType & "'"
616:                     End If
617:                 Else
618:                     Set pTable = pFWs.OpenTable("TITLE24K_NEW2")
619:                     pQryStr = "QUAD = '" & strQuadName & "' And PROJECTTYPE = '" & strType & "'"
620:                 End If
621:             ElseIf (Me.QuadScale = odnr100K) Then
622:                 Set pTable = pFWs.OpenTable("TITLE100K")
623:                 pQryStr = "QUAD = '" & strQuadName & "' And TYPE = '" & strType & "'"
624:             End If
625:             If (pQryStr <> "") Then
626:                 pQf.WhereClause = pQryStr
627:                 If (pTable.RowCount(pQf) > 0) Then
628:                     Set pCursor = pTable.Search(pQf, False)
629:                     Set pRow = pCursor.NextRow
630:                     Set QuadLayoutTextRecord = pRow
631:                 Else
632:                     Set QuadLayoutTextRecord = Nothing
633:                 End If
634:             Else
635:                 Set QuadLayoutTextRecord = Nothing
636:             End If
637:         Else
638:             Set QuadLayoutTextRecord = Nothing
639:         End If
640:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadLayoutTextRecord " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property


Public Function CountyList() As Collection
    On Error GoTo ErrorHandler
    
651:     If (m_blnIsloaded) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
        Dim pF As IFeature, pFCursor As IFeatureCursor, lngFieldIdx As Long
        Dim pSpF As ISpatialFilter, pGeoCol As IGeometryCollection
        Dim pGeoCol2 As IGeometryCollection
        Dim pPolygon As IPolygon, pGeo As IGeometry
        Dim pSegCol As ISegmentCollection
        Dim vKey As Variant, pQuad As ODNRQuad, pCountyList As Collection
659:         Set pFlyr = m_pParent.CountyLayer
660:         If Not (pFlyr Is Nothing) Then
661:             Set pFc = pFlyr.FeatureClass
662:             lngFieldIdx = pFc.FindField("COUNTY")
            
664:             Set pSegCol = New Ring
665:             Set pGeoCol = New Polygon
666:             For Each vKey In m_pQuadList.Keys
667:                 Set pQuad = m_pQuadList.Item(vKey)
668:                 Set pGeoCol2 = pQuad.QuadBoundary
669:                 pGeoCol.AddGeometryCollection pGeoCol2
670:             Next
    
672:             Set pPolygon = pGeoCol
673:             Set pGeo = pPolygon
674:             Set pSpF = New SpatialFilter
675:             Set pSpF.Geometry = pGeo
676:             pSpF.GeometryField = "Shape"
677:             pSpF.SpatialRel = esriSpatialRelIntersects
            
679:             Set pCountyList = New Collection
680:             Set pFCursor = pFc.Search(pSpF, False)
681:             Set pF = pFCursor.NextFeature
682:             Do While Not pF Is Nothing
683:                 pCountyList.Add Item:=pF.Value(lngFieldIdx)
684:                 Set pF = pFCursor.NextFeature
685:             Loop
686:             Set CountyList = pCountyList
687:         Else
688:             Set CountyList = Nothing
689:         End If
690:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".CountyList " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


