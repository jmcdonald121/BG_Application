VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRStateLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRStateLayers
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that serves individual ODNRStateLayers.
'
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pParent As ODNRProject
Private m_lngActiveMap As ODNRMapType
Private m_lngLayerIdx As Long
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRStateLayers"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
29:     m_blnIsloaded = False
30:     m_lngLayerIdx = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
40:     m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get Parent() As ODNRProject
    On Error GoTo ErrorHandler
    
50:     If (m_blnIsloaded) Then Set Parent = m_pParent

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Parent " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ActiveMapName() As String
    On Error GoTo ErrorHandler
    
60:     If (m_blnIsloaded) Then
        Dim pMapTypeDomain As ICodedValueDomain, lngIdx As Long
62:         Set pMapTypeDomain = gODNRProjectDb.MapTypeDomain
63:         For lngIdx = 0 To pMapTypeDomain.CodeCount - 1
64:             If (pMapTypeDomain.Value(lngIdx) = m_lngActiveMap) Then
65:                 ActiveMapName = pMapTypeDomain.Name(lngIdx)
66:                 Exit For
67:             End If
68:         Next
69:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMapName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let ActiveMapName(RHS As String)
    On Error GoTo ErrorHandler
    
79:     If (m_blnIsloaded) Then
        Dim pMapTypeDomain As ICodedValueDomain, lngIdx As Long
81:         Set pMapTypeDomain = gODNRProjectDb.MapTypeDomain
82:         For lngIdx = 0 To pMapTypeDomain.CodeCount - 1
83:             If (pMapTypeDomain.Name(lngIdx) = RHS) Then
84:                 m_lngActiveMap = pMapTypeDomain.Value(lngIdx)
85:                 m_lngLayerIdx = 0
86:                 Exit For
87:             End If
88:         Next
89:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMapName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ActiveMap() As ODNRMapType
    On Error GoTo ErrorHandler
    
99:     If (m_blnIsloaded) Then
100:         ActiveMap = m_lngActiveMap
101:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Let ActiveMap(RHS As ODNRMapType)
    On Error GoTo ErrorHandler
    
111:     If (m_blnIsloaded) Then
112:         If (RHS <> m_lngActiveMap) Then
113:             m_lngActiveMap = RHS
114:             m_lngLayerIdx = 0
115:         End If
116:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveMap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get LayerCount(lngMapType As ODNRMapType) As Long
    On Error GoTo ErrorHandler
    
126:     If (m_blnIsloaded) Then
        Dim pLayerList As Dictionary
128:         Set pLayerList = GetMapLayers(lngMapType)
129:         LayerCount = pLayerList.Count
130:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function GetLayerByName(strName As String) As ODNRStateLayer
    On Error GoTo ErrorHandler

140:     If (m_blnIsloaded) Then
        Dim pLayerList As Dictionary, vKey As Variant, pLayersTbl As ITable
        Dim pODNRLayer As ODNRStateLayer, pLayer As ILayer
        Dim str24KField As String, str100KField As String, pFieldNames As String
        Dim lngExportMethod As Long
        Dim lngId As Long, pRow As IRow, lngFieldIdx As Long
        Dim pMemBlobStrm As IMemoryBlobStream, strFilePath As String, strTempDir As String
        Dim strLayerType As String
        Dim strLayerTypeName As String
149:         Set pLayerList = GetMapLayers(m_lngActiveMap)
150:         Set pLayersTbl = gODNRProjectDb.LayerObjectsTable
151:         If (pLayerList.Exists(strName)) Then
152:             Set pODNRLayer = New ODNRStateLayer
153:             Set pLayer = FindLayer(strName)
154:             lngId = pLayerList.Item(strName)
155:             Set pRow = pLayersTbl.GetRow(lngId)
156:             If (pLayer Is Nothing) Then 'If the layer is missing from the map
157:                 lngFieldIdx = pLayersTbl.FindField("LAYEROBJECT")
158:                 Set pMemBlobStrm = pRow.Value(lngFieldIdx)
159:                 strTempDir = gODNRProjectDb.TempDirectory
160:                 strFilePath = strTempDir & "\" & strName & ".lyr"
161:                 pMemBlobStrm.SaveToFile strFilePath
                'Added section, 20060106, Jim McDonald
163:                 lngFieldIdx = pLayersTbl.FindField("LAYERTYPE")
164:                 strLayerType = pRow.Value(lngFieldIdx)
165:                 strLayerTypeName = LayerTypeName(strLayerType)
166:                 If (IsNull(pRow.Value(lngFieldIdx) = False)) Then
167:                     If (strLayerTypeName = "Boundary" Or strLayerTypeName = "Base 100") Then
168:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRBasemapWorkspace, g_strGeoDB_Path)
169:                     ElseIf (strLayerTypeName = "Bedrock Geology" Or strLayerTypeName = "Bedrock Topography") Then
170:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRGeologyWorkspace, g_strGeoDB_Path)
171:                     ElseIf (strLayerTypeName = "AUM") Then
172:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRAumWorkspace, g_strGeoDB_Path)
173:                     ElseIf (strLayerTypeName = "Oil-and-Gas Wells") Then
174:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNROGWellsWorkspace, g_strGeoDB_Path)
175:                     End If
176:                 End If
                'Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRGeologyWorkspace, g_strGeoDB_Path)
                'End Added Section, 20060106, Jim McDonald
179:             Else
                'CheckLayerDataSource pLayer 'This will fail in the future.  Need to fix so as to be able to detect different data sourcs-SDE, PGB, different databases, etc...
                'ChangeMapLayerWs pLayer, ODNRGeologyWorkspace
182:                 lngFieldIdx = pLayersTbl.FindField("LAYERTYPE")
183:                 If (IsNull(pRow.Value(lngFieldIdx)) = False) Then strLayerType = pRow.Value(lngFieldIdx)
184:                 strLayerTypeName = LayerTypeName(strLayerType)
185:                 If (strLayerTypeName = "Bedrock Geology" Or strLayerTypeName = "Bedrock Topography") Then
186:                     ChangeMapLayerWs pLayer, ODNRGeologyWorkspace
187:                 ElseIf (strLayerTypeName = "Boundary" Or strLayerTypeName = "Base 100") Then
188:                     ChangeMapLayerWs pLayer, ODNRBasemapWorkspace
189:                 ElseIf (strLayerTypeName = "AUM") Then
190:                     ChangeMapLayerWs pLayer, ODNRAumWorkspace
191:                 ElseIf (strLayerTypeName = "Oil-and-Gas Wells") Then
192:                     ChangeMapLayerWs pLayer, ODNROGWellsWorkspace
193:                 End If
                'MsgBox "strName = " & strName & vbNewLine & "strLayerType = " & strLayerType & vbNewLine & "strLayerTypeName = " & strLayerTypeName
195:             End If
196:             lngFieldIdx = pLayersTbl.FindField("QUAD24KFIELD")
197:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then str24KField = pRow.Value(lngFieldIdx)
198:             lngFieldIdx = pLayersTbl.FindField("QUAD100KFIELD")
199:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then str100KField = pRow.Value(lngFieldIdx)
200:             lngFieldIdx = pLayersTbl.FindField("EXPORTMETHOD")
201:             lngExportMethod = 0
202:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then lngExportMethod = pRow.Value(lngFieldIdx)
'            MsgBox "Layer.Name = " & pLayer.Name & _
'                vbCrLf & "Quad24KField = " & str24KField & _
'                vbCrLf & "Quad100KField = " & str100KField & _
'                vbCrLf & "ExportMethod = " & lngExportMethod
207:             pODNRLayer.LoadFromLayer m_pApp, Me, pLayer, odnrZoomLevelAll, str24KField, str100KField, m_lngActiveMap, lngExportMethod
208:             Set GetLayerByName = pODNRLayer
209:         Else
210:             Set GetLayerByName = Nothing
211:         End If
212:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".GetLayerByName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function NextLayer() As ODNRStateLayer
    On Error GoTo ErrorHandler
    
222:     If (m_blnIsloaded) Then
        Dim pLayerList As Dictionary, pLayersTbl As ITable
        Dim pODNRLayer As ODNRStateLayer, pLayer As ILayer, strName As String
        Dim str24KField As String, str100KField As String, pFieldNames As String
        Dim strLayerType As String 'Added 20051221, Jim McDonald
        Dim blnInMap As Boolean
        Dim lngExportMethod As Long
        Dim lngId As Long, pRow As IRow, lngFieldIdx As Long
        Dim pMemBlobStrm As IMemoryBlobStream, strFilePath As String, strTempDir As String
        Dim strLayerTypeName As String
        
233:         Set pLayerList = GetMapLayers(m_lngActiveMap)
234:         If (m_lngLayerIdx < pLayerList.Count) Then
235:             strName = pLayerList.Keys(m_lngLayerIdx)
236:             Set pODNRLayer = New ODNRStateLayer
237:             Set pLayer = FindLayer(strName)
238:             Set pLayersTbl = gODNRProjectDb.LayerObjectsTable
239:             lngId = pLayerList.Item(strName)
240:             Set pRow = pLayersTbl.GetRow(lngId)
241:             blnInMap = False
242:             If (pLayer Is Nothing) Then
                'If the layer is missing from the project MXD, then re-attach the layer using the
                'layer list in the projects database
245:                 lngFieldIdx = pLayersTbl.FindField("LAYEROBJECT")
246:                 Set pMemBlobStrm = pRow.Value(lngFieldIdx)
247:                 strTempDir = gODNRProjectDb.TempDirectory
248:                 strFilePath = strTempDir & "\" & strName & ".lyr"
249:                 pMemBlobStrm.SaveToFile strFilePath
                'Added section, 20060106, Jim McDonald
251:                 lngFieldIdx = pLayersTbl.FindField("LAYERTYPE")
252:                 strLayerType = pRow.Value(lngFieldIdx)
253:                 strLayerTypeName = LayerTypeName(strLayerType)
254:                 If (IsNull(pRow.Value(lngFieldIdx) = False)) Then
255:                     If (strLayerTypeName = "Boundary" Or strLayerTypeName = "Base 100") Then
256:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRBasemapWorkspace, g_strGeoDB_Path)
257:                     ElseIf (strLayerTypeName = "Bedrock Geology" Or strLayerTypeName = "Bedrock Topography") Then
258:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRGeologyWorkspace, g_strGeoDB_Path)
259:                     ElseIf (strLayerTypeName = "AUM") Then
260:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRAumWorkspace, g_strGeoDB_Path)
261:                     ElseIf (strLayerTypeName = "Oil-and-Gas Wells") Then
262:                         Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNROGWellsWorkspace, g_strGeoDB_Path)
263:                     End If
264:                 End If
                'Set pLayer = ODNR_Common.LoadFromGxFile(strFilePath, ODNRGeologyWorkspace, g_strGeoDB_Path)
                'This will fail in the future.  Need to fix so as to be able to detect different data sourcs-SDE, PGB, different databases, etc...
                'End Added section, 20060106, Jim McDonald
268:             Else
269:                 blnInMap = True
                'CheckLayerDataSource pLayer 'This will fail in the future.  Need to fix so as to be able to detect different data sourcs-SDE, PGB, different databases, etc...
                'ChangeMapLayerWs pLayer, ODNRGeologyWorkspace
272:                 lngFieldIdx = pLayersTbl.FindField("LAYERTYPE")
273:                 If (IsNull(pRow.Value(lngFieldIdx)) = False) Then strLayerType = pRow.Value(lngFieldIdx)
274:                 strLayerTypeName = LayerTypeName(strLayerType)
275:                 If (strLayerTypeName = "Bedrock Geology" Or strLayerTypeName = "Bedrock Topography") Then
276:                     ChangeMapLayerWs pLayer, ODNRGeologyWorkspace
277:                 ElseIf (strLayerTypeName = "Boundary" Or strLayerTypeName = "Base 100") Then
278:                     ChangeMapLayerWs pLayer, ODNRBasemapWorkspace
279:                 ElseIf (strLayerTypeName = "AUM") Then
280:                     ChangeMapLayerWs pLayer, ODNRAumWorkspace
281:                 ElseIf (strLayerTypeName = "Oil-and-Gas Wells") Then
282:                     ChangeMapLayerWs pLayer, ODNROGWellsWorkspace
283:                 End If
                'MsgBox "strName = " & strName & vbNewLine & "strLayerType = " & strLayerType & vbNewLine & "strLayerTypeName = " & strLayerTypeName
285:             End If
286:             lngFieldIdx = pLayersTbl.FindField("QUAD24KFIELD")
287:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then str24KField = pRow.Value(lngFieldIdx)
288:             lngFieldIdx = pLayersTbl.FindField("QUAD100KFIELD")
289:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then str100KField = pRow.Value(lngFieldIdx)
290:             lngFieldIdx = pLayersTbl.FindField("EXPORTMETHOD")
291:             lngExportMethod = 0
292:             If (IsNull(pRow.Value(lngFieldIdx)) = False) Then lngExportMethod = pRow.Value(lngFieldIdx)
'            MsgBox "Layer.Name = " & pLayer.Name & _
'                vbCrLf & "Quad24KField = " & str24KField & _
'                vbCrLf & "Quad100KField = " & str100KField & _
'                vbCrLf & "ExportMethod = " & lngExportMethod
297:             pODNRLayer.LoadFromLayer m_pApp, Me, pLayer, odnrZoomLevelAll, str24KField, str100KField, m_lngActiveMap, lngExportMethod
298:             Set NextLayer = pODNRLayer
299:             m_lngLayerIdx = m_lngLayerIdx + 1
300:         Else
301:             Set NextLayer = Nothing
302:         End If
303:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".NextLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub Reset()
    On Error GoTo ErrorHandler
    
313:     If (m_blnIsloaded) Then m_lngLayerIdx = 0
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Reset " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub Load(pApp As esriFramework.IApplication, pParent As ODNRProject)
    On Error GoTo ErrorHandler
    
323:     If Not (pParent Is Nothing) Then
324:         Set m_pApp = pApp
325:         Set m_pParent = pParent
326:         m_lngActiveMap = odnrGeologyMap
327:         m_blnIsloaded = True
328:     End If

    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Load " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function FindLayer(strName As String) As ILayer
    On Error GoTo ErrorHandler
    
338:     If (m_blnIsloaded) Then
        Dim pMxDoc As IMxDocument, pMap As IMap, pLayer As ILayer, lngIdx As Long
        Dim strMapName As String
341:         Set pMxDoc = m_pApp.Document
342:         Set pMap = m_pParent.ProjectMap(m_lngActiveMap)
343:         If Not (pMap Is Nothing) Then
344:             For lngIdx = 0 To pMap.LayerCount - 1
345:                 If (strName = pMap.Layer(lngIdx).Name) Then
346:                     Set pLayer = pMap.Layer(lngIdx)
347:                     Exit For
348:                 End If
349:             Next
350:         End If
351:         If Not (pLayer Is Nothing) Then
352:             Set FindLayer = pLayer
353:         Else
354:             Set FindLayer = Nothing
355:         End If
356:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".FindLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function GetMapLayers(lngMapType As ODNRMapType) As Dictionary
    On Error GoTo ErrorHandler
    
366:     If (lngMapType <> odnrMapTypeUnknown) Then
        Dim pMapsTbl As ITable, pLayersTbl As ITable
        Dim pQf As IQueryFilter, pCursor As ICursor, pRow As IRow
        Dim lngId As Long, pLayerRow As IRow
        Dim strProjectCode As String, pLayerList As Dictionary
371:         Set pMapsTbl = gODNRProjectDb.MapLayersTable
372:         Set pLayersTbl = gODNRProjectDb.LayerObjectsTable
373:         strProjectCode = gODNRProjectDb.ActiveProjectCode
374:         Set pQf = New QueryFilter
375:         pQf.WhereClause = "[PROJECTCODE] = '" & strProjectCode & "' And [MAPTYPE] = " & lngMapType
376:         Set pCursor = pMapsTbl.Search(pQf, False)
377:         Set pRow = pCursor.NextRow
378:         Set pLayerList = New Dictionary
379:         Do While Not pRow Is Nothing
380:             lngId = pRow.Value(3)
381:             Set pLayerRow = pLayersTbl.GetRow(lngId)
382:             pLayerList.Add Key:=pLayerRow.Value(1), Item:=lngId
383:             Set pRow = pCursor.NextRow
384:         Loop
385:         Set GetMapLayers = pLayerList
386:     End If

    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".GetMapLayers " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ChangeMapLayerWs(pLayer As ILayer, pWsProj As IWorkspace)
'------------------------------------------------------------------------------
'Title: ChangeMapLayerWs
'Date: 20051220
'Version: 1.0
'Abstract: This proceedure will examine the workspace properties and
'paths on map layers, and if necessary, change the workspace paths/properties
'to the project workspace path/properties.
'------------------------------------------------------------------------------
'James McDonald
'GIMS Specialist
'Ohio Division of Geological Survey
'2045 Morse Road
'Columbus, OH  43229-6693
'Ph. (614) 265-6601
'E-mail: jim.mcdonald@dnr.state.oh.us
'------------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Dim pFL As IFeatureLayer
    Dim pFlNew As IFeatureLayer
    Dim pDsFl As IDataset
    Dim pDsFlNew As IDataset
    Dim pFc As IFeatureClass
    Dim pFCNew As IFeatureClass
    Dim pWs As IWorkspace
    Dim pWsNew As IWorkspace
    Dim pFWs As IFeatureWorkspace
    Dim pWsFact As IWorkspaceFactory
    Dim pSqlSyn1 As ISQLSyntax
    Dim pSqlSyn2 As ISQLSyntax
    Dim pConnProp As IPropertySet
    Dim pConnPropProj As IPropertySet
    Dim strFullName As String, strDB As String, strOwner As String, strTable As String
    Dim strFcName As String
    Dim i As Long
    Dim pDsFc As IDataset
    Dim intLen As Integer
    Dim intLoc As Integer
    Dim strPos As String
    
    Dim varNames As Variant, varValues As Variant
    
    'Added 20060217, Delete Later after testing. Jim McDonald
    Dim lngI As Long
    
439:     Set pFL = pLayer
440:     Set pDsFl = pFL
    
442:     Set pFc = pFL.FeatureClass
443:     Set pDsFc = pFc
444:     Set pWs = pDsFc.Workspace
445:     Set pConnProp = pWs.ConnectionProperties
446:     Set pConnPropProj = pWsProj.ConnectionProperties
'    Set pWs = pDsFl.Workspace
'    Set pSqlSyn1 = pWs
'    Set pSqlSyn2 = pWsProj
450:     If (pWs.Type = esriLocalDatabaseWorkspace) Then
451:         Set pWsFact = New AccessWorkspaceFactory
452:         Set pWs = pWsFact.Open(pConnProp, 0)
453:     ElseIf (pWs.Type = esriRemoteDatabaseWorkspace) Then
454:         Set pWsFact = New SdeWorkspaceFactory
455:         Set pWs = pWsFact.Open(pConnProp, 0)
456:     End If
    
458:     strFcName = pDsFc.Name
'    MsgBox strFcName

461:     If (TypeOf pLayer Is IRasterLayer) Then 'Check to see if the layer is a raster layer. Maybe implement code layer to change raster layer workspaces.
        Exit Sub
463:     ElseIf ((TypeOf pLayer Is IFeatureLayer)) Then 'Check to see if the layer is a feature layer
464:         If (Not pWsProj Is Nothing) Then 'Check to see if the Project Workspace exists
465:             If (pWsProj.Type = esriFileSystemWorkspace) Then 'Workspace type is file based, such as shape files or coverages. We currently do not deal with file based types.
                Exit Sub
467:             ElseIf (pWsProj.Type = esriLocalDatabaseWorkspace) Then 'Workspace type is Personnel Geodatabase
468:                 If (Not pConnPropProj.IsEqual(pConnProp)) Then
469:                     If (pWs.Type = esriLocalDatabaseWorkspace) Then
'                        Set pWsNew = pWsProj
471:                         Set pFWs = pWsProj
472:                         Set pFCNew = pFWs.OpenFeatureClass(strFcName)
473:                         For lngI = 0 To 1000000
474:                         Next
475:                         Set pFL.FeatureClass = pFCNew
476:                     ElseIf (pWs.Type = esriRemoteDatabaseWorkspace) Then
477:                         intLen = Len(strFcName)
478:                         For i = 1 To intLen
479:                             strPos = Mid(strFcName, i, 1)
480:                             If (strPos = ".") Then intLoc = i
481:                         Next i
482:                         strFcName = Right(strFcName, intLen - intLoc)
'                        Set pWsNew = pWsProj
484:                         Set pFWs = pWsProj
485:                         For lngI = 0 To 1000000
486:                         Next
487:                         Set pFCNew = pFWs.OpenFeatureClass(strFcName)
488:                         Set pDsFc = pFCNew
489:                         For lngI = 0 To 1000000
490:                         Next
491:                         Set pFL.FeatureClass = pFCNew
492:                     End If
493:                 Else
                    Exit Sub
495:                 End If
496:             ElseIf (pWsProj.Type = esriRemoteDatabaseWorkspace) Then
497:                 If (Not pConnPropProj.IsEqual(pConnProp)) Then
498:                     If (pWs.Type = esriLocalDatabaseWorkspace) Then
'                        Set pWsNew = pWsProj
500:                         Set pFWs = pWsProj
501:                         Set pFCNew = pFWs.OpenFeatureClass(strFcName)
502:                         For lngI = 0 To 1000000
503:                         Next
504:                         Set pFL.FeatureClass = pFCNew
505:                     ElseIf (pWs.Type = esriRemoteDatabaseWorkspace) Then
506:                         intLen = Len(strFcName)
507:                         For i = 1 To intLen
508:                             strPos = Mid(strFcName, i, 1)
509:                             If (strPos = ".") Then intLoc = i
510:                         Next i
511:                         strFcName = Right(strFcName, intLen - intLoc)
'                        Set pWsNew = pWsProj
513:                         Set pFWs = pWsProj
514:                         Set pFCNew = pFWs.OpenFeatureClass(strFcName)
515:                         For lngI = 0 To 1000000
516:                         Next
517:                         Set pFL.FeatureClass = pFCNew
518:                     End If
519:                 Else
                    Exit Sub
521:                 End If
522:             End If
523:         Else 'If the Project Workspace does not exist, then exit the subroutine
            Exit Sub
525:         End If
526:     Else
        Exit Sub
528:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ChangeMapLayerWs " & pLayer.Name & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get LayerTypeName(strLyrValue As String) As String
    On Error GoTo ErrorHandler
    
538:     If (m_blnIsloaded) Then
        Dim pLayerTypeDomain As ICodedValueDomain, lngIdx As Long
540:         Set pLayerTypeDomain = gODNRProjectDb.LayerTypeDomain
541:         For lngIdx = 0 To pLayerTypeDomain.CodeCount - 1
542:             If (pLayerTypeDomain.Value(lngIdx) = strLyrValue) Then
543:                 LayerTypeName = pLayerTypeDomain.Name(lngIdx)
544:                 Exit For
545:             End If
546:         Next
547:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerTypeName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
