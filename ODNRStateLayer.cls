VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRStateLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************
'
'   Program:    ODNRStateLayer
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that provides common functionality
'               for the project layers covering the entire state.
'               This object is owned by the ODNRStateLayers object.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pParent As ODNRStateLayers
Private m_pLayer As ILayer
Private m_str24KFieldName As String
Private m_str100KFieldName As String
Private m_pExportMethod As ODNRQuadExportMethod
Private m_pVisibilityLevel As ODNRVisibilityLevel
Private m_lngMapType As ODNRMapType
Private m_blnIsloaded As Boolean

Const c_strModuleName As String = "ODNRStateLayer"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
34:     m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
44:     Set m_pLayer = Nothing
45:     m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get Name() As String
    On Error GoTo ErrorHandler
    
55:     If (m_blnIsloaded) Then Name = m_pLayer.Name
        
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Name " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayer() As ILayer
    On Error GoTo ErrorHandler
    
65:     If (m_blnIsloaded) Then
66:         Set ESRILayer = m_pLayer
67:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ESRILayerType() As esriDatasetType
    On Error GoTo ErrorHandler
    
77:     If (m_blnIsloaded) Then
        Dim pDs As IDataset, pFlyr As IFeatureLayer, pRLyr As IRasterLayer
79:         If (TypeOf m_pLayer Is IFeatureLayer) Then
80:             Set pFlyr = m_pLayer
81:             Set pDs = pFlyr
82:         ElseIf (TypeOf m_pLayer Is IRasterLayer) Then
83:             Set pRLyr = m_pLayer
84:             Set pDs = pRLyr
85:         End If
86:         If Not (pDs Is Nothing) Then
87:             ESRILayerType = pDs.Type
88:         Else
89:             ESRILayerType = esriDTAny
90:         End If
91:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ESRILayerType " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get VisibilityLevel() As ODNRVisibilityLevel
    On Error GoTo ErrorHandler
    
101:     If (m_blnIsloaded) Then VisibilityLevel = m_pVisibilityLevel
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".VisibilityLevel " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get InMap() As Boolean
    On Error GoTo ErrorHandler
    
111:     InMap = False
112:     If (m_blnIsloaded) Then
        Dim pMap As IMap, pLayer As ILayer, lngIdx As Long
114:         Set pMap = m_pParent.Parent.ProjectMap(m_lngMapType)
115:         If Not (pMap Is Nothing) Then
116:             For lngIdx = 0 To pMap.LayerCount - 1
117:                 Set pLayer = pMap.Layer(lngIdx)
118:                 If (pLayer Is m_pLayer) Then
119:                     InMap = True
120:                     Exit For
121:                 End If
122:             Next
123:         End If
124:     End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".InMap " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get CanQueryBy24KQuad() As Boolean
    On Error GoTo ErrorHandler
    
134:     CanQueryBy24KQuad = False
135:     If (m_blnIsloaded) Then
136:         If (TypeOf m_pLayer Is IFeatureLayer) Then
137:             If (m_str24KFieldName <> "") And (m_pExportMethod = odnrExportMethodField) Then
                Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
139:                 Set pFlyr = m_pLayer
140:                 Set pFc = pFlyr.FeatureClass
141:                 If (pFc.FindField(m_str24KFieldName) > -1) Then
142:                     CanQueryBy24KQuad = True
143:                 End If
144:             ElseIf (m_pExportMethod = odnrExportMethodSpatial) Then
145:                 CanQueryBy24KQuad = True
146:             ElseIf (m_pExportMethod = odnrExportMethodClip) Then
147:                 CanQueryBy24KQuad = True
148:             End If
149:         End If
150:     End If
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".CanQueryBy24KQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get CanQueryBy100KQuad() As Boolean
    On Error GoTo ErrorHandler
    
159:     CanQueryBy100KQuad = False
160:     If (m_blnIsloaded) Then
161:         If (TypeOf m_pLayer Is IFeatureLayer) Then
162:             If (m_str100KFieldName <> "") And (m_pExportMethod = odnrExportMethodField) Then
                Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
164:                 Set pFlyr = m_pLayer
165:                 Set pFc = pFlyr.FeatureClass
166:                 If (pFc.FindField(m_str100KFieldName) > -1) Then
167:                     CanQueryBy100KQuad = True
168:                 End If
169:             ElseIf (m_pExportMethod = odnrExportMethodSpatial) Then
170:                 CanQueryBy100KQuad = True
171:             ElseIf (m_pExportMethod = odnrExportMethodClip) Then
172:                 CanQueryBy100KQuad = True
173:             End If
174:         End If
175:     End If
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".CanQueryBy100KQuad " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get Quad24KFieldName() As String
    On Error GoTo ErrorHandler

184:     If (m_blnIsloaded) Then Quad24KFieldName = m_str24KFieldName

    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Quad24KFieldName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get Quad100KFieldName() As String
    On Error GoTo ErrorHandler
    
194:     If (m_blnIsloaded) Then Quad100KFieldName = m_str100KFieldName
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".Quad100KFieldName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Sub LimitToQuads()
    On Error GoTo ErrorHandler
    
204:     If (m_blnIsloaded) And (TypeOf m_pLayer Is IFeatureLayer) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
        Dim pFlyrDef As IFeatureLayerDefinition
        Dim strQryDef As String, lngIdx As Long
        Dim pQuad As ODNRQuad, strFieldName As String
209:         Set pFlyr = m_pLayer
210:         Set pFc = pFlyr.FeatureClass
211:         If (m_pExportMethod = odnrExportMethodField) Then
212:             If (m_pParent.Parent.QuadScale = odnr24K) Then
213:                 strFieldName = m_str24KFieldName
214:             ElseIf (m_pParent.Parent.QuadScale = odnr100K) Then
215:                 strFieldName = m_str100KFieldName
216:             End If
217:             If (pFc.FindField(strFieldName) > -1) And (m_pParent.Parent.Quads.QuadCount > 0) Then
218:                 Set pFlyrDef = pFlyr
219:                 m_pParent.Parent.Quads.Reset
220:                 Set pQuad = m_pParent.Parent.Quads.NextQuad
221:                 Do While Not pQuad Is Nothing
222:                     If (strQryDef = "") Then
223:                         strQryDef = strFieldName & " in ('" & pQuad.QuadName & "'"
224:                     Else
225:                         strQryDef = strQryDef & ",'" & pQuad.QuadName & "'"
226:                     End If
227:                     Set pQuad = m_pParent.Parent.Quads.NextQuad
228:                 Loop
229:                 strQryDef = strQryDef & ")"
230:                 pFlyrDef.DefinitionExpression = strQryDef
231:             End If
232:         End If
233:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LimitToQuads " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function QuadsSelectionSet() As ISelectionSet
    On Error GoTo ErrorHandler
    
243:     Set QuadsSelectionSet = Nothing
244:     If (m_blnIsloaded) And (TypeOf m_pLayer Is IFeatureLayer) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass
        Dim strQryDef As String, lngIdx As Long
        Dim pQuad As ODNRQuad, strFieldName As String
        Dim pQf As IQueryFilter, pSpF As ISpatialFilter
        Dim pGeoCol As IGeometryCollection, pPolygon As IPolygon
        Dim pCtrPt As IPoint, pTrans2D As ITransform2D, pEnv As IEnvelope, pArea As IArea
251:         Set pFlyr = m_pLayer
252:         Set pFc = pFlyr.FeatureClass
253:         If (m_pExportMethod = odnrExportMethodField) Then
254:             If (m_pParent.Parent.QuadScale = odnr24K) Then
255:                 strFieldName = m_str24KFieldName
256:             ElseIf (m_pParent.Parent.QuadScale = odnr100K) Then
257:                 strFieldName = m_str100KFieldName
258:             End If
259:             If (pFc.FindField(strFieldName) > -1) And (m_pParent.Parent.Quads.QuadCount > 0) Then
260:                 m_pParent.Parent.Quads.Reset
261:                 Set pQuad = m_pParent.Parent.Quads.NextQuad
262:                 Do While Not pQuad Is Nothing
263:                     If (strQryDef = "") Then
264:                         strQryDef = strFieldName & " in ('" & pQuad.QuadName & "'"
265:                     Else
266:                         strQryDef = strQryDef & ",'" & pQuad.QuadName & "'"
267:                     End If
268:                     Set pQuad = m_pParent.Parent.Quads.NextQuad
269:                 Loop
270:                 strQryDef = strQryDef & ")"
271:                 Set pQf = New QueryFilter
272:                 pQf.WhereClause = strQryDef
273:                 Set QuadsSelectionSet = pFc.Select(pQf, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
274:             End If
275:         ElseIf (m_pExportMethod = odnrExportMethodSpatial) Then
276:             If (m_pParent.Parent.Quads.QuadCount > 0) Then
277:                 m_pParent.Parent.Quads.Reset
278:                 Set pQuad = m_pParent.Parent.Quads.NextQuad
279:                 Set pGeoCol = New Polygon
280:                 Do While Not pQuad Is Nothing
281:                     If (pFc.ShapeType = esriGeometryPoint) Then
282:                         pGeoCol.AddGeometryCollection pQuad.QuadBoundary
283:                     Else
284:                         Set pEnv = pQuad.QuadBoundary.Envelope
285:                         Set pArea = pEnv
286:                         Set pCtrPt = pArea.Centroid
287:                         Set pTrans2D = pQuad.QuadBoundary
288:                         With pTrans2D
289:                             .Scale pCtrPt, 0.99, 0.99
290:                         End With
291:                         Set pPolygon = pTrans2D
292:                         pGeoCol.AddGeometryCollection pPolygon
293:                     End If
294:                     Set pQuad = m_pParent.Parent.Quads.NextQuad
295:                 Loop
296:                 Set pPolygon = pGeoCol
297:                 Set pSpF = New SpatialFilter
298:                 pSpF.GeometryField = pFc.shapeFieldName
299:                 pSpF.SpatialRel = esriSpatialRelIntersects
300:                 Set pSpF.Geometry = pPolygon
301:                 Set QuadsSelectionSet = pFc.Select(pSpF, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
302:             End If
303:         End If
304:     End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadsSelectionSet " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ShowAllFeatures()
    On Error GoTo ErrorHandler
    
314:     If (m_blnIsloaded) Then
315:         If (TypeOf m_pLayer Is IFeatureLayer) Then
            Dim pFlyr As IFeatureLayer, pFlyrDef As IFeatureLayerDefinition
317:             Set pFlyr = m_pLayer
318:             Set pFlyrDef = pFlyr
319:             pFlyrDef.DefinitionExpression = ""
320:         End If
321:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".ShowAllFeatures " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Public Function Export(strExportDir As String, strName As String) As String
'    On Error GoTo ErrorHandler
'
'    If (m_blnIsloaded) Then
'        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pOutFc As IFeatureClass
'        Dim pDs As IDataset, pDsName As IDatasetName, pOutDs As IDataset
'        Dim pInFcName As IFeatureClassName, pOutFcName As IFeatureClassName
'        Dim pEnumInvalidObj As IEnumInvalidObject, pInvalidObjInfo As IInvalidObjectInfo
'        Dim pInWs As IWorkspace, pOutWs As IWorkspace, pWsFact As IWorkspaceFactory
'        Dim pSelSet As ISelectionSet, pOutWsName As IWorkspaceName
'        Dim pInFields As IFields, pOutFields As IFields
'        Dim pFieldChecker As IFieldChecker, pFDataCon As IFeatureDataConverter2
'
'        Dim pRLyr As IRasterLayer, strFilePath As String, strDest As String
'        Dim pFSO As FileSystemObject, strBaseFileName As String
'        Dim strMsg As String
'
'        strMsg = "Initializing export of " & Me.Name & "..."
'
'        If (m_pLayer.Valid) Then
'            If (TypeOf m_pLayer Is IFeatureLayer) Then
'                strMsg = strMsg & vbCrLf & vbTab & "exported featurelayer: " & m_pLayer.Name
'                Set pFlyr = m_pLayer
'                Set pFc = pFlyr.FeatureClass
'                Set pSelSet = QuadsSelectionSet
'
'                Set pDs = pFc
'                Set pInFcName = pDs.FullName
'                Set pInWs = pDs.Workspace
'
'                Set pWsFact = New ShapefileWorkspaceFactory
'                Set pOutWs = pWsFact.OpenFromFile(strExportDir, 0)
'                Set pDs = pOutWs
'                Set pOutWsName = pDs.FullName
'
'                Set pOutFcName = New FeatureClassName
'                Set pDsName = pOutFcName
'                pDsName.Name = strName
'                Set pDsName.WorkspaceName = pOutWsName
'
'                Set pInFields = pFc.Fields
'                Set pFieldChecker = New FieldChecker
'                pFieldChecker.InputWorkspace = pInWs
'                Set pFieldChecker.ValidateWorkspace = pOutWs
'                pFieldChecker.Validate pInFields, Nothing, pOutFields
'
'                Set pFDataCon = New FeatureDataConverter
'                Set pEnumInvalidObj = pFDataCon.ConvertFeatureClass(pInFcName, Nothing, pSelSet, Nothing, pOutFcName, Nothing, pOutFields, "", 100, 0)
'                Set pInvalidObjInfo = pEnumInvalidObj.Next
'                If (pInvalidObjInfo Is Nothing) Then
'                    strMsg = strMsg & " successfully."
'                Else
'                    strMsg = strMsg & " unsuccessfully."
'                End If
'
'                Set pInWs = Nothing
'                Set pOutWs = Nothing
'                Set pFieldChecker = Nothing
'                Set pFDataCon = Nothing
'                Set pEnumInvalidObj = Nothing
'
'            ElseIf (TypeOf m_pLayer Is IRasterLayer) Then
'                strMsg = strMsg & vbCrLf & vbTab & "exported rasterlayer: " & m_pLayer.Name
'                Set pRLyr = m_pLayer
'                Set pFSO = New FileSystemObject
'                strFilePath = pRLyr.FilePath
'                strBaseFileName = pFSO.GetBaseName(strFilePath)
'                strDest = strExportDir & "\"
'                strFilePath = Left(strFilePath, Len(strFilePath) - 3) & "*"
'                pFSO.CopyFile strFilePath, strDest, True
'                If (pFSO.FileExists(strDest & strBaseFileName)) Then
'                    strMsg = strMsg & " successfully."
'                Else
'                    strMsg = strMsg & " unsuccessfully."
'                End If
'                Set pFSO = Nothing
'            End If
'        Else
'            strMsg = strMsg & vbCrLf & vbTab & "layer: " & m_pLayer.Name & " is invalid."
'        End If
'        Export = strMsg
'    End If
'
'    Exit Function
'
'ErrorHandler:
'    HandleError True, c_strModuleName & ".Export " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Function

Public Sub LoadFromLayer(pApp As esriFramework.IApplication, pParent As ODNRStateLayers, pLayer As ILayer, pVisibilityLevel As ODNRVisibilityLevel, str24KField As String, str100KField As String, lngMapType As ODNRMapType, pExportMethod As ODNRQuadExportMethod)
    On Error GoTo ErrorHandler
    
420:     If (m_blnIsloaded = False) Then
421:         If Not (pLayer Is Nothing) Then
422:             Set m_pApp = pApp
423:             Set m_pParent = pParent
424:             m_pVisibilityLevel = pVisibilityLevel
425:             m_pExportMethod = pExportMethod
426:             m_str24KFieldName = str24KField
427:             m_str100KFieldName = str100KField
428:             m_lngMapType = lngMapType
429:             Set m_pLayer = pLayer
430:             If Not (m_pLayer Is Nothing) Then
431:                 m_blnIsloaded = True
432:             End If
433:         End If
434:     End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadFromLayer " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function ExportByClip(strExportDir As String) As String
     On Error GoTo ErrorHandler
    
    Dim pIntersectFClass As IFeatureClass
    Dim pFilter As IQueryFilter
    Dim pQfClip As IQueryFilter
    Dim pIntersectFCursor As IFeatureCursor
    Dim pIntersectFeature As IFeature
    Dim pIntersectTopo As ITopologicalOperator
    Dim pTopoOp2 As ITopologicalOperator2
    Dim pInterTopoOp2 As ITopologicalOperator2
    Dim pOverlayTopoOp As ITopologicalOperator
    Dim pPolygon As IPolygon
    Dim pSpatialFilter As ISpatialFilter
    Dim pIntersectFields As IFields
    Dim pOverlayFCursor As IFeatureCursor
    Dim pOverlayFClass As IFeatureClass
    Dim pOverlayFields As IFields
    Dim pOverlayFeature As IFeature
    Dim pOutputfClass As IFeatureClass
    Dim newGeometry As IGeometry
    Dim newGeomFix As IGeometry
    Dim pInterGeomFix As IGeometry
    Dim pOverlayIndex As Long
    Dim pOldFieldsCount As Long
    Dim Item As Long
    Dim pOldField As IField
    Dim pNewFeatureCursor As IFeatureCursor, pNewBuffer As IFeatureBuffer
    Dim intPolyCount As Integer
    Dim intNewBufferIdx As Integer
    Dim pRing() As IRing
    Dim pIntRing() As IRing
    Dim pNewPolygon As IPolygon2
    Dim l As Long
    Dim pArea As IArea
    Dim pPolygon2 As IGeometryCollection
    Dim pPolyline As IGeometryCollection
    Dim pPolyline2 As IGeometryCollection
    Dim pPolyline3 As IPolyline
    Dim pOverlayFLayer As IFeatureLayer
    Dim pQuad As ODNRQuad
    Dim pStateLayer As ODNRStateLayer
    Dim strDefQry As String
    Dim pNewGFlyr As IGeoFeatureLayer
    Dim pGFlyr As IGeoFeatureLayer
    Dim strField As String
    Dim pAnnoProps As IAnnotateLayerPropertiesCollection
    Dim pAnnoLabelProps As IAnnotateLayerProperties
    Dim pLabelEngine As ILabelEngineLayerProperties 'Delete in the future
    Dim blnDispAnno As Boolean
    Dim pFRenderer As IFeatureRenderer
    Dim n As Integer
    Dim intOverFeatCount As Integer
    Dim pNewlayer As IFeatureLayer
    Dim pIntersectFLyr As IFeatureLayer
    Dim pStateLayers As ODNRStateLayers
    Dim strQryDef As String
    Dim pOverlayLayer As ILayer
    
500:     Set pOverlayFLayer = m_pLayer
501:     Set pOverlayLayer = pOverlayFLayer
502:     Set pIntersectFLyr = m_pParent.Parent.QuadFeatureLayer
503:     Set pIntersectFClass = pIntersectFLyr.FeatureClass 'Quadrangle feature class
504:     Set pOverlayFClass = pOverlayFLayer.FeatureClass 'Geology feature classes
505:     Set pOutputfClass = NewShapeFile(pOverlayFClass, strExportDir) 'New Geology Shape file feature classes
    
    'Maintain the original fields collection
508:     Set pOverlayFields = pOverlayFClass.Fields
509:     pOldFieldsCount = pOverlayFClass.Fields.FieldCount
    
511:     Set pStateLayer = gODNRProject.StateLayers.GetLayerByName(m_pParent.Parent.QuadFeatureLayer.Name)
512:     If Not (pStateLayer Is Nothing) Then
513:         If (m_pParent.Parent.QuadScale = odnr24K) Then
514:             strField = pStateLayer.Quad24KFieldName
515:         ElseIf (m_pParent.Parent.QuadScale = odnr100K) Then
516:             strField = pStateLayer.Quad100KFieldName
517:         End If
518:     Else
        Exit Function
520:     End If
    
522:     If (m_pParent.Parent.Quads.QuadCount > 0) Then
        
524:         strQryDef = ""
525:         m_pParent.Parent.Quads.Reset
526:         Set pQuad = m_pParent.Parent.Quads.NextQuad
527:         Do While Not pQuad Is Nothing
528:             If (strQryDef = "") Then
529:                 strQryDef = strField & " = '" & pQuad.QuadName & "'"
530:             Else
531:                 strQryDef = strQryDef & " OR " & strField & " = '" & pQuad.QuadName & "'"
532:             End If
533:             Set pQuad = m_pParent.Parent.Quads.NextQuad
534:         Loop
        
        ' get the buffer record for the Quad layer
537:         Set pQfClip = New QueryFilter
538:         pQfClip.WhereClause = strQryDef
539:         Set pIntersectFCursor = pIntersectFLyr.Search(pQfClip, False)
        
        'Create an insert cursor to hold the new features
542:         Set pNewFeatureCursor = pOutputfClass.Insert(True)
        
        'Start with buffer feature for the Quad layer
545:         Set pIntersectFeature = pIntersectFCursor.NextFeature
        
547:         While Not pIntersectFeature Is Nothing
548:             Set pInterTopoOp2 = pIntersectFeature.ShapeCopy
549:             pInterTopoOp2.IsKnownSimple = False
550:             pInterTopoOp2.Simplify
551:             Set pIntersectTopo = pInterTopoOp2
552:             Set pInterGeomFix = pIntersectTopo

            'get the intersecting layer features - Quad Layer
555:             Set pSpatialFilter = New SpatialFilter
556:             pSpatialFilter.GeometryField = pIntersectFClass.shapeFieldName
'            Set pSpatialFilter.Geometry = pIntersectFeature.Shape
558:             Set pSpatialFilter.Geometry = pInterGeomFix
559:             pSpatialFilter.SpatialRel = esriSpatialRelIntersects
560:             Set pOverlayFCursor = pOverlayFClass.Search(pSpatialFilter, False) 'Find all the Geology features that intersect the quad layer
            'Cycle through the features intersecting that buffer poly
562:             Set pOverlayFeature = pOverlayFCursor.NextFeature
563:             intPolyCount = 0
564:             While Not pOverlayFeature Is Nothing

566:                 If (pOutputfClass.ShapeType = esriGeometryPolygon) Then
567:                     Set pTopoOp2 = pOverlayFeature.ShapeCopy
568:                     pTopoOp2.IsKnownSimple = False
569:                     pTopoOp2.Simplify
570:                     Set newGeomFix = pTopoOp2
571:                     Set newGeometry = pIntersectTopo.Intersect(newGeomFix, esriGeometry2Dimension)
572:                     Set pOverlayTopoOp = newGeometry

574:                     If (newGeometry.IsEmpty = False) Then
575:                         Set pPolygon = newGeometry
                        ReDim pRing(pPolygon.ExteriorRingCount - 1)

578:                         Set pNewPolygon = pPolygon
579:                         pNewPolygon.QueryExteriorRingsEx pPolygon.ExteriorRingCount, pRing(0)
580:                         For l = 0 To pPolygon.ExteriorRingCount - 1
581:                             Set pPolygon2 = New Polygon
582:                             pPolygon2.AddGeometry pRing(l)

584:                             If (pPolygon.InteriorRingCount(pRing(l)) > 0) Then
                                ReDim pIntRing(pPolygon.InteriorRingCount(pRing(l)) - 1)
586:                                 pNewPolygon.QueryInteriorRingsEx pRing(l), pPolygon.InteriorRingCount(pRing(l)), pIntRing(0)
587:                                 For n = 0 To pPolygon.InteriorRingCount(pRing(l)) - 1
588:                                     pPolygon2.AddGeometry pIntRing(n)
589:                                 Next
590:                             End If

592:                             Set pArea = pPolygon2
                            'Calculate attributes in the new feature based on the overlay layer - The Quad layer
594:                             Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
595:                             Set pNewBuffer.Shape = pPolygon2

                            'Calculate the original attributes based on the overlay - The Geology Layer
598:                             For Item = 0 To pOldFieldsCount - 1
599:                                 Set pOldField = pOverlayFClass.Fields.Field(Item)
600:                                 intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
601:                                 If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                        And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                        And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
604:                                     pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
605:                                     If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                        Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
607:                                         If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
608:                                             pNewBuffer.Value(intNewBufferIdx) = 0
609:                                         Else
610:                                             pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
611:                                         End If
612:                                     Else
613:                                         If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
614:                                             pNewBuffer.Value(intNewBufferIdx) = ""
615:                                         Else
616:                                             pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
617:                                         End If
618:                                     End If
619:                                 ElseIf UCase(pOldField.Name) = "SHAPE_AREA" Or UCase(pOldField.Name) = "SHAPE.AREA" Then
620:                                     intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_area")
621:                                     pNewBuffer.Value(intNewBufferIdx) = pArea.Area
622:                                 ElseIf UCase(pOldField.Name) = "SHAPE_LENGTH" Then
623:                                     intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_leng")
624:                                     pNewBuffer.Value(intNewBufferIdx) = pPolygon.Length
625:                                 ElseIf UCase(pOldField.Name) = "SHAPE.LEN" Then
626:                                     intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_len")
627:                                     pNewBuffer.Value(intNewBufferIdx) = pPolygon.Length
628:                                 End If
629:                             Next Item

                            'Store the feature
632:                             pNewFeatureCursor.InsertFeature pNewBuffer
633:                         Next l
634:                     End If
635:                 ElseIf (pOutputfClass.ShapeType = esriGeometryPolyline) Then
636:                     Set newGeometry = pIntersectTopo.Intersect(pOverlayFeature.Shape, pOverlayFeature.Shape.Dimension)
637:                     Set pPolyline = newGeometry
638:                     For l = 0 To pPolyline.GeometryCount - 1
639:                         Set pPolyline2 = New Polyline
640:                         pPolyline2.AddGeometry pPolyline.Geometry(l)
641:                         Set pPolyline3 = pPolyline2

                        'Calculate attributes in the new feature based on the overlay layer - The Quad layer
644:                         Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
645:                         Set pNewBuffer.Shape = pPolyline2

                        'Calculate the original attributes based on the overlay - The Geology Layer
648:                         For Item = 0 To pOldFieldsCount - 1
649:                             Set pOldField = pOverlayFClass.Fields.Field(Item)
650:                             intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
651:                             If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                    And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                    And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
654:                                 pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
655:                                 If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                    Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
657:                                     If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
658:                                         pNewBuffer.Value(intNewBufferIdx) = 0
659:                                     Else
660:                                         pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
661:                                     End If
662:                                 Else
663:                                     If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
664:                                         pNewBuffer.Value(intNewBufferIdx) = ""
665:                                     Else
666:                                         pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
667:                                     End If
668:                                 End If
669:                             ElseIf UCase(pOldField.Name) = "SHAPE_LENGTH" Then
670:                                 intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_leng")
671:                                 pNewBuffer.Value(intNewBufferIdx) = pPolyline3.Length
672:                             ElseIf UCase(pOldField.Name) = "SHAPE.LEN" Then
673:                                 intNewBufferIdx = pNewBuffer.Fields.FindField("SHAPE_len")
674:                                 pNewBuffer.Value(intNewBufferIdx) = pPolyline3.Length
675:                             End If
676:                         Next Item

                        'Store the feature
679:                         pNewFeatureCursor.InsertFeature pNewBuffer
680:                         Next l
681:                 ElseIf (pOutputfClass.ShapeType = esriGeometryPoint) Then
682:                     Set newGeometry = pIntersectTopo.Intersect(pOverlayFeature.Shape, pOverlayFeature.Shape.Dimension)
683:                     Set pNewBuffer = pOutputfClass.CreateFeatureBuffer
684:                     Set pNewBuffer.Shape = newGeometry

                    'Calculate the original attributes based on the overlay - The Geology Layer
687:                     For Item = 0 To pOldFieldsCount - 1
688:                         Set pOldField = pOverlayFClass.Fields.Field(Item)
689:                         intNewBufferIdx = pNewBuffer.Fields.FindField(Left(pOldField.Name, 10))
690:                         If pOldField.Type <> esriFieldTypeGeometry And pOldField.Type <> esriFieldTypeOID _
                                    And UCase(pOldField.Name) <> "SHAPE_LENGTH" And UCase(pOldField.Name) <> "SHAPE_AREA" _
                                    And UCase(pOldField.Name) <> "SHAPE.LEN" And UCase(pOldField.Name) <> "SHAPE.AREA" Then
693:                             pOverlayIndex = pOverlayFields.FindField(pOldField.Name)
694:                             If pOldField.Type = esriFieldTypeDouble Or pOldField.Type = esriFieldTypeInteger _
                                Or pOldField.Type = esriFieldTypeSingle Or pOldField.Type = esriFieldTypeSmallInteger Then
696:                                 If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
697:                                     pNewBuffer.Value(intNewBufferIdx) = 0
698:                                 Else
699:                                     pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
700:                                 End If
701:                             Else
702:                                 If IsNull(pOverlayFeature.Value(pOverlayIndex)) Then
703:                                     pNewBuffer.Value(intNewBufferIdx) = ""
704:                                 Else
705:                                     pNewBuffer.Value(intNewBufferIdx) = pOverlayFeature.Value(pOverlayIndex)
706:                                 End If
707:                             End If
708:                         End If
709:                     Next Item

                    'Store the feature
712:                     pNewFeatureCursor.InsertFeature pNewBuffer
713:                 End If
714:                 Set pOverlayFeature = pOverlayFCursor.NextFeature
715:                 intPolyCount = intPolyCount + 1
716:             Wend
717:             Set pIntersectFeature = pIntersectFCursor.NextFeature
718:         Wend
719:     End If
    
721:     ExportByClip = "  : The number of features = " & intPolyCount
    
    Exit Function
  
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportByClip " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function NewShapeFile(pTemplateFC As IFeatureClass, strExportDir As String) As IFeatureClass
    On Error GoTo ErrorHandler
    
    Dim pFcExists As IFeatureClass
    Dim pFcDs As IDataset
    Dim pWsDs As IDataset
    Dim pTemplateDS As IDataset
    Dim pWsInDs As IDataset
    Dim pWsExists As IWorkspace
    Dim pWsIn As IWorkspace
    Dim pFWsIn As IFeatureWorkspace
    Dim pFWs As IFeatureWorkspace
    Dim pNewFeatClass As IFeatureClass
    Dim strFcName As String
    Dim i As Integer
    Dim intLen As Integer
    Dim intLoc As Integer
    Dim strPos As String
    'Objects used for the metadata
    Dim pNameInput As IName
    Dim pNameOutput As IName
    Dim pDsInput As IDataset
    Dim pDsOutput As IDataset
    Dim pMetadataInput As IMetadata
    Dim pMetadataOutput As IMetadata
    Dim pPsInput As IPropertySet
    Dim pPsOutput As IPropertySet
    
    'Open the input feature class workspace
758:     Set pTemplateDS = pTemplateFC
759:     Set pWsIn = pTemplateDS.Workspace
    
    'Get the template feature class name. Check to see if the feature class is in SDE, PGDB, or a file based name.
762:     If (pWsIn.Type = esriLocalDatabaseWorkspace) Then
763:         strFcName = pTemplateDS.Name
764:     ElseIf (pWsIn.Type = esriRemoteDatabaseWorkspace) Then
765:         strFcName = pTemplateDS.Name
766:         intLen = Len(strFcName)
767:         For i = 1 To intLen
768:             strPos = Mid(strFcName, i, 1)
769:             If (strPos = ".") Then intLoc = i
770:         Next i
771:         strFcName = Right(strFcName, intLen - intLoc)
772:     End If
    
    Dim strFcFullName As String
775:     strFcFullName = strExportDir & "\" & strFcName
    
    'Open the shape file workspace
    Dim pOutShpWspName As IWorkspaceName
779:     Set pOutShpWspName = New WorkspaceName
780:     pOutShpWspName.PathName = strExportDir
781:     pOutShpWspName.WorkspaceFactoryProgID = "esriCore.ShapefileWorkspaceFactory.1"
    Dim pName As IName
783:     Set pName = pOutShpWspName
    Dim pShapeWorkspace As IWorkspace
785:     Set pShapeWorkspace = pName.Open
    
    'Check to see if shape files exists, if so then delete
    Dim pWsFExists As IWorkspaceFactory
    Dim pTruncPath As String
790:     If InStr(1, strFcFullName, ".shp") > 0 Then
791:       pTruncPath = Left(strFcFullName, InStr(1, strFcFullName, ".shp") - 1)
792:     Else
793:       pTruncPath = strFcFullName
794:     End If
      
    Dim fs As Object
797:     Set fs = CreateObject("Scripting.FileSystemObject")
798:     If fs.FileExists(pTruncPath & ".shp") Or fs.FileExists(pTruncPath & ".dbf") Or _
        fs.FileExists(pTruncPath & ".shx") Then
800:         Set pWsFExists = New ShapefileWorkspaceFactory
801:         Set pWsExists = pWsFExists.OpenFromFile(strExportDir, 0)
        'Note: Remove EntryName function
803:         Set pFWs = pWsExists
804:         Set pFcExists = pFWs.OpenFeatureClass(strFcName)
805:         Set pFcDs = pFcExists
806:         pFcDs.Delete
807:     End If

    'Validate the fields
    Dim pOutputFields As IFields
    Dim pFieldChecker As IFieldChecker
812:     Set pFieldChecker = New FieldChecker
813:     Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
    
    'Clone the fields of the Template feature class.  Create new FID field for the new shape file.
    Dim pErrorEnum As IEnumFieldError
    Dim pNewFields As IFields, pField As IField
    Dim pNewFieldsEdit As IFieldsEdit
819:     Set pNewFields = New Fields
820:     Set pNewFieldsEdit = pNewFields
821:     pNewFieldsEdit.FieldCount = pTemplateFC.Fields.FieldCount
    
    Dim pClone As IClone
    Dim pCloneField As IField
    Dim pOIDFld As IField
    Dim pOIDFldEdit As IFieldEdit
    Dim j As Integer
828:     For j = 0 To pTemplateFC.Fields.FieldCount - 1
829:         If (pTemplateFC.Fields.Field(j).Type = esriFieldTypeOID) Then
830:             Set pOIDFld = New Field
831:             Set pOIDFldEdit = pOIDFld
832:             With pOIDFldEdit
833:                 .Name = "FID"
834:                 .Type = esriFieldTypeOID
835:             End With
836:             Set pNewFieldsEdit.Field(j) = pOIDFld
837:         Else
838:             Set pClone = pTemplateFC.Fields.Field(j)
839:             Set pCloneField = pClone.Clone
840:             Set pNewFieldsEdit.Field(j) = pCloneField
841:         End If
842:     Next
    
844:     pFieldChecker.Validate pNewFields, pErrorEnum, pOutputFields
    
    'Create the output featureclass
    Dim pUID As IUID
848:     Set pUID = pTemplateFC.clsid
    Dim shapeFieldName As String
850:     shapeFieldName = "Shape"
    
    Dim pFeatureWorkspace As IFeatureWorkspace
853:     Set pFeatureWorkspace = pShapeWorkspace
854:     Set pNewFeatClass = pFeatureWorkspace.CreateFeatureClass(strFcName, pOutputFields, pUID, Nothing, esriFTSimple, shapeFieldName, "")
    
    'Create Metadata for Output Feature Class
    'Setup metadata for input feature class
858:     Set pDsInput = pTemplateFC
859:     Set pNameInput = pDsInput.FullName
860:     Set pMetadataInput = pNameInput
861:     Set pPsInput = pMetadataInput.Metadata
    'Setup metadata for output feature class
863:     Set pDsOutput = pNewFeatClass
864:     Set pNameOutput = pDsOutput.FullName
865:     Set pMetadataOutput = pNameOutput
866:     pMetadataOutput.Synchronize esriMSAAlways, 0 'Create metadata XML file for output feature class
867:     pMetadataOutput.Metadata = pPsInput 'Copy input metadata records to output metadata records
868:     pMetadataOutput.Synchronize esriMSAAlways, 0 'Update the metadata XML file
   
870:     Set NewShapeFile = pNewFeatClass
    
    Exit Function
  
ErrorHandler:
    HandleError True, c_strModuleName & ".NewShapeFile " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function Export2(strExportDir As String, strName As String) As String
    On Error GoTo ErrorHandler
    
881:     If (m_blnIsloaded) Then
        Dim pFlyr As IFeatureLayer, pFc As IFeatureClass, pOutFc As IFeatureClass
        Dim pDs As IDataset, pDsName As IDatasetName, pOutDs As IDataset
        Dim pInFcName As IFeatureClassName, pOutFcName As IFeatureClassName
        Dim pEnumInvalidObj As IEnumInvalidObject, pInvalidObjInfo As IInvalidObjectInfo
        Dim pInWs As IWorkspace, pOutWs As IWorkspace, pWsFact As IWorkspaceFactory
        Dim pSelSet As ISelectionSet, pOutWsName As IWorkspaceName
        Dim pInFields As IFields, pOutFields As IFields
        Dim pFieldChecker As IFieldChecker, pFDataCon As IFeatureDataConverter2
        
        Dim pRLyr As IRasterLayer, strFilePath As String, strDest As String
        Dim pFSO As FileSystemObject, strBaseFileName As String
        Dim strMsg As String
        
895:         strMsg = "Initializing export of " & Me.Name & "..."
        
897:         If (m_pLayer.Valid) Then
898:             If (TypeOf m_pLayer Is IFeatureLayer) Then
899:                 If (m_pExportMethod = odnrExportMethodSpatial Or m_pExportMethod = odnrExportMethodField Or m_pExportMethod = odnrExportMethodAll) Then
900:                     strMsg = strMsg & vbCrLf & vbTab & "exported featurelayer: " & m_pLayer.Name
901:                     Set pFlyr = m_pLayer
902:                     Set pFc = pFlyr.FeatureClass
903:                     Set pSelSet = QuadsSelectionSet
    
905:                     Set pDs = pFc
906:                     Set pInFcName = pDs.FullName
907:                     Set pInWs = pDs.Workspace
                    
909:                     Set pWsFact = New ShapefileWorkspaceFactory
910:                     Set pOutWs = pWsFact.OpenFromFile(strExportDir, 0)
911:                     Set pDs = pOutWs
912:                     Set pOutWsName = pDs.FullName
                    
914:                     Set pOutFcName = New FeatureClassName
915:                     Set pDsName = pOutFcName
916:                     pDsName.Name = strName
917:                     Set pDsName.WorkspaceName = pOutWsName
                    
919:                     Set pInFields = pFc.Fields
920:                     Set pFieldChecker = New FieldChecker
921:                     pFieldChecker.InputWorkspace = pInWs
922:                     Set pFieldChecker.ValidateWorkspace = pOutWs
923:                     pFieldChecker.Validate pInFields, Nothing, pOutFields
                    
925:                     Set pFDataCon = New FeatureDataConverter
926:                     Set pEnumInvalidObj = pFDataCon.ConvertFeatureClass(pInFcName, Nothing, pSelSet, Nothing, pOutFcName, Nothing, pOutFields, "", 100, 0)
927:                     Set pInvalidObjInfo = pEnumInvalidObj.Next
928:                     If (pInvalidObjInfo Is Nothing) Then
929:                         strMsg = strMsg & " successfully."
930:                     Else
931:                         strMsg = strMsg & " unsuccessfully."
932:                     End If
                    
934:                     Set pInWs = Nothing
935:                     Set pOutWs = Nothing
936:                     Set pFieldChecker = Nothing
937:                     Set pFDataCon = Nothing
938:                     Set pEnumInvalidObj = Nothing
939:                 ElseIf (m_pExportMethod = odnrExportMethodClip) Then
940:                     strMsg = strMsg & ExportByClip(strExportDir)
941:                 ElseIf (m_pExportMethod = odnrExportMethodNone) Then
942:                     strMsg = strMsg & "not exported"
943:                 End If
944:             ElseIf (TypeOf m_pLayer Is IRasterLayer) Then
945:                 strMsg = strMsg & vbCrLf & vbTab & "exported rasterlayer: " & m_pLayer.Name
946:                 Set pRLyr = m_pLayer
947:                 Set pFSO = New FileSystemObject
948:                 strFilePath = pRLyr.FilePath
949:                 strBaseFileName = pFSO.GetBaseName(strFilePath)
950:                 strDest = strExportDir & "\"
951:                 strFilePath = Left(strFilePath, Len(strFilePath) - 3) & "*"
952:                 pFSO.CopyFile strFilePath, strDest, True
953:                 If (pFSO.FileExists(strDest & strBaseFileName)) Then
954:                     strMsg = strMsg & " successfully."
955:                 Else
956:                     strMsg = strMsg & " unsuccessfully."
957:                 End If
958:                 Set pFSO = Nothing
959:             End If
960:         Else
961:             strMsg = strMsg & vbCrLf & vbTab & "layer: " & m_pLayer.Name & " is invalid."
962:         End If
963:         Export2 = strMsg
964:     End If

    Exit Function
    
ErrorHandler:
    HandleError True, c_strModuleName & ".Export2 " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
