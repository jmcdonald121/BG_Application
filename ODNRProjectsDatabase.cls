VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODNRProjectsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************
'
'   Program:    ODNRProjectsDatabase
'   Author:     Greg Palovchik
'   Company:    Taratec Corporation
'               1251 Dublin Rd.
'               Columbus, OH 43215
'               (614) 291-2229
'   Date:       June 16, 2004
'   Purpose:    Creates an object that represents an ODNR Geology
'               Projects database.  This objects serves the projects.
'   Called from:
'
'*****************************************

Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_blnIsloaded As Boolean
Private m_pProjectsTable As ITable
Private m_pMapsTable As ITable
Private m_pLayersTable As ITable

Private m_pProjectTypeDomain As ICodedValueDomain
Private m_pQuadScaleDomain As ICodedValueDomain
Private m_pMapTypeDomain As ICodedValueDomain
Private m_pLayerTypeDomain As ICodedValueDomain
Private m_pExportMethodDomain As ICodedValueDomain
Private m_lngOpenDocId As Long
Private m_lngProjectId As Long

Const c_strModuleName As String = "ODNRProjectsDatabase"

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_lngOpenDocId = 0
    m_lngProjectId = 0
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Initialize " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    m_blnIsloaded = False
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".Class_Terminate " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Property Get ProjectCount() As Long
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        ProjectCount = m_pProjectsTable.RowCount(Nothing)
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectCount " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ActiveProjectCode() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) And (m_lngProjectId > 0) Then
        Dim pRow As IRow, lngFieldIdx As Long
        Set pRow = m_pProjectsTable.GetRow(m_lngProjectId)
        If Not (pRow Is Nothing) Then
            lngFieldIdx = m_pProjectsTable.FindField("PROJECTCODE")
            ActiveProjectCode = pRow.Value(lngFieldIdx)
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveProjectCode " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ActiveProjectName() As String
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) And (m_lngProjectId > 0) Then
        Dim pRow As IRow, lngFieldIdx As Long
        Set pRow = m_pProjectsTable.GetRow(m_lngProjectId)
        If Not (pRow Is Nothing) Then
            lngFieldIdx = m_pProjectsTable.FindField("PROJECTNAME")
            ActiveProjectName = pRow.Value(lngFieldIdx)
        End If
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ActiveProjectName " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ProjectsTable() As ITable
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set ProjectsTable = m_pProjectsTable
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectsTable " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get MapLayersTable() As ITable
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set MapLayersTable = m_pMapsTable
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".MapLayersTable " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get LayerObjectsTable() As ITable
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set LayerObjectsTable = m_pLayersTable
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerObjectsTable " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ProjectTypeDomain() As ICodedValueDomain
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set ProjectTypeDomain = m_pProjectTypeDomain
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectTypeDomain " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get QuadScaleDomain() As ICodedValueDomain
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set QuadScaleDomain = m_pQuadScaleDomain
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".QuadScaleDomain " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get MapTypeDomain() As ICodedValueDomain
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set MapTypeDomain = m_pMapTypeDomain
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".MapTypeDomain " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get ExportMethodDomain() As ICodedValueDomain
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set ExportMethodDomain = m_pExportMethodDomain
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".ExportMethodDomain " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Function LoadProject() As Boolean
    On Error GoTo ErrorHandler
    'Must OpenDocument first before loading project.
    If (m_blnIsloaded) And (m_lngOpenDocId > 0) Then
        Dim pRow As IRow, lngFieldIdx As Long
        Dim lngScale As Long, lngType As Long
        m_lngProjectId = m_lngOpenDocId
        m_lngOpenDocId = 0
        Set pRow = m_pProjectsTable.GetRow(m_lngProjectId)
        If Not (pRow Is Nothing) Then
            lngFieldIdx = m_pProjectsTable.FindField("PROJECTSCALE")
            lngScale = pRow.Value(lngFieldIdx)
            lngFieldIdx = m_pProjectsTable.FindField("PROJECTTYPE")
            lngType = pRow.Value(lngFieldIdx)
            Set gODNRProject = New ODNRProject
            gODNRProject.OpenProject m_pApp, lngScale, lngType
            ODNR_Common.VerifyProjectLayers
            ODNR_Common.UpdateProjectCombo
            ODNR_Common.LoadQuadCombo
        Else
            Set gODNRProject = Nothing
            m_lngProjectId = 0
        End If
    Else
        Set gODNRProject = Nothing
        m_lngProjectId = 0
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadProject " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub CloseProject()
    On Error GoTo ErrorHandler

    If (m_blnIsloaded) And (m_lngProjectId > 0) Then
        Set gODNRProject = Nothing
        m_lngProjectId = 0
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadProject " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Sub OpenDocument(strCode As String)
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        If (strCode <> Me.ActiveProjectCode) Then
            Dim pQf As IQueryFilter, pCursor As ICursor, pRow As IRow, lngFieldIdx As Long
            Dim strTempDir, pMemBlobStrm As IMemoryBlobStream, strFilePath As String
            Dim lngScale As Long, lngType As Long
            Set pQf = New QueryFilter
            pQf.WhereClause = "[PROJECTCODE] = '" & strCode & "'"
            strTempDir = TempDirectory
            If (m_pProjectsTable.RowCount(pQf) = 1) Then
                Set pCursor = m_pProjectsTable.Search(pQf, False)
                Set pRow = pCursor.NextRow
                lngFieldIdx = m_pProjectsTable.FindField("MXDFILE")
                Set pMemBlobStrm = pRow.Value(lngFieldIdx)
                strFilePath = strTempDir & "\" & strCode & ".mxd"
                If (IsOpenDocumentPath(strFilePath) = False) Then
                    pMemBlobStrm.SaveToFile strFilePath
                    m_pApp.OpenDocument strFilePath
                    'Opening the document cause an event to fire in ArcMap ("m_pDocEvents_OpenDocument")
                    'in the Select_DataDir_Cmd class.  Once the document is
                    'finished opening it will attempt to create the project by
                    'calling LoadProject in this class.
                Else
                    'Load project is open doc event is not firing here.
                    Me.LoadProject
                End If
                m_lngOpenDocId = pRow.OID
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError True, c_strModuleName & ".OpenDocument " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function ProjectCodeList() As Collection
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Dim pCursor As ICursor, pRow As IRow, lngFieldIdx As Long
        Dim pTblSort As ITableSort, pCodeList As Collection
        If (m_pProjectsTable.RowCount(Nothing) > 0) Then
            Set pTblSort = New esriGeoDatabase.TableSort
            With pTblSort
                .Fields = "PROJECTCODE"
                .Ascending("PROJECTCODE") = True
                .CaseSensitive("PROJECTCODE") = False
                Set .Table = m_pProjectsTable
            End With
            pTblSort.Sort Nothing
            Set pCursor = pTblSort.Rows
            Set pRow = pCursor.NextRow
            lngFieldIdx = m_pProjectsTable.FindField("PROJECTCODE")
            Set pCodeList = New Collection
            Do While Not pRow Is Nothing
                pCodeList.Add Item:=pRow.Value(lngFieldIdx)
                Set pRow = pCursor.NextRow
            Loop
            Set ProjectCodeList = pCodeList
        End If
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".ProjectCodeList " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function LoadDatabase(pApp As esriFramework.IApplication, pWs As IWorkspace) As Boolean
    On Error GoTo ErrorHandler
    
    If (Not pWs Is Nothing) And (Not pApp Is Nothing) Then
        Dim pDs As IDataset, pEnumDs As IEnumDataset, pTbl As ITable
        Dim pWsDomains As IWorkspaceDomains, pDomain As ICodedValueDomain
        Dim pTypeDomain As ICodedValueDomain, pScaleDomain As ICodedValueDomain, pMapTypeDomain As ICodedValueDomain
        Dim pExportMethodDomain As ICodedValueDomain, pLayerTypeDomain As ICodedValueDomain
        Dim blnProjectsLoaded As Boolean, blnMapsLoaded As Boolean, blnLayersLoaded As Boolean
        Set m_pApp = pApp
        If (pWs.Type = esriLocalDatabaseWorkspace) Or (pWs.Type = esriRemoteDatabaseWorkspace) Then
            Set pEnumDs = pWs.Datasets(esriDTTable)
            Set pDs = pEnumDs.Next
            Do While Not pDs Is Nothing
                If (pDs.Name = "ODNRProjects") Then
                    Set pTbl = pDs
                    If (pTbl.FindField("MXDFILE") > -1) And _
                        (pTbl.FindField("PROJECTNAME") > -1) And _
                        (pTbl.FindField("PROJECTTYPE") > -1) And _
                        (pTbl.FindField("PROJECTSCALE") > -1) And _
                        (pTbl.FindField("PROJECTCODE") > -1) And _
                        (pTbl.FindField("MXDFILEDATE") > -1) And _
                        (pTbl.FindField("DESCRIPTION") > -1) Then
                        Set pWsDomains = pWs
                        Set pDomain = pWsDomains.DomainByName("ProjectType")
                        If Not (pDomain Is Nothing) Then
                            Set pTypeDomain = pDomain
                        End If
                        Set pDomain = pWsDomains.DomainByName("ProjectScale")
                        If Not (pDomain Is Nothing) Then
                            Set pScaleDomain = pDomain
                        End If
                        Set pDomain = pWsDomains.DomainByName("MapType")
                        If Not (pDomain Is Nothing) Then
                            Set pMapTypeDomain = pDomain
                        End If
                        Set pDomain = pWsDomains.DomainByName("QuadExportMethod")
                        If Not (pDomain Is Nothing) Then
                            Set pExportMethodDomain = pDomain
                        End If
                        'Added 20051221, Jim McDonald
                        Set pDomain = pWsDomains.DomainByName("LayerType")
                        If Not (pDomain Is Nothing) Then
                            Set pLayerTypeDomain = pDomain
                        End If

                        If (Not pTypeDomain Is Nothing) And (Not pScaleDomain Is Nothing) And (Not pMapTypeDomain Is Nothing) And (Not pExportMethodDomain Is Nothing) And (Not pLayerTypeDomain Is Nothing) Then
                            Set m_pProjectsTable = pTbl
                            Set m_pProjectTypeDomain = pTypeDomain
                            Set m_pQuadScaleDomain = pScaleDomain
                            Set m_pMapTypeDomain = pMapTypeDomain
                            Set m_pExportMethodDomain = pExportMethodDomain
                            Set m_pLayerTypeDomain = pLayerTypeDomain
                            blnProjectsLoaded = True
                        End If
                    End If
                ElseIf (pDs.Name = "ODNRMapLayers") Then
                    Set pTbl = pDs
                    If (pTbl.FindField("PROJECTCODE") > -1) And _
                        (pTbl.FindField("MAPTYPE") > -1) And _
                        (pTbl.FindField("LAYEROBJECTID") > -1) Then
                        Set m_pMapsTable = pTbl
                        blnMapsLoaded = True
                    End If
                ElseIf (pDs.Name = "ODNRLayerObjects") Then
                    Set pTbl = pDs
                    If (pTbl.FindField("LAYERNAME") > -1) And _
                        (pTbl.FindField("LAYEROBJECT") > -1) And _
                        (pTbl.FindField("LAYERTYPE") > -1) And _
                        (pTbl.FindField("QUAD24KFIELD") > -1) And _
                        (pTbl.FindField("QUAD100KFIELD") > -1) Then
                        'Added 20051221, Jim McDonald
'                        Set pWsDomains = pWs
'                        Set pDomain = pWsDomains.DomainByName("LayerType")
'                        If Not (pDomain Is Nothing) Then
'                            Set pLayerTypeDomain = pDomain
'                        End If
                        Set m_pLayersTable = pTbl
                        blnLayersLoaded = True
                    End If
                End If
                Set pDs = pEnumDs.Next
            Loop
        End If
    End If
    If (blnProjectsLoaded) And (blnMapsLoaded) And (blnLayersLoaded) Then
        m_blnIsloaded = True
    End If
    LoadDatabase = m_blnIsloaded
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".LoadDatabase " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function TempDirectory() As String
    On Error GoTo ErrorHandler

    Dim strSysTempDir As String, strTempDir As String, pFSO As FileSystemObject
    Dim pFolder As Folder, blnValid As Boolean, intCount As Integer
    Set pFSO = New FileSystemObject
    strSysTempDir = Environ$("TEMP")
    If (strSysTempDir = "") Then strSysTempDir = Environ$("TMP")
    If (strSysTempDir = "") Then
        If (pFSO.FolderExists("C:\Temp")) Then
            strSysTempDir = "C:\Temp"
        Else
            pFSO.CreateFolder "C:\Temp"
            If (pFSO.FolderExists("C:\Temp")) Then strSysTempDir = "C:\Temp"
        End If
    End If
    blnValid = False
    intCount = 1
    strTempDir = strSysTempDir & "\Geology_Mxds"
    If (pFSO.FolderExists(strTempDir) = False) Then
        pFSO.CreateFolder strTempDir
    End If

    TempDirectory = strTempDir
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".TempDirectory " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function IsOpenDocumentPath(strFilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pTemplates As ITemplates
    Dim strOpenFilePath As String, strFileName As String
    Set pTemplates = m_pApp.Templates
    strOpenFilePath = pTemplates.Item(pTemplates.Count - 1)
    If (strFilePath = strOpenFilePath) Then
        IsOpenDocumentPath = True
    Else
        IsOpenDocumentPath = False
    End If
    
    Exit Function
ErrorHandler:
    HandleError True, c_strModuleName & ".IsOpenDocumentPath " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Property Get LayerTypeDomain() As ICodedValueDomain
    On Error GoTo ErrorHandler
    
    If (m_blnIsloaded) Then
        Set LayerTypeDomain = m_pLayerTypeDomain
    End If
    
    Exit Property
ErrorHandler:
    HandleError True, c_strModuleName & ".LayerTypeDomain " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

