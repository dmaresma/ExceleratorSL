Attribute VB_Name = "SemanticView"
' SEMANTIC VIEW / DYNAMIC PIVOT PANEL
' ------------------------------------------------------------------
' Implements the "SSAS-like" dimension/measure picker for Snowflake
' Semantic Views (SEMANTIC_VIEW SQL construct). Since Snowflake does
' not expose an MDX/XMLA endpoint, Excel's native PivotTable cannot be
' wired directly to a semantic view as an OLAP cube. Instead:
'   1. This module lists the Dimensions/Facts/Metrics available on a
'      chosen semantic view (via DESCRIBE SEMANTIC VIEW).
'   2. The user checks/selects which Dimensions and Metrics to bring
'      back (SemanticPivotForm - see setup instructions below).
'   3. A single SEMANTIC_VIEW(...) query is built and executed through
'      the existing Query.ExecuteSQL pipeline (same as every other
'      query in this add-in), landing pre-aggregated rows on a
'      worksheet as an Excel Table.
'   4. A native Excel PivotTable is created on top of that Table so the
'      user gets drag & drop dimension/measure slicing exactly like an
'      SSAS cube browser - but running against locally cached results.
'      Re-opening the form and picking different dimensions/metrics
'      re-issues the SEMANTIC_VIEW query and rebuilds the PivotTable.
'
' --- UserForm setup (build once in the VBA IDE; can't be hand-authored
'     as text because MSForms UserForms store their binary layout in
'     the paired .frx, not in the .frm) ---
'   Form name: SemanticPivotForm
'   Controls :
'     cbDatabases   (ComboBox)   - reuse FormCommon.getDatabasesCombobox
'     cbSchemas     (ComboBox)   - reuse FormCommon.getSchemasCombobox
'     cbSemanticViews (ComboBox) - populated by SemanticView.getSemanticViewsCombobox
'     lbDimensions  (ListBox, MultiSelectMulti, 2 columns: display/qualified ref)
'     lbMetrics     (ListBox, MultiSelectMulti, 2 columns: display/qualified ref)
'     btBuildPivot  (CommandButton) -> calls SemanticView.BuildAndInsertPivot
'     btCancel      (CommandButton) -> Me.Hide
'   Code-behind (paste into the form module):
'
'     Private Sub cbDatabases_Click()
'         Call StatusForm.execMethod("FormCommon", "getSchemasCombobox", cbSchemas, cbDatabases.value)
'         If cbSchemas.ListCount > 0 Then cbSchemas.ListIndex = 0
'     End Sub
'     Private Sub cbSchemas_Click()
'         Call SemanticView.getSemanticViewsCombobox(cbSemanticViews, cbDatabases.value, cbSchemas.value)
'         If cbSemanticViews.ListCount > 0 Then cbSemanticViews.ListIndex = 0
'     End Sub
'     Private Sub cbSemanticViews_Click()
'         Call SemanticView.populateDimensionsAndMetrics(lbDimensions, lbMetrics, cbDatabases.value, cbSchemas.value, cbSemanticViews.value)
'     End Sub
'     Private Sub UserForm_Initialize()
'         Call FormCommon.setUserFormPosition(Me)
'         Call FormCommon.initializeDBObjectsComboBoxes(cbDatabases, cbSchemas, cbSemanticViews) ' cbSemanticViews reused as 3rd combo, then repopulated below
'         Call SemanticView.getSemanticViewsCombobox(cbSemanticViews, cbDatabases.value, cbSchemas.value)
'         If cbSemanticViews.ListCount > 0 Then
'             cbSemanticViews.ListIndex = 0
'             Call SemanticView.populateDimensionsAndMetrics(lbDimensions, lbMetrics, cbDatabases.value, cbSchemas.value, cbSemanticViews.value)
'         End If
'     End Sub
'     Private Sub btCancel_Click()
'         Me.Hide
'     End Sub
'     Private Sub btBuildPivot_Click()
'         Dim dimsCSV As String, metricsCSV As String
'         dimsCSV = SemanticView.getSelectedRefsCSV(lbDimensions)
'         metricsCSV = SemanticView.getSelectedRefsCSV(lbMetrics)
'         If dimsCSV = "" And metricsCSV = "" Then
'             MsgBox "Please select at least one Dimension or Metric."
'             Exit Sub
'         End If
'         Me.Hide
'         Set StatusForm = Nothing
'         Call StatusForm.execMethod("SemanticView", "BuildAndInsertPivot", cbDatabases.value, cbSchemas.value, cbSemanticViews.value, dimsCSV, metricsCSV)
'     End Sub
'
' --- Ribbon wiring ---
'   Add a button in the ribbon customUI (Office RibbonX Editor on the
'   .xlam) with onAction="OpenSemanticPivotFormFromRibbon", then add
'   the matching Sub to PublicMacros.bas (already done).
' ------------------------------------------------------------------

Dim dictDescribeCache As New Scripting.Dictionary ' caches DESCRIBE SEMANTIC VIEW output per db-schema-view

' Populates a combobox with the semantic views available in the given database/schema
Public Sub getSemanticViewsCombobox(ByRef cbViews As ComboBox, database As String, schema As String)
    Dim sql As String
    Dim arrViews As Variant

    Call StatusForm.Update_Status("Getting Semantic Views...")
    cbViews.Clear
    On Error GoTo ErrorHandlerNoSemanticViews
    sql = "show semantic views in schema """ & database & """.""" & schema & """"
    Utils.execSQLFireAndForget (sql)
    sql = "select ""name"" from table(result_scan(last_query_id())) order by 1"
    arrViews = Utils.execSQLToArray(sql)
    For i = LBound(arrViews) To UBound(arrViews, 2)
        cbViews.AddItem (arrViews(0, i))
    Next i
ErrorHandlerNoSemanticViews:
    StatusForm.Hide
End Sub

' Runs (and caches) DESCRIBE SEMANTIC VIEW for the given view and returns the raw 2D array
' Columns (0-based): 0=object_kind, 1=object_name, 2=parent_entity, 3=property, 4=property_value
Private Function getDescribeArray(database As String, schema As String, view As String) As Variant
    Dim sql As String
    Dim key As String

    key = database & "-" & schema & "-" & view
    If dictDescribeCache.Exists(key) Then
        getDescribeArray = dictDescribeCache(key)
    Else
        sql = "describe semantic view """ & database & """.""" & schema & """.""" & view & """"
        getDescribeArray = Utils.execSQLToArray(sql)
        dictDescribeCache.Add Item:=getDescribeArray, key:=key
    End If
End Function

' Fills two listboxes with the Dimensions and Metrics defined on the semantic view.
' Each row added has 2 columns: column 0 = friendly display name ("logical_table.name"),
' column 1 = qualified reference to use in the SEMANTIC_VIEW(...) SQL clause.
Public Sub populateDimensionsAndMetrics(ByRef lbDimensions As ListBox, ByRef lbMetrics As ListBox, database As String, schema As String, view As String)
    Call StatusForm.Update_Status("Getting Dimensions and Metrics for " & view & "...")
    On Error GoTo ErrorHandlerNoMetadata
    Call fillListBoxFromDescribe(lbDimensions, database, schema, view, "DIMENSION")
    Call fillListBoxFromDescribe(lbMetrics, database, schema, view, "METRIC")
ErrorHandlerNoMetadata:
    StatusForm.Hide
End Sub

Private Sub fillListBoxFromDescribe(ByRef lb As ListBox, database As String, schema As String, view As String, objectKind As String)
    Dim arrDescribe As Variant
    Dim dictSeen As New Scripting.Dictionary
    Dim objectName As String
    Dim parentEntity As String
    Dim uniqueKey As String

    lb.Clear
    lb.ColumnCount = 2
    lb.ColumnWidths = "200,0" ' hide the 2nd (qualified ref) column, used only for building SQL

    arrDescribe = getDescribeArray(database, schema, view)
    If IsEmpty(arrDescribe) Then Exit Sub

    For i = LBound(arrDescribe, 2) To UBound(arrDescribe, 2)
        If UCase(arrDescribe(0, i)) = objectKind Then
            objectName = arrDescribe(1, i)
            parentEntity = arrDescribe(2, i)
            uniqueKey = parentEntity & "." & objectName
            If Not dictSeen.Exists(uniqueKey) Then
                dictSeen.Add Item:=True, key:=uniqueKey
                lb.AddItem
                lb.list(lb.ListCount - 1, 0) = uniqueKey
                lb.list(lb.ListCount - 1, 1) = uniqueKey
            End If
        End If
    Next i
End Sub

' Returns a comma separated list of the qualified references (column 1) for all selected rows in a listbox
Public Function getSelectedRefsCSV(ByRef lb As ListBox) As String
    Dim result As String
    result = ""
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            result = result & ", " & lb.list(i, 1)
        End If
    Next i
    If result <> "" Then
        result = Right(result, Len(result) - 2) ' strip leading ", "
    End If
    getSelectedRefsCSV = result
End Function

' Builds the SEMANTIC_VIEW(...) SQL statement for the chosen dimensions/metrics
Public Function BuildSemanticViewSQL(database As String, schema As String, view As String, dimensionsCSV As String, metricsCSV As String) As String
    Dim sql As String
    sql = "SELECT * FROM SEMANTIC_VIEW(" & vbCrLf
    sql = sql & "  """ & database & """.""" & schema & """.""" & view & """" & vbCrLf
    If metricsCSV <> "" Then
        sql = sql & "  METRICS " & metricsCSV & vbCrLf
    End If
    If dimensionsCSV <> "" Then
        sql = sql & "  DIMENSIONS " & dimensionsCSV & vbCrLf
    End If
    sql = sql & ")"
    BuildSemanticViewSQL = sql
End Function

' Executes the SEMANTIC_VIEW query (through the standard Query.ExecuteSQL pipeline so it
' benefits from the same status form, error handling and Excel Table formatting as any other
' query) then adds a native PivotTable on top of the results, with Dimensions on Rows and
' Metrics as Values - giving the SSAS-cube-like drag & drop experience for further slicing.
Public Sub BuildAndInsertPivot(database As String, schema As String, view As String, dimensionsCSV As String, metricsCSV As String)
    Dim sql As String
    Dim resultsWorksheet As Worksheet

    sql = BuildSemanticViewSQL(database, schema, view, dimensionsCSV, metricsCSV)
    Call Query.ExecuteSQL(sql)

    If CustomRange(sgRangeResultsWorksheet) = "" Then
        Set resultsWorksheet = ActiveSheet
    Else
        Set resultsWorksheet = Utils.getWorksheet(CustomRange(sgRangeResultsWorksheet))
    End If

    Call InsertPivotFromWorksheet(resultsWorksheet, dimensionsCSV, metricsCSV)
End Sub

' Creates (or replaces) a PivotTable on a new sheet, sourced from the flattened SEMANTIC_VIEW
' results, with each selected dimension as a Row field and each selected metric as a Value field.
Private Sub InsertPivotFromWorksheet(sourceWs As Worksheet, dimensionsCSV As String, metricsCSV As String)
    Dim pivotWs As Worksheet
    Dim pivotCache As pivotCache
    Dim pivotTbl As pivotTable
    Dim pivotWsName As String
    Dim arrDims() As String
    Dim arrMetrics() As String
    Dim fieldName As String

    On Error GoTo ErrorHandlerPivot

    pivotWsName = "Pivot_" & sourceWs.CodeName
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(pivotWsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandlerPivot
    Set pivotWs = Worksheets.Add(After:=sourceWs)
    pivotWs.name = pivotWsName

    Set pivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceWs.UsedRange)
    Set pivotTbl = pivotCache.CreatePivotTable(TableDestination:=pivotWs.range("A3"), TableName:="SemanticPivot_" & pivotWs.CodeName)

    If dimensionsCSV <> "" Then
        arrDims = Split(dimensionsCSV, ",")
        For i = LBound(arrDims) To UBound(arrDims)
            fieldName = getResultColumnNameFromRef(pivotTbl, Trim(arrDims(i)))
            If fieldName <> "" Then
                pivotTbl.PivotFields(fieldName).Orientation = xlRowField
            End If
        Next i
    End If

    If metricsCSV <> "" Then
        arrMetrics = Split(metricsCSV, ",")
        For i = LBound(arrMetrics) To UBound(arrMetrics)
            fieldName = getResultColumnNameFromRef(pivotTbl, Trim(arrMetrics(i)))
            If fieldName <> "" Then
                pivotTbl.AddDataField pivotTbl.PivotFields(fieldName), , xlSum
            End If
        Next i
    End If

    pivotWs.Activate
    Exit Sub
ErrorHandlerPivot:
    Call Utils.handleError("Error building PivotTable from Semantic View results. ", err)
End Sub

' The SEMANTIC_VIEW() query returns columns named after the unqualified dimension/metric name
' (e.g. "ORDERS.ORDER_DATE" -> column "ORDER_DATE"). This maps a qualified reference back to the
' actual PivotTable field name, tolerating a name clash by falling back to the qualified form.
Private Function getResultColumnNameFromRef(pivotTbl As pivotTable, qualifiedRef As String) As String
    Dim unqualifiedName As String
    Dim dotPos As Integer

    dotPos = InStrRev(qualifiedRef, ".")
    If dotPos > 0 Then
        unqualifiedName = Mid(qualifiedRef, dotPos + 1)
    Else
        unqualifiedName = qualifiedRef
    End If

    On Error Resume Next
    getResultColumnNameFromRef = ""
    If Not pivotTbl.PivotFields(unqualifiedName) Is Nothing Then
        getResultColumnNameFromRef = unqualifiedName
    End If
    If getResultColumnNameFromRef = "" Then
        If Not pivotTbl.PivotFields(qualifiedRef) Is Nothing Then
            getResultColumnNameFromRef = qualifiedRef
        End If
    End If
End Function

' Clears the DESCRIBE SEMANTIC VIEW metadata cache (call if a semantic view definition changes)
Public Sub dropSemanticViewCache()
    dictDescribeCache.RemoveAll
End Sub
