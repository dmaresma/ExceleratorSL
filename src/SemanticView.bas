Attribute VB_Name = "SemanticView"
' SEMANTIC VIEW SUPPORT
' ------------------------------------------------------------------
' Snowflake Semantic Views are a distinct object kind. They are not listed in
' INFORMATION_SCHEMA.TABLES, they have no INFORMATION_SCHEMA.COLUMNS entry, and
' they cannot be read with a plain SELECT ... FROM. They are:
'   - listed    with SHOW SEMANTIC VIEWS
'   - described with DESCRIBE SEMANTIC VIEW
'   - queried   with SEMANTIC_VIEW( <view> FACTS ... DIMENSIONS ... )
'
' Rather than a dedicated form, they are exposed through the existing Execute SQL
' form, which already carries the database/schema/object comboboxes, the column
' picker, the saved searches and the execution pipeline:
'   - FormCommon.getTablesCombobox appends them to the Table / View dropdown
'   - FormCommon.getColumnArray feeds the column picker from DESCRIBE
'   - SQLForm builds a SEMANTIC_VIEW(...) statement instead of a SELECT ... FROM
'
' SHOW and DESCRIBE both go through execSQLFireAndForget + RESULT_SCAN, the same
' pattern getDatabasesCombobox uses, because the ODBC driver does not expose the
' result set of those commands directly.
' ------------------------------------------------------------------

Public Const sgSemanticObjectFact As String = "FACT"
Public Const sgSemanticObjectDimension As String = "DIMENSION"

Dim dictViewsCache As New Scripting.Dictionary    ' "database-schema" -> array of semantic view names
Dim dictObjectsCache As New Scripting.Dictionary  ' "database-schema-view" -> array of (reference, object kind)


' Utils.execSQLToArray leaves its return value as an uninitialised array when the query
' returns no row. That is neither Empty nor safe to pass to LBound, so every caller has
' to go through this guard before looping.
Public Function hasRows(arr As Variant) As Boolean
    On Error GoTo NoRows
    If IsEmpty(arr) Then
        hasRows = False
        Exit Function
    End If
    hasRows = (UBound(arr, 2) >= LBound(arr, 2))
    Exit Function
NoRows:
    hasRows = False
End Function


' Clears the semantic view metadata caches. Called from FormCommon's cache dropping subs
' so a changed semantic view definition is picked up like any other object change.
Public Sub dropSemanticViewCache()
    dictViewsCache.RemoveAll
    dictObjectsCache.RemoveAll
End Sub


' Returns a 2 dimensional array whose row 0 holds the names of the semantic views in the
' given database/schema, or Empty when there are none.
Public Function getSemanticViewNames(database As String, schema As String) As Variant
    Dim sql As String
    Dim key As String

    key = database & "-" & schema
    If dictViewsCache.Exists(key) Then
        getSemanticViewNames = dictViewsCache(key)
        Exit Function
    End If

    ' A role without access, or an account without semantic views, makes SHOW fail. That must
    ' not take down the Table / View dropdown, so the failure is cached as "none" instead.
    On Error GoTo ErrorHandlerNoSemanticViews
    sql = "show semantic views in schema """ & database & """.""" & schema & """"
    Utils.execSQLFireAndForget (sql)
    sql = "WITH sv (name) as (select ""name"" from table(result_scan(last_query_id()))) " & _
          "select name from sv order by 1"
    getSemanticViewNames = Utils.execSQLToArray(sql)
    dictViewsCache.Add Item:=getSemanticViewNames, key:=key
    Exit Function

ErrorHandlerNoSemanticViews:
    getSemanticViewNames = Empty
    If Not dictViewsCache.Exists(key) Then
        dictViewsCache.Add Item:=Empty, key:=key
    End If
End Function


' True when the given object name is a semantic view of that database/schema. This is what
' every caller branches on to decide between the regular table path and the semantic path.
Public Function isSemanticView(database As String, schema As String, name As String) As Boolean
    Dim arrViews As Variant

    isSemanticView = False
    If name = "" Then Exit Function

    arrViews = getSemanticViewNames(database, schema)
    If Not hasRows(arrViews) Then Exit Function

    For i = LBound(arrViews, 2) To UBound(arrViews, 2)
        If UCase(arrViews(0, i)) = UCase(name) Then
            isSemanticView = True
            Exit Function
        End If
    Next i
End Function


' Returns the selectable objects of a semantic view as a 2 dimensional array:
'   (0, i) = qualified reference used in SEMANTIC_VIEW(...), e.g. "FACT.QTY_DELIVERED"
'   (1, i) = object kind, "DIMENSION" or "FACT"
' That 2 column shape is exactly what SelectColumnsForm already displays for a regular
' table (column name + data type), so the column picker needs no layout change.
Public Function getSemanticObjects(database As String, schema As String, view As String) As Variant
    Dim sql As String
    Dim key As String

    key = database & "-" & schema & "-" & view
    If dictObjectsCache.Exists(key) Then
        getSemanticObjects = dictObjectsCache(key)
        Exit Function
    End If

    On Error GoTo ErrorHandlerNoMetadata
    sql = "describe semantic view """ & database & """.""" & schema & """.""" & view & """"
    Utils.execSQLFireAndForget (sql)
    ' DESCRIBE returns one row per object *property* (TABLE, EXPRESSION, DATA_TYPE,
    ' ACCESS_MODIFIER), hence the DISTINCT on the object itself.
    sql = "WITH d (kind, nm, parent) as (" & _
          "select ""object_kind"", ""object_name"", ""parent_entity"" " & _
          "from table(result_scan(last_query_id()))) " & _
          "select distinct iff(parent is null, nm, parent || '.' || nm) as ref, kind from d " & _
          "where kind in ('" & sgSemanticObjectDimension & "','" & sgSemanticObjectFact & "') " & _
          "order by kind, ref"
    getSemanticObjects = Utils.execSQLToArray(sql)
    dictObjectsCache.Add Item:=getSemanticObjects, key:=key
    Exit Function

ErrorHandlerNoMetadata:
    getSemanticObjects = Empty
End Function


' Builds a SEMANTIC_VIEW(...) statement from a comma separated list of qualified references,
' as produced by SelectColumnsForm. Each reference is routed to the FACTS or the DIMENSIONS
' clause according to its object kind - the two clauses are not interchangeable.
Public Function BuildSemanticViewSQL(database As String, schema As String, view As String, refsCSV As String) As String
    Dim arrObjects As Variant
    Dim arrRefs() As String
    Dim dictKind As New Scripting.Dictionary
    Dim factsCSV As String
    Dim dimensionsCSV As String
    Dim ref As String

    arrObjects = getSemanticObjects(database, schema, view)
    If hasRows(arrObjects) Then
        For i = LBound(arrObjects, 2) To UBound(arrObjects, 2)
            dictKind(UCase(arrObjects(0, i))) = UCase(arrObjects(1, i))
        Next i
    End If

    arrRefs = Split(refsCSV, ",")
    For i = LBound(arrRefs) To UBound(arrRefs)
        ref = Trim(arrRefs(i))
        If ref <> "" Then
            ' An unknown reference can only come from a stale cache. Sending it as a dimension
            ' lets Snowflake report the real problem instead of dropping it silently.
            ' VBA's And is not short-circuiting, so Exists has to gate the lookup in its own
            ' If - reading a missing key would otherwise insert it into the dictionary.
            If dictKind.Exists(UCase(ref)) Then
                If dictKind(UCase(ref)) = sgSemanticObjectFact Then
                    factsCSV = factsCSV & ", " & ref
                Else
                    dimensionsCSV = dimensionsCSV & ", " & ref
                End If
            Else
                dimensionsCSV = dimensionsCSV & ", " & ref
            End If
        End If
    Next i

    BuildSemanticViewSQL = buildStatement(database, schema, view, factsCSV, dimensionsCSV)
End Function


' "All Columns" on a semantic view means every fact and every dimension: SEMANTIC_VIEW(...)
' has no SELECT * equivalent, the objects have to be named explicitly.
Public Function BuildSemanticViewSQLAllObjects(database As String, schema As String, view As String) As String
    Dim arrObjects As Variant
    Dim factsCSV As String
    Dim dimensionsCSV As String

    arrObjects = getSemanticObjects(database, schema, view)
    If Not hasRows(arrObjects) Then
        BuildSemanticViewSQLAllObjects = ""
        Exit Function
    End If

    For i = LBound(arrObjects, 2) To UBound(arrObjects, 2)
        If UCase(arrObjects(1, i)) = sgSemanticObjectFact Then
            factsCSV = factsCSV & ", " & arrObjects(0, i)
        Else
            dimensionsCSV = dimensionsCSV & ", " & arrObjects(0, i)
        End If
    Next i

    BuildSemanticViewSQLAllObjects = buildStatement(database, schema, view, factsCSV, dimensionsCSV)
End Function


Private Function buildStatement(database As String, schema As String, view As String, factsCSV As String, dimensionsCSV As String) As String
    Dim sql As String

    sql = "SELECT * FROM SEMANTIC_VIEW(" & vbCrLf
    sql = sql & "  """ & database & """.""" & schema & """.""" & view & """" & vbCrLf
    If factsCSV <> "" Then
        sql = sql & "  FACTS " & stripLeadingComma(factsCSV) & vbCrLf
    End If
    If dimensionsCSV <> "" Then
        sql = sql & "  DIMENSIONS " & stripLeadingComma(dimensionsCSV) & vbCrLf
    End If
    sql = sql & ")"
    buildStatement = sql
End Function


Private Function stripLeadingComma(csv As String) As String
    stripLeadingComma = Trim(csv)
    If Left(stripLeadingComma, 1) = "," Then
        stripLeadingComma = Trim(Mid(stripLeadingComma, 2))
    End If
End Function
