Attribute VB_Name = "PMTool"
Option Explicit

Private Const ERROR_BASE As Long = vbObjectError + 2400
Private Const ITEM_CAPACITY As Long = 2000
Private Const ID_COUNTER_MAX As Long = 999999999
Private Const EXPORT_CHARACTER_MAX As Long = 5000000
Private Const EXPORT_BYTE_MAX As Long = 20000000

Private Function Tbl(ByVal tableName As String) As Excel.ListObject
    Dim requested As String
    requested = Trim$(tableName)
    If Len(requested) = 0 Then
        Err.Raise ERROR_BASE + 1, "PMTool.Tbl", "A table name is required."
    End If

    Dim sheet As Excel.Worksheet
    Dim table As Excel.ListObject
    Dim found As Excel.ListObject
    Dim matches As Long
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, requested, vbTextCompare) = 0 Then
                matches = matches + 1
                Set found = table
            End If
        Next table
    Next sheet

    If matches = 0 Then
        Err.Raise ERROR_BASE + 2, "PMTool.Tbl", _
                  "Required table '" & requested & "' was not found."
    End If
    If matches > 1 Then
        Err.Raise ERROR_BASE + 3, "PMTool.Tbl", _
                  "Table name '" & requested & "' is duplicated."
    End If
    Set Tbl = found
End Function

Private Function ColOf(ByVal table As Excel.ListObject, _
                      ByVal header As String) As Long
    If table Is Nothing Then
        Err.Raise ERROR_BASE + 4, "PMTool.ColOf", "A table object is required."
    End If
    Dim requested As String
    requested = Trim$(header)
    If Len(requested) = 0 Then
        Err.Raise ERROR_BASE + 5, "PMTool.ColOf", "A column header is required."
    End If

    Dim index As Long
    Dim matches As Long
    Dim foundIndex As Long
    For index = 1 To table.ListColumns.Count
        If StrComp(table.ListColumns(index).Name, requested, vbTextCompare) = 0 Then
            matches = matches + 1
            foundIndex = index
        End If
    Next index

    If matches = 0 Then
        Err.Raise ERROR_BASE + 6, "PMTool.ColOf", _
                  "Required column '" & requested & "' is missing from " & _
                  table.Name & "."
    End If
    If matches > 1 Then
        Err.Raise ERROR_BASE + 7, "PMTool.ColOf", _
                  "Column '" & requested & "' is duplicated in " & _
                  table.Name & "."
    End If
    ColOf = foundIndex
End Function

Private Function ConfigRange(ByVal settingName As String) As Excel.Range
    Dim requested As String
    requested = Trim$(settingName)
    If Len(requested) = 0 Then
        Err.Raise ERROR_BASE + 8, "PMTool.ConfigRange", _
                  "A configuration name is required."
    End If

    Dim target As Excel.Range
    On Error GoTo accessFailure
    Set target = ThisWorkbook.Names(requested).RefersToRange
    On Error GoTo 0

    If target Is Nothing Then
        Err.Raise ERROR_BASE + 9, "PMTool.ConfigRange", _
                  "Configuration name '" & requested & "' has no cell target."
    End If
    If target.Cells.CountLarge <> 1 Then
        Err.Raise ERROR_BASE + 10, "PMTool.ConfigRange", _
                  "Configuration name '" & requested & "' must refer to one cell."
    End If
    Set ConfigRange = target
    Exit Function

accessFailure:
    Dim failureDescription As String
    failureDescription = Err.description
    On Error GoTo 0
    Err.Raise ERROR_BASE + 11, "PMTool.ConfigRange", _
              "Configuration name '" & requested & "' is unavailable: " & _
              failureDescription
End Function

Private Function CfgVal(ByVal settingName As String) As Variant
    Dim value As Variant
    value = ConfigRange(settingName).value
    If IsError(value) Then
        Err.Raise ERROR_BASE + 12, "PMTool.CfgVal", _
                  "Configuration '" & settingName & "' contains an Excel error."
    End If
    CfgVal = value
End Function

Private Function CfgLong(ByVal settingName As String) As Long
    Dim value As Variant
    value = CfgVal(settingName)
    If IsEmpty(value) Or IsNull(value) Then
        Err.Raise ERROR_BASE + 13, "PMTool.CfgLong", _
                  "Configuration '" & settingName & "' is blank."
    End If
    If Len(Trim$(CStr(value))) = 0 Then
        Err.Raise ERROR_BASE + 13, "PMTool.CfgLong", _
                  "Configuration '" & settingName & "' is blank."
    End If
    If Not IsNumeric(value) Then
        Err.Raise ERROR_BASE + 14, "PMTool.CfgLong", _
                  "Configuration '" & settingName & "' must be an integer."
    End If
    Dim numericValue As Double
    numericValue = CDbl(value)
    If numericValue <> Fix(numericValue) Or numericValue < 1 Or _
       numericValue > 2147483646# Then
        Err.Raise ERROR_BASE + 15, "PMTool.CfgLong", _
                  "Configuration '" & settingName & _
                  "' must be an integer from 1 to 2147483646."
    End If
    CfgLong = CLng(numericValue)
End Function

Private Function CfgText(ByVal settingName As String) As String
    Dim value As Variant
    value = CfgVal(settingName)
    If IsEmpty(value) Or IsNull(value) Then
        Err.Raise ERROR_BASE + 16, "PMTool.CfgText", _
                  "Configuration '" & settingName & "' is blank."
    End If
    Dim text As String
    text = CStr(value)
    If Len(text) = 0 Or text <> Trim$(text) Then
        Err.Raise ERROR_BASE + 17, "PMTool.CfgText", _
                  "Configuration '" & settingName & _
                  "' must contain trimmed text."
    End If
    CfgText = text
End Function

Private Function CfgIdPrefix(ByVal settingName As String) As String
    Dim prefix As String
    prefix = CfgText(settingName)
    If Len(prefix) > 8 Then
        Err.Raise ERROR_BASE + 49, "PMTool.CfgIdPrefix", _
                  "Configuration '" & settingName & _
                  "' must contain no more than 8 characters."
    End If
    If InStr(1, prefix, ",", vbBinaryCompare) > 0 Or _
       InStr(1, prefix, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, prefix, vbLf, vbBinaryCompare) > 0 Or _
       InStr(1, prefix, vbTab, vbBinaryCompare) > 0 Then
        Err.Raise ERROR_BASE + 50, "PMTool.CfgIdPrefix", _
                  "Configuration '" & settingName & _
                  "' cannot contain commas, tabs or line breaks."
    End If
    CfgIdPrefix = prefix
End Function

Private Sub SetCfg(ByVal settingName As String, ByVal value As Variant)
    If IsError(value) Then
        Err.Raise ERROR_BASE + 18, "PMTool.SetCfg", _
                  "An Excel error cannot be written to configuration."
    End If
    ConfigRange(settingName).value = value
End Sub

Public Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsError(value) Then
        Err.Raise ERROR_BASE + 19, "PMTool.IsBlankValue", _
                  "An Excel error cannot be treated as a blank value."
    End If
    If IsEmpty(value) Or IsNull(value) Then
        IsBlankValue = True
    ElseIf VarType(value) = vbString Then
        IsBlankValue = (Len(Trim$(CStr(value))) = 0)
    End If
End Function

Private Function ConfiguredText(ByVal value As Variant, _
                                ByVal fieldName As String, _
                                ByVal sourceName As String) As String
    If IsError(value) Then
        Err.Raise ERROR_BASE + 90, sourceName, _
                  fieldName & " contains an Excel error."
    End If
    If IsBlankValue(value) Then Exit Function
    If VarType(value) <> vbString Then
        Err.Raise ERROR_BASE + 83, sourceName, _
                  fieldName & " must contain text or be blank."
    End If
    Dim result As String
    result = CStr(value)
    If result <> Trim$(result) Then
        Err.Raise ERROR_BASE + 84, sourceName, _
                  fieldName & " cannot contain leading or trailing spaces."
    End If
    ConfiguredText = result
End Function

Private Function ConfiguredBoolean(ByVal value As Variant, _
                                   ByVal fieldName As String, _
                                   ByVal sourceName As String) As Boolean
    If IsError(value) Then
        Err.Raise ERROR_BASE + 98, sourceName, _
                  fieldName & " contains an Excel error."
    End If
    If VarType(value) <> vbBoolean Then
        Err.Raise ERROR_BASE + 99, sourceName, _
                  fieldName & " must contain a TRUE or FALSE checkbox value."
    End If
    ConfiguredBoolean = CBool(value)
End Function

Public Function ValidatedIdentifierText(ByVal value As Variant, _
                                        ByVal fieldName As String, _
                                        ByVal sourceName As String) As String
    Dim identifier As String
    identifier = ConfiguredText(value, fieldName, sourceName)
    If Len(identifier) = 0 Then
        Err.Raise ERROR_BASE + 85, sourceName, _
                  fieldName & " is required."
    End If
    If Len(identifier) > 255 Or _
       InStr(1, identifier, ",", vbBinaryCompare) > 0 Or _
       InStr(1, identifier, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, identifier, vbLf, vbBinaryCompare) > 0 Or _
       InStr(1, identifier, vbTab, vbBinaryCompare) > 0 Then
        Err.Raise ERROR_BASE + 86, sourceName, _
                  fieldName & _
                  " must contain at most 255 characters and cannot contain " & _
                  "commas, tabs or line breaks."
    End If
    ValidatedIdentifierText = identifier
End Function

Private Sub AccumulateTextValue(ByVal value As Variant, _
                                ByVal requested As String, _
                                ByVal fieldName As String, _
                                ByRef matches As Long, _
                                ByRef nonblankCount As Long, _
                                ByRef lastText As String)
    If IsError(value) Then
        Err.Raise ERROR_BASE + 66, "PMTool.AccumulateTextValue", _
                  fieldName & " contains an Excel error."
    End If
    If IsEmpty(value) Or IsNull(value) Then Exit Sub
    If VarType(value) <> vbString Then
        Err.Raise ERROR_BASE + 67, "PMTool.AccumulateTextValue", _
                  fieldName & " must contain text or blank cells."
    End If

    Dim text As String
    text = CStr(value)
    If Len(text) = 0 Then Exit Sub
    If text <> Trim$(text) Then
        Err.Raise ERROR_BASE + 82, "PMTool.AccumulateTextValue", _
                  fieldName & " contains leading or trailing spaces."
    End If
    nonblankCount = nonblankCount + 1
    lastText = text
    If StrComp(text, requested, vbTextCompare) = 0 Then
        matches = matches + 1
    End If
End Sub

Private Sub TextRangeStats(ByVal values As Excel.Range, _
                           ByVal requested As String, _
                           ByVal fieldName As String, _
                           ByRef matches As Long, _
                           ByRef nonblankCount As Long, _
                           ByRef lastText As String)
    If values Is Nothing Then
        Err.Raise ERROR_BASE + 64, "PMTool.TextRangeStats", _
                  fieldName & " has no cell range."
    End If
    If values.Areas.Count <> 1 Then
        Err.Raise ERROR_BASE + 65, "PMTool.TextRangeStats", _
                  fieldName & " must use one contiguous cell range."
    End If

    matches = 0
    nonblankCount = 0
    lastText = ""
    Dim data As Variant
    data = values.Value2
    If values.CountLarge = 1 Then
        AccumulateTextValue data, requested, fieldName, _
                            matches, nonblankCount, lastText
        Exit Sub
    End If

    Dim rowIndex As Long
    Dim columnIndex As Long
    For rowIndex = LBound(data, 1) To UBound(data, 1)
        For columnIndex = LBound(data, 2) To UBound(data, 2)
            AccumulateTextValue data(rowIndex, columnIndex), _
                                requested, fieldName, matches, _
                                nonblankCount, lastText
        Next columnIndex
    Next rowIndex
End Sub

Public Function ExactTextCount(ByVal values As Excel.Range, _
                               ByVal requested As String, _
                               ByVal fieldName As String) As Long
    Dim matches As Long
    Dim nonblankCount As Long
    Dim lastText As String
    TextRangeStats values, requested, fieldName, _
                   matches, nonblankCount, lastText
    ExactTextCount = matches
End Function

Private Function TableTextRow(ByVal table As Excel.ListObject, _
                              ByVal identityHeader As String, _
                              ByVal requested As String, _
                              ByVal fieldName As String, _
                              ByVal sourceName As String) As Long
    If table Is Nothing Then
        Err.Raise ERROR_BASE + 100, sourceName, _
                  "The configuration table is unavailable."
    End If
    Dim identityIndex As Long
    identityIndex = ColOf(table, identityHeader)
    If table.DataBodyRange Is Nothing Then
        Err.Raise ERROR_BASE + 101, sourceName, _
                  table.Name & " has no configuration rows."
    End If

    Dim data As Variant
    data = table.DataBodyRange.Value2
    Dim rowIndex As Long
    Dim identity As String
    Dim matches As Long
    Dim matchedRow As Long
    For rowIndex = 1 To table.ListRows.Count
        identity = ConfiguredText( _
            data(rowIndex, identityIndex), _
            table.Name & "[" & identityHeader & "] row " & rowIndex, _
            sourceName)
        If Len(identity) = 0 Then
            Err.Raise ERROR_BASE + 102, sourceName, _
                      table.Name & " contains a blank " & identityHeader & "."
        End If
        If StrComp(identity, requested, vbTextCompare) = 0 Then
            matches = matches + 1
            matchedRow = rowIndex
        End If
    Next rowIndex
    If matches = 0 Then
        Err.Raise ERROR_BASE + 103, sourceName, _
                  fieldName & " '" & requested & "' is not defined in Config."
    End If
    If matches > 1 Then
        Err.Raise ERROR_BASE + 104, sourceName, _
                  fieldName & " '" & requested & "' is duplicated in Config."
    End If
    TableTextRow = matchedRow
End Function

Public Sub ItemStatusRoles(ByVal value As Variant, _
                           ByRef isActive As Boolean, _
                           ByRef isDone As Boolean, _
                           ByRef isCancelled As Boolean)
    isActive = False
    isDone = False
    isCancelled = False
    Dim requested As String
    requested = ConfiguredText( _
        value, "Items status", "PMTool.ItemStatusRoles")
    If Len(requested) = 0 Then Exit Sub
    Dim table As Excel.ListObject
    Set table = Tbl("tblStatuses")
    Dim rowIndex As Long
    rowIndex = TableTextRow( _
        table, "Status", requested, "Items status", "PMTool.ItemStatusRoles")
    Dim data As Variant
    data = table.DataBodyRange.Value2
    isActive = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsActive")), _
        "Status '" & requested & "' IsActive", "PMTool.ItemStatusRoles")
    isDone = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsDone")), _
        "Status '" & requested & "' IsDone", "PMTool.ItemStatusRoles")
    isCancelled = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsCancelled")), _
        "Status '" & requested & "' IsCancelled", "PMTool.ItemStatusRoles")
    Dim isDeleted As Boolean
    isDeleted = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsDeleted")), _
        "Status '" & requested & "' IsDeleted", "PMTool.ItemStatusRoles")
    If isActive And isDone Then
        Err.Raise ERROR_BASE + 105, "PMTool.ItemStatusRoles", _
                  "Status '" & requested & _
                  "' cannot be both active and done."
    End If
    If isCancelled And Not isDone Then
        Err.Raise ERROR_BASE + 106, "PMTool.ItemStatusRoles", _
                  "Status '" & requested & _
                  "' must be marked done when it is cancelled."
    End If
    If isDeleted And (isActive Or Not isDone Or Not isCancelled) Then
        Err.Raise ERROR_BASE + 131, "PMTool.ItemStatusRoles", _
                  "Status '" & requested & _
                  "' must be inactive, done and cancelled when it is deleted."
    End If
End Sub

Public Function ItemStatusIsDeleted(ByVal value As Variant) As Boolean
    Dim requested As String
    requested = ConfiguredText( _
        value, "Items status", "PMTool.ItemStatusIsDeleted")
    If Len(requested) = 0 Then Exit Function
    Dim table As Excel.ListObject
    Set table = Tbl("tblStatuses")
    Dim rowIndex As Long
    rowIndex = TableTextRow( _
        table, "Status", requested, "Items status", _
        "PMTool.ItemStatusIsDeleted")
    Dim data As Variant
    data = table.DataBodyRange.Value2
    ItemStatusIsDeleted = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsDeleted")), _
        "Status '" & requested & "' IsDeleted", _
        "PMTool.ItemStatusIsDeleted")
End Function

Public Function RaidStatusIsClosed(ByVal value As Variant) As Boolean
    Dim requested As String
    requested = ConfiguredText( _
        value, "RAID status", "PMTool.RaidStatusIsClosed")
    If Len(requested) = 0 Then Exit Function
    Dim table As Excel.ListObject
    Set table = Tbl("tblRaidStatuses")
    Dim rowIndex As Long
    rowIndex = TableTextRow( _
        table, "RaidStatus", requested, "RAID status", _
        "PMTool.RaidStatusIsClosed")
    Dim data As Variant
    data = table.DataBodyRange.Value2
    RaidStatusIsClosed = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsClosed")), _
        "RAID status '" & requested & "' IsClosed", _
        "PMTool.RaidStatusIsClosed")
    Dim isDeleted As Boolean
    isDeleted = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsDeleted")), _
        "RAID status '" & requested & "' IsDeleted", _
        "PMTool.RaidStatusIsClosed")
    If isDeleted And Not RaidStatusIsClosed Then
        Err.Raise ERROR_BASE + 132, "PMTool.RaidStatusIsClosed", _
                  "RAID status '" & requested & _
                  "' must be closed when it is deleted."
    End If
End Function

Public Function RaidStatusIsDeleted(ByVal value As Variant) As Boolean
    Dim requested As String
    requested = ConfiguredText( _
        value, "RAID status", "PMTool.RaidStatusIsDeleted")
    If Len(requested) = 0 Then Exit Function
    Dim table As Excel.ListObject
    Set table = Tbl("tblRaidStatuses")
    Dim rowIndex As Long
    rowIndex = TableTextRow( _
        table, "RaidStatus", requested, "RAID status", _
        "PMTool.RaidStatusIsDeleted")
    Dim data As Variant
    data = table.DataBodyRange.Value2
    RaidStatusIsDeleted = ConfiguredBoolean( _
        data(rowIndex, ColOf(table, "IsDeleted")), _
        "RAID status '" & requested & "' IsDeleted", _
        "PMTool.RaidStatusIsDeleted")
End Function

Public Function IsBlockedDeliveryHealth(ByVal value As Variant) As Boolean
    Dim requested As String
    requested = ConfiguredText( _
        value, "Delivery Health", "PMTool.IsBlockedDeliveryHealth")
    If Len(requested) = 0 Then Exit Function

    Dim healthTable As Excel.ListObject
    Set healthTable = Tbl("tblDeliveryHealth")
    Dim healthRange As Excel.Range
    Set healthRange = healthTable.ListColumns( _
        ColOf(healthTable, "Delivery Health")).DataBodyRange
    Dim matches As Long
    Dim healthCount As Long
    Dim blockedValue As String
    TextRangeStats healthRange, requested, "Delivery Health", _
                   matches, healthCount, blockedValue
    If matches = 0 Then
        Err.Raise ERROR_BASE + 76, "PMTool.IsBlockedDeliveryHealth", _
                  "Delivery Health '" & requested & _
                  "' is not defined in Config."
    End If
    If matches > 1 Then
        Err.Raise ERROR_BASE + 77, "PMTool.IsBlockedDeliveryHealth", _
                  "Delivery Health '" & requested & _
                  "' is duplicated in Config."
    End If
    If healthCount < 4 Then
        Err.Raise ERROR_BASE + 79, "PMTool.IsBlockedDeliveryHealth", _
                  "Config requires at least four Delivery Health values."
    End If
    IsBlockedDeliveryHealth = _
        (StrComp(requested, blockedValue, vbTextCompare) = 0)
End Function

Public Function ItemTypeLevel(ByVal value As Variant) As Long
    ' Mirrors fnTypeLevel: the first configured match wins and blank,
    ' unknown or unusable types resolve to level 0.
    Dim requested As String
    requested = ConfiguredText( _
        value, "Items type", "PMTool.ItemTypeLevel")
    If Len(requested) = 0 Then Exit Function

    Dim table As Excel.ListObject
    Set table = Tbl("tblTypes")
    If table.DataBodyRange Is Nothing Then Exit Function
    Dim data As Variant
    data = table.DataBodyRange.Value2
    Dim nameIndex As Long
    Dim levelIndex As Long
    nameIndex = ColOf(table, "Type")
    levelIndex = ColOf(table, "Level")

    Dim rowIndex As Long
    Dim nameValue As Variant
    Dim levelValue As Variant
    For rowIndex = 1 To table.ListRows.Count
        nameValue = data(rowIndex, nameIndex)
        If IsError(nameValue) Then Exit Function
        If Not IsBlankValue(nameValue) Then
            If StrComp(Trim$(CStr(nameValue)), requested, _
                       vbTextCompare) = 0 Then
                levelValue = data(rowIndex, levelIndex)
                If IsError(levelValue) Then Exit Function
                If Not IsNumeric(levelValue) Then Exit Function
                If CDbl(levelValue) <> Fix(CDbl(levelValue)) Then Exit Function
                If levelValue < 1 Or levelValue > 6 Then Exit Function
                ItemTypeLevel = CLng(levelValue)
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Public Sub ApplyItemLevelPresentation(ByVal dataSheet As Excel.Worksheet, _
                                      ByVal rowNumber As Long, _
                                      ByVal titleColumn As Long, _
                                      ByVal level As Long)
    If dataSheet Is Nothing Then
        Err.Raise ERROR_BASE + 128, "PMTool.ApplyItemLevelPresentation", _
                  "A worksheet is required."
    End If
    If rowNumber < 1 Or titleColumn < 1 Then
        Err.Raise ERROR_BASE + 129, "PMTool.ApplyItemLevelPresentation", _
                  "The Items presentation target is invalid."
    End If
    If level < 0 Or level > 6 Then
        Err.Raise ERROR_BASE + 130, "PMTool.ApplyItemLevelPresentation", _
                  "Item levels run from 0 (unassigned) to 6."
    End If

    ' Level ramp; matches the HIERARCHY design tokens in the build.
    Dim titleSize As Double
    Dim levelRowHeight As Double
    Select Case level
        Case 1
            titleSize = 12
            levelRowHeight = 30
        Case 2
            titleSize = 11
            levelRowHeight = 28
        Case 3
            titleSize = 10.5
            levelRowHeight = 26
        Case Else
            titleSize = 10
            levelRowHeight = 24
    End Select

    dataSheet.Rows(rowNumber).RowHeight = levelRowHeight
    With dataSheet.Cells(rowNumber, titleColumn)
        If level >= 1 Then
            .IndentLevel = level - 1
        Else
            .IndentLevel = 0
        End If
        .Font.Bold = (level >= 1 And level <= 3)
        .Font.Size = titleSize
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
End Sub

Public Function NextUniqueIdInRange(ByVal idRange As Excel.Range, _
                                    ByVal fieldName As String, _
                                    ByVal prefixSetting As String, _
                                    ByVal counterSetting As String, _
                                    Optional ByVal padWidth As Long = 0) As String
    If idRange Is Nothing Then
        Err.Raise ERROR_BASE + 68, "PMTool.NextUniqueIdInRange", _
                  fieldName & " has no cell range."
    End If
    If idRange.Areas.Count <> 1 Then
        Err.Raise ERROR_BASE + 69, "PMTool.NextUniqueIdInRange", _
                  fieldName & " must use one contiguous cell range."
    End If
    If padWidth < 0 Or padWidth > 12 Then
        Err.Raise ERROR_BASE + 24, "PMTool.NextUniqueIdInRange", _
                  "ID padding must be between 0 and 12 characters."
    End If

    Dim prefix As String
    prefix = CfgIdPrefix(prefixSetting)
    Dim counter As Long
    counter = CfgLong(counterSetting)
    If counter >= ID_COUNTER_MAX Then
        Err.Raise ERROR_BASE + 26, "PMTool.NextUniqueIdInRange", _
                  "The ID counter has reached its supported maximum."
    End If
    Dim candidate As String
    Dim suffix As String
    Dim occurrences As Long

    Do
        If padWidth = 0 Then
            suffix = CStr(counter)
        Else
            suffix = Format$(counter, String$(padWidth, "0"))
        End If
        candidate = prefix & suffix
        If Len(candidate) > 255 Then
            Err.Raise ERROR_BASE + 25, "PMTool.NextUniqueIdInRange", _
                      "Generated ID exceeds Excel's 255-character cell limit."
        End If
        occurrences = ExactTextCount(idRange, candidate, fieldName)
        If occurrences = 0 Then Exit Do
        If counter + 1 >= ID_COUNTER_MAX Then
            Err.Raise ERROR_BASE + 26, "PMTool.NextUniqueIdInRange", _
                      "The ID counter has reached its supported maximum."
        End If
        counter = counter + 1
    Loop

    SetCfg counterSetting, counter + 1
    NextUniqueIdInRange = candidate
End Function

Private Sub CalculateExportView()
    Tbl("tblItems").Range.Calculate
    Tbl("tblRAID").Range.Calculate
    ThisWorkbook.Worksheets("Calc").Calculate
    ThisWorkbook.Worksheets("Overview").Calculate
End Sub

Private Function NormalizeMacMarkdownPath(ByVal path As String) As String
    Dim result As String
    result = path
    Dim extensionPosition As Long
    extensionPosition = InStrRev(result, ".")
    If extensionPosition > 0 Then
        Select Case LCase$(Mid$(result, extensionPosition))
            Case ".xlsx", ".xlsm", ".xlsb", ".xls"
                result = Left$(result, extensionPosition - 1)
        End Select
    End If
    NormalizeMacMarkdownPath = result
End Function

Private Function SelectedMarkdownPath(ByVal suggested As String) As String
    Dim selectedPath As Variant
#If Mac Then
    selectedPath = Application.GetSaveAsFilename( _
        InitialFileName:=suggested, _
        title:="Export project status", _
        ButtonText:="Export")
    If VarType(selectedPath) = vbBoolean Then Exit Function
    If Len(Trim$(CStr(selectedPath))) = 0 Then Exit Function
    SelectedMarkdownPath = EnsureMdExtension( _
        NormalizeMacMarkdownPath(CStr(selectedPath)))
#Else
    selectedPath = Application.GetSaveAsFilename( _
        InitialFileName:=suggested, _
        FileFilter:="Markdown (*.md),*.md", _
        title:="Export project status")
    If VarType(selectedPath) = vbBoolean Then Exit Function
    If Len(Trim$(CStr(selectedPath))) = 0 Then Exit Function
    SelectedMarkdownPath = EnsureMdExtension(CStr(selectedPath))
#End If
End Function

Public Sub ExportMarkdown()
    On Error GoTo exportFailure
    Dim exportStage As String
    Dim suggested As String
    suggested = "PM_Status_" & Format$(Date, "yyyy-mm-dd") & ".md"
    Dim savePath As String
    exportStage = "choosing the Markdown destination"
    savePath = SelectedMarkdownPath(suggested)
    If Len(savePath) = 0 Then Exit Sub
    exportStage = "checking the Markdown destination"
    If FileExists(savePath) Then
        If MsgBox("Replace the existing Markdown file?" & vbCrLf & _
                  savePath, vbQuestion + vbYesNo, _
                  "Markdown export") <> vbYes Then
            Exit Sub
        End If
    End If
    exportStage = "recalculating the export data"
    CalculateExportView
    Dim markdown As String
    exportStage = "building the Markdown document"
    markdown = "# PM workbook status" & vbLf & vbLf
    markdown = markdown & "> Exported " & _
        Format$(Now, "d mmm yyyy \a\t HH:mm") & " from " & _
        MdCell(ThisWorkbook.Name) & "." & vbLf & vbLf
    markdown = markdown & "## Overview" & vbLf & vbLf
    markdown = markdown & PanelMd("Executive Status Summary", "A2:D7")
    markdown = markdown & PanelMd("Top RAID", "F2:K7")
    markdown = markdown & PanelMd("Coming up", "M2:O7")
    markdown = markdown & PanelMd("Recent progress", "Q2:U7")
    markdown = markdown & "## Source data" & vbLf & vbLf
    markdown = markdown & ItemsMd()
    markdown = markdown & RaidMd()

    exportStage = "writing the Markdown file"
    WriteUtf8 savePath, markdown
    MsgBox "Markdown exported." & vbCrLf & savePath, _
           vbInformation, "Markdown export"
    Exit Sub

exportFailure:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.description
    MsgBox "The Markdown export failed." & vbCrLf & vbCrLf & _
           "While " & exportStage & "." & vbCrLf & _
           "Error " & CStr(failureNumber) & " in " & _
           failureSource & vbCrLf & failureDescription & vbCrLf & vbCrLf & _
           "Workbook data was not changed.", _
           vbExclamation, "Markdown export"
End Sub

Private Function PanelMd(ByVal title As String, ByVal address As String) As String
    Dim sheet As Excel.Worksheet
    Set sheet = ThisWorkbook.Worksheets("Overview")
    Dim panel As Excel.Range
    Set panel = sheet.Range(address)
    If panel.Rows.Count < 2 Or panel.Columns.Count < 1 Then
        Err.Raise ERROR_BASE + 27, "PMTool.PanelMd", _
                  "Panel range " & address & " has invalid dimensions."
    End If

    Dim markdown As String
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim line As String
    Dim hasRows As Boolean
    Dim rowHasData As Boolean
    markdown = "### " & title & vbLf & vbLf & "|"
    For columnIndex = 1 To panel.Columns.Count
        markdown = markdown & " " & _
            MdCell(panel.Cells(1, columnIndex).value) & " |"
    Next columnIndex
    markdown = markdown & vbLf & "|"
    For columnIndex = 1 To panel.Columns.Count
        If IsDateHeader(CStr(panel.Cells(1, columnIndex).value)) Then
            markdown = markdown & " ---: |"
        Else
            markdown = markdown & " --- |"
        End If
    Next columnIndex
    markdown = markdown & vbLf

    For rowIndex = 2 To panel.Rows.Count
        rowHasData = False
        For columnIndex = 1 To panel.Columns.Count
            If Not IsBlankValue(panel.Cells(rowIndex, columnIndex).value) Then
                rowHasData = True
            End If
        Next columnIndex
        If rowHasData Then
            If IsBlankValue(panel.Cells(rowIndex, 1).value) Then
                Err.Raise ERROR_BASE + 107, "PMTool.PanelMd", _
                          title & " row " & rowIndex & _
                          " has data but no first-column label."
            End If
            hasRows = True
            line = "|"
            For columnIndex = 1 To panel.Columns.Count
                line = line & " " & _
                    MdCell(panel.Cells(rowIndex, columnIndex).value) & " |"
            Next columnIndex
            markdown = markdown & line & vbLf
        End If
    Next rowIndex
    If Not hasRows Then
        line = "| _No records._ |"
        For columnIndex = 2 To panel.Columns.Count
            line = line & "  |"
        Next columnIndex
        markdown = markdown & line & vbLf
    End If
    PanelMd = markdown & vbLf
End Function

Private Function ItemsMd() As String
    Dim table As Excel.ListObject
    Set table = Tbl("tblItems")
    Dim markdown As String
    markdown = "### Items" & vbLf & vbLf
    If table.DataBodyRange Is Nothing Then
        ItemsMd = markdown & "_No items recorded._" & vbLf & vbLf
        Exit Function
    End If

    markdown = markdown & _
        "| ID | Work item | Parent | Scope | Status | Priority | Owner | " & _
        "Schedule | Delivery Health | Blockers |" & vbLf
    markdown = markdown & _
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |" & vbLf

    Dim data As Variant
    data = table.DataBodyRange.value
    Dim idIndex As Long
    Dim typeIndex As Long
    Dim titleIndex As Long
    Dim parentIndex As Long
    Dim scopeIndex As Long
    Dim statusIndex As Long
    Dim priorityIndex As Long
    Dim ownerIndex As Long
    Dim startIndex As Long
    Dim dueIndex As Long
    Dim healthIndex As Long
    Dim waitingIndex As Long
    Dim blockedByIndex As Long
    Dim latestIndex As Long
    idIndex = ColOf(table, "ID")
    typeIndex = ColOf(table, "Type")
    titleIndex = ColOf(table, "Title")
    parentIndex = ColOf(table, "Parent")
    scopeIndex = ColOf(table, "Scope")
    statusIndex = ColOf(table, "Status")
    priorityIndex = ColOf(table, "Priority")
    ownerIndex = ColOf(table, "Owner")
    startIndex = ColOf(table, "Start")
    dueIndex = ColOf(table, "Due")
    healthIndex = ColOf(table, "Delivery Health")
    waitingIndex = ColOf(table, "WaitingOn")
    blockedByIndex = ColOf(table, "BlockedBy")
    latestIndex = ColOf(table, "Latest Status")
    Dim inputIndexes As Variant
    inputIndexes = Array( _
        idIndex, typeIndex, titleIndex, parentIndex, statusIndex, _
        priorityIndex, ownerIndex, startIndex, dueIndex, healthIndex, _
        blockedByIndex, latestIndex)
    Dim identifiers As New Collection
    Dim validatedCount As Long
    validatedCount = ValidateItemsHierarchy( _
        table, data, "PMTool.ItemsMd", identifiers)
    If validatedCount = 0 Then
        ItemsMd = "### Items" & vbLf & vbLf & _
                  "_No items recorded._" & vbLf & vbLf
        Exit Function
    End If

    Dim rowLines() As String
    Dim noteLines() As String
    ReDim rowLines(0 To table.ListRows.Count - 1)
    ReDim noteLines(0 To table.ListRows.Count - 1)
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim noteCount As Long
    Dim itemId As String
    Dim itemTitle As String
    Dim latestStatus As Variant
    Dim workItem As String
    Dim hasEnteredData As Boolean
    For rowIndex = 1 To table.ListRows.Count
        hasEnteredData = ArrayRowHasEnteredData( _
            data, inputIndexes, rowIndex, "Items")
        If hasEnteredData Then
            itemId = ValidatedIdentifierText( _
                data(rowIndex, idIndex), _
                "Items row " & rowIndex & " ID", "PMTool.ItemsMd")
            itemTitle = ConfiguredText( _
                data(rowIndex, titleIndex), _
                "Item '" & itemId & "' Title", "PMTool.ItemsMd")
            workItem = MdWorkItem( _
                data(rowIndex, typeIndex), itemTitle)
            rowLines(rowCount) = "| " & MdCell(itemId, False, True) & _
                " | " & workItem & " | " & _
                MdCell(data(rowIndex, parentIndex), False, True) & " | " & _
                MdCell(data(rowIndex, scopeIndex), False, True) & " | " & _
                MdCell(data(rowIndex, statusIndex), False, True) & " | " & _
                MdCell(data(rowIndex, priorityIndex), False, True) & " | " & _
                MdCell(data(rowIndex, ownerIndex), False, True) & " | " & _
                MdSchedule(data(rowIndex, startIndex), _
                           data(rowIndex, dueIndex)) & " | " & _
                MdCell(data(rowIndex, healthIndex), False, True) & " | " & _
                MdCell(data(rowIndex, waitingIndex), False, True) & " |"
            rowCount = rowCount + 1

            latestStatus = data(rowIndex, latestIndex)
            If Not IsBlankValue(latestStatus) Then
                noteLines(noteCount) = "- **" & MdCell(itemId) & " " & _
                    ChrW(&H2014) & " " & _
                    MdCell(itemTitle) & ":** " & _
                    MdCell(latestStatus)
                noteCount = noteCount + 1
            End If
        End If
    Next rowIndex
    ReDim Preserve rowLines(0 To rowCount - 1)
    markdown = markdown & Join(rowLines, vbLf) & vbLf & vbLf
    If noteCount > 0 Then
        ReDim Preserve noteLines(0 To noteCount - 1)
        markdown = markdown & "#### Latest status notes" & vbLf & vbLf & _
                   Join(noteLines, vbLf) & vbLf & vbLf
    End If
    ItemsMd = markdown
End Function

Private Function RaidMd() As String
    Dim table As Excel.ListObject
    Set table = Tbl("tblRAID")
    Dim markdown As String
    markdown = "### RAID" & vbLf & vbLf
    If table.DataBodyRange Is Nothing Then
        RaidMd = markdown & "_No RAID items recorded._" & vbLf & vbLf
        Exit Function
    End If

    markdown = markdown & _
        "| RAID ID | Type | Item | Scope | Status | Severity | Owner | Next review |" & vbLf
    markdown = markdown & _
        "| --- | --- | --- | --- | --- | --- | --- | ---: |" & vbLf
    Dim data As Variant
    data = table.DataBodyRange.value
    Dim idIndex As Long
    Dim typeIndex As Long
    Dim titleIndex As Long
    Dim scopeIndex As Long
    Dim statusIndex As Long
    Dim severityIndex As Long
    Dim ownerIndex As Long
    Dim reviewIndex As Long
    Dim detailIndex As Long
    Dim responseIndex As Long
    Dim relatedIndex As Long
    Dim probabilityIndex As Long
    Dim impactIndex As Long
    idIndex = ColOf(table, "RaidID")
    typeIndex = ColOf(table, "Type")
    titleIndex = ColOf(table, "Title")
    scopeIndex = ColOf(table, "Scope")
    statusIndex = ColOf(table, "Status")
    severityIndex = ColOf(table, "Severity")
    ownerIndex = ColOf(table, "Owner")
    reviewIndex = ColOf(table, "NextReview")
    detailIndex = ColOf(table, "Detail")
    responseIndex = ColOf(table, "Response")
    relatedIndex = ColOf(table, "RelatedID")
    probabilityIndex = ColOf(table, "Prob")
    impactIndex = ColOf(table, "Impact")
    Dim inputIndexes As Variant
    inputIndexes = Array( _
        idIndex, typeIndex, titleIndex, detailIndex, relatedIndex, _
        ownerIndex, statusIndex, probabilityIndex, impactIndex, _
        responseIndex, reviewIndex)

    Dim itemsTable As Excel.ListObject
    Dim itemData As Variant
    Dim itemIdentifiers As New Collection
    Dim validatedItemCount As Long
    Set itemsTable = Tbl("tblItems")
    If Not itemsTable.DataBodyRange Is Nothing Then
        itemData = itemsTable.DataBodyRange.value
        validatedItemCount = ValidateItemsHierarchy( _
            itemsTable, itemData, "PMTool.RaidMd", itemIdentifiers)
    End If

    Dim rowLines() As String
    Dim noteBlocks() As String
    ReDim rowLines(0 To table.ListRows.Count - 1)
    ReDim noteBlocks(0 To table.ListRows.Count - 1)
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim noteCount As Long
    Dim raidId As String
    Dim raidTitle As String
    Dim detail As Variant
    Dim response As Variant
    Dim relatedId As Variant
    Dim noteBlock As String
    Dim hasEnteredData As Boolean
    Dim identifiers As New Collection
    Dim relatedText As String
    Dim relatedRow As Long
    Dim relatedFound As Boolean
    For rowIndex = 1 To table.ListRows.Count
        hasEnteredData = ArrayRowHasEnteredData( _
            data, inputIndexes, rowIndex, "RAID")
        If hasEnteredData Then
            raidId = ValidatedIdentifierText( _
                data(rowIndex, idIndex), _
                "RAID row " & rowIndex & " ID", "PMTool.RaidMd")
            AddUniqueIdentifier identifiers, raidId, _
                                "RAID ID", "PMTool.RaidMd"
            raidTitle = ConfiguredText( _
                data(rowIndex, titleIndex), _
                "RAID item '" & raidId & "' Title", "PMTool.RaidMd")
            If Len(raidTitle) = 0 Then
                Err.Raise ERROR_BASE + 109, "PMTool.RaidMd", _
                          "RAID item '" & raidId & "' has no Title."
            End If
            relatedId = data(rowIndex, relatedIndex)
            If Not IsBlankValue(relatedId) Then
                relatedText = ValidatedIdentifierText( _
                    relatedId, "RAID item '" & raidId & "' RelatedID", _
                    "PMTool.RaidMd")
                relatedFound = False
                relatedRow = IdentifierRow( _
                    itemIdentifiers, relatedText, relatedFound, _
                    "PMTool.RaidMd")
                If Not relatedFound Or relatedRow < 1 Or _
                   validatedItemCount = 0 Then
                    Err.Raise ERROR_BASE + 125, "PMTool.RaidMd", _
                              "RAID item '" & raidId & "' RelatedID '" & _
                              relatedText & "' does not match an Items ID."
                End If
                relatedId = relatedText
            End If
            rowLines(rowCount) = "| " & MdCell(raidId, False, True) & " | " & _
                MdCell(data(rowIndex, typeIndex), False, True) & " | " & _
                MdCell(raidTitle, False, True) & " | " & _
                MdCell(data(rowIndex, scopeIndex), False, True) & " | " & _
                MdCell(data(rowIndex, statusIndex), False, True) & " | " & _
                MdCell(data(rowIndex, severityIndex), False, True) & " | " & _
                MdCell(data(rowIndex, ownerIndex), False, True) & " | " & _
                MdCell(data(rowIndex, reviewIndex), True, True) & " |"
            rowCount = rowCount + 1

            detail = data(rowIndex, detailIndex)
            response = data(rowIndex, responseIndex)
            If Not IsBlankValue(detail) Or Not IsBlankValue(response) Or _
               Not IsBlankValue(relatedId) Then
                noteBlock = "- **" & MdCell(raidId) & " " & _
                    ChrW(&H2014) & " " & _
                    MdCell(raidTitle) & "**"
                If Not IsBlankValue(detail) Then
                    noteBlock = noteBlock & vbLf & _
                                "  - Detail: " & MdCell(detail)
                End If
                If Not IsBlankValue(response) Then
                    noteBlock = noteBlock & vbLf & _
                                "  - Response: " & MdCell(response)
                End If
                If Not IsBlankValue(relatedId) Then
                    noteBlock = noteBlock & vbLf & _
                                "  - Related item: " & MdCell(relatedId)
                End If
                noteBlocks(noteCount) = noteBlock
                noteCount = noteCount + 1
            End If
        End If
    Next rowIndex
    If rowCount = 0 Then
        RaidMd = "### RAID" & vbLf & vbLf & _
                 "_No RAID items recorded._" & vbLf & vbLf
        Exit Function
    End If
    ReDim Preserve rowLines(0 To rowCount - 1)
    markdown = markdown & Join(rowLines, vbLf) & vbLf & vbLf
    If noteCount > 0 Then
        ReDim Preserve noteBlocks(0 To noteCount - 1)
        markdown = markdown & "#### RAID detail and response" & vbLf & vbLf & _
                   Join(noteBlocks, vbLf & vbLf) & vbLf & vbLf
    End If
    RaidMd = markdown
End Function

Private Function MdWorkItem(ByVal itemType As Variant, _
                            ByVal title As Variant) As String
    Dim typeText As String
    Dim titleText As String
    typeText = MdCell(itemType)
    titleText = MdCell(title)
    If Len(typeText) = 0 Then
        MdWorkItem = titleText
    ElseIf Len(titleText) = 0 Then
        MdWorkItem = typeText
    Else
        MdWorkItem = typeText & " " & ChrW(&H2014) & " " & titleText
    End If
End Function

Private Function MdSchedule(ByVal startValue As Variant, _
                            ByVal dueValue As Variant) As String
    Dim hasStart As Boolean
    Dim hasDue As Boolean
    hasStart = HasDate(startValue, "Start")
    hasDue = HasDate(dueValue, "Due")
    If hasStart And hasDue Then
        MdSchedule = MdCell(startValue, True) & " " & ChrW(&H2192) & " " & _
                     MdCell(dueValue, True)
    ElseIf hasStart Then
        MdSchedule = "From " & MdCell(startValue, True)
    ElseIf hasDue Then
        MdSchedule = "Due " & MdCell(dueValue, True)
    Else
        MdSchedule = ChrW(&H2014)
    End If
End Function

Private Function HasDate(ByVal value As Variant, _
                         ByVal fieldName As String) As Boolean
    If IsBlankValue(value) Then
        HasDate = False
        Exit Function
    End If
    If Not IsDate(value) Then
        Err.Raise ERROR_BASE + 29, "PMTool.HasDate", _
                  fieldName & " contains a non-date value."
    End If
    HasDate = True
End Function

Private Function IsDateHeader(ByVal header As String) As Boolean
    Select Case LCase$(Trim$(header))
        Case "due", "date", "completed", "next review"
            IsDateHeader = True
    End Select
End Function

Private Function EnsureMdExtension(ByVal path As String) As String
    Dim result As String
    result = path
    If Len(Trim$(result)) = 0 Then
        Err.Raise ERROR_BASE + 30, "PMTool.EnsureMdExtension", _
                  "An export path is required."
    End If
    If LCase$(Right$(result, 3)) <> ".md" Then result = result & ".md"
    EnsureMdExtension = result
End Function

Private Function SanitizeMarkdownControls(ByVal text As String) As String
    Dim position As Long
    Dim codeUnit As Long
    Dim character As String
    Dim result As String
    For position = 1 To Len(text)
        character = Mid$(text, position, 1)
        codeUnit = AscW(character)
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        If (codeUnit >= 0 And codeUnit <= 8) Or _
           (codeUnit >= 11 And codeUnit <= 12) Or _
           (codeUnit >= 14 And codeUnit <= 31) Or _
           codeUnit = 127 Then
            result = result & "[U+" & _
                Right$("0000" & Hex$(codeUnit), 4) & "]"
        Else
            result = result & character
        End If
    Next position
    SanitizeMarkdownControls = result
End Function

Private Function MdCell(ByVal value As Variant, _
                        Optional ByVal asDate As Boolean = False, _
                        Optional ByVal dashWhenBlank As Boolean = False) As String
    If IsError(value) Then
        Err.Raise ERROR_BASE + 31, "PMTool.MdCell", _
                  "A workbook error value cannot be exported."
    End If

    Dim text As String
    If IsEmpty(value) Or IsNull(value) Then
        text = ""
    ElseIf asDate Then
        If Len(Trim$(CStr(value))) = 0 Then
            text = ""
        ElseIf Not IsDate(value) Then
            Err.Raise ERROR_BASE + 32, "PMTool.MdCell", _
                      "A date column contains a non-date value."
        Else
            text = Format$(CDate(value), "d mmm yyyy")
        End If
    ElseIf VarType(value) = vbBoolean Then
        If CBool(value) Then text = "Yes" Else text = "No"
    Else
        text = CStr(value)
    End If

    text = SanitizeMarkdownControls(text)
    If Len(Trim$(text)) = 0 Then
        If dashWhenBlank Then MdCell = ChrW(&H2014) Else MdCell = ""
        Exit Function
    End If
    text = Replace(text, "\", "\\")
    text = Replace(text, "|", "\|")
    text = Replace(text, "`", "\`")
    text = Replace(text, vbTab, " ")
    text = Replace(text, vbCrLf, "<br>")
    text = Replace(text, vbLf, "<br>")
    text = Replace(text, vbCr, "<br>")
    MdCell = Trim$(text)
End Function

Private Function FileExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then
        Err.Raise ERROR_BASE + 80, "PMTool.FileExists", _
                  "A file path is required."
    End If

    On Error GoTo probeFailure
    Dim fileSize As Long
    fileSize = FileLen(path)
    FileExists = True
    Exit Function

probeFailure:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.description
    On Error GoTo 0
    Select Case failureNumber
        Case 53, 76
            FileExists = False
        Case Else
            Err.Raise failureNumber, "PMTool.FileExists", _
                      "The export path could not be inspected: " & _
                      failureDescription
    End Select
End Function

Private Sub ReadFileBytes(ByVal path As String, _
                          ByRef bytes() As Byte, _
                          ByRef byteCount As Long)
    If Not FileExists(path) Then
        Err.Raise ERROR_BASE + 35, "PMTool.ReadFileBytes", _
                  "The existing export file was not found."
    End If

    byteCount = FileLen(path)
    If byteCount > EXPORT_BYTE_MAX Then
        Err.Raise ERROR_BASE + 41, "PMTool.ReadFileBytes", _
                  "The existing Markdown file exceeds the supported size."
    End If
    If byteCount = 0 Then Exit Sub
    ReDim bytes(0 To byteCount - 1)

    Dim fileNumber As Integer
    Dim fileOpen As Boolean
    On Error GoTo readFailure
    fileNumber = FreeFile
    Open path For Binary Access Read As #fileNumber
    fileOpen = True
    Get #fileNumber, 1, bytes
    Close #fileNumber
    fileOpen = False
    Exit Sub

readFailure:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.description
    If fileOpen Then
        On Error GoTo closeFailure
        Close #fileNumber
        fileOpen = False
    End If
    On Error GoTo 0
    Err.Raise failureNumber, "PMTool.ReadFileBytes", failureDescription

closeFailure:
    failureDescription = failureDescription & _
        "; file close failed: " & Err.description
    On Error GoTo 0
    Err.Raise failureNumber, "PMTool.ReadFileBytes", failureDescription
End Sub

Private Sub WriteFileBytes(ByVal path As String, _
                           ByRef bytes() As Byte, _
                           ByVal byteCount As Long, _
                           ByRef destinationChanged As Boolean)
    If Len(path) = 0 Then
        Err.Raise ERROR_BASE + 36, "PMTool.WriteFileBytes", _
                  "A file path is required."
    End If
    If byteCount < 0 Then
        Err.Raise ERROR_BASE + 37, "PMTool.WriteFileBytes", _
                  "The byte count cannot be negative."
    End If

    Dim fileNumber As Integer
    Dim fileOpen As Boolean
    destinationChanged = False
    On Error GoTo writeFailure
    If FileExists(path) Then
        Kill path
        destinationChanged = True
    End If
    fileNumber = FreeFile
    Open path For Binary Access Write As #fileNumber
    fileOpen = True
    destinationChanged = True
    If byteCount > 0 Then Put #fileNumber, 1, bytes
    Close #fileNumber
    fileOpen = False
    If FileLen(path) <> byteCount Then
        Err.Raise ERROR_BASE + 40, "PMTool.WriteFileBytes", _
                  "The saved Markdown file has an unexpected size."
    End If
    Exit Sub

writeFailure:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.description
    If fileOpen Then
        On Error GoTo closeFailure
        Close #fileNumber
        fileOpen = False
    End If
    On Error GoTo 0
    Err.Raise failureNumber, "PMTool.WriteFileBytes", failureDescription

closeFailure:
    failureDescription = failureDescription & _
        "; file close failed: " & Err.description
    On Error GoTo 0
    Err.Raise failureNumber, "PMTool.WriteFileBytes", failureDescription
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    If Len(path) = 0 Then
        Err.Raise ERROR_BASE + 33, "PMTool.WriteUtf8", _
                  "An export path is required."
    End If
    If Len(text) = 0 Then
        Err.Raise ERROR_BASE + 34, "PMTool.WriteUtf8", _
                  "The Markdown document is empty."
    End If
    If Len(text) > EXPORT_CHARACTER_MAX Then
        Err.Raise ERROR_BASE + 52, "PMTool.WriteUtf8", _
                  "The Markdown document exceeds the supported export size."
    End If

    Dim bytes() As Byte
    Dim byteCount As Long
    Dim textIndex As Long
    Dim codeUnit As Long
    Dim lowUnit As Long
    Dim codePoint As Long
    ReDim bytes(0 To Len(text) * 4 - 1)
    textIndex = 1
    Do While textIndex <= Len(text)
        codeUnit = AscW(Mid$(text, textIndex, 1))
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        codePoint = codeUnit
        If codeUnit >= &HD800 And codeUnit <= &HDBFF Then
            If textIndex = Len(text) Then
                Err.Raise ERROR_BASE + 38, "PMTool.WriteUtf8", _
                          "The Markdown text ends with an unpaired surrogate."
            End If
            lowUnit = AscW(Mid$(text, textIndex + 1, 1))
            If lowUnit < 0 Then lowUnit = lowUnit + 65536
            If lowUnit < &HDC00 Or lowUnit > &HDFFF Then
                Err.Raise ERROR_BASE + 39, "PMTool.WriteUtf8", _
                          "The Markdown text contains an unpaired surrogate."
            End If
            codePoint = &H10000 + (codeUnit - &HD800) * &H400 + _
                        (lowUnit - &HDC00)
            textIndex = textIndex + 1
        ElseIf codeUnit >= &HDC00 And codeUnit <= &HDFFF Then
            Err.Raise ERROR_BASE + 39, "PMTool.WriteUtf8", _
                      "The Markdown text contains an unpaired surrogate."
        End If

        If codePoint < &H80 Then
            bytes(byteCount) = codePoint
            byteCount = byteCount + 1
        ElseIf codePoint < &H800 Then
            bytes(byteCount) = &HC0 Or (codePoint \ &H40)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or (codePoint And &H3F)
            byteCount = byteCount + 1
        ElseIf codePoint < &H10000 Then
            bytes(byteCount) = &HE0 Or (codePoint \ &H1000)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or ((codePoint \ &H40) And &H3F)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or (codePoint And &H3F)
            byteCount = byteCount + 1
        Else
            bytes(byteCount) = &HF0 Or (codePoint \ &H40000)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or ((codePoint \ &H1000) And &H3F)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or ((codePoint \ &H40) And &H3F)
            byteCount = byteCount + 1
            bytes(byteCount) = &H80 Or (codePoint And &H3F)
            byteCount = byteCount + 1
        End If
        textIndex = textIndex + 1
    Loop
    ReDim Preserve bytes(0 To byteCount - 1)

    Dim originalBytes() As Byte
    Dim originalByteCount As Long
    Dim hadOriginal As Boolean
    Dim originalCaptured As Boolean
    Dim destinationTouched As Boolean
    Dim restorationTouched As Boolean
    Dim restorationDescription As String
    Dim incompleteRemoved As Boolean
    On Error GoTo writeFailure
    hadOriginal = FileExists(path)
    If hadOriginal Then
        ReadFileBytes path, originalBytes, originalByteCount
        originalCaptured = True
    End If

    destinationTouched = False
    WriteFileBytes path, bytes, byteCount, destinationTouched
    Exit Sub

writeFailure:
    Dim originalNumber As Long
    Dim originalSource As String
    Dim originalDescription As String
    originalNumber = Err.Number
    originalSource = Err.Source
    originalDescription = Err.description

    If destinationTouched Then
        On Error GoTo restoreFailure
        If hadOriginal And originalCaptured Then
            WriteFileBytes path, originalBytes, originalByteCount, _
                           restorationTouched
        ElseIf FileExists(path) Then
            Kill path
        End If
    End If
    On Error GoTo 0
    Err.Raise originalNumber, originalSource, originalDescription

restoreFailure:
    restorationDescription = Err.description
    On Error GoTo removalFailure
    If FileExists(path) Then
        Kill path
        incompleteRemoved = True
    End If
    On Error GoTo 0
    originalDescription = originalDescription & _
        "; destination restoration failed: " & restorationDescription
    If incompleteRemoved Then
        originalDescription = originalDescription & _
            "; the incomplete destination was removed"
    Else
        originalDescription = originalDescription & _
            "; the destination remains absent"
    End If
    Err.Raise originalNumber, originalSource, originalDescription

removalFailure:
    originalDescription = originalDescription & _
        "; destination restoration failed: " & restorationDescription & _
        "; incomplete destination removal also failed: " & Err.description
    On Error GoTo 0
    Err.Raise originalNumber, originalSource, originalDescription
End Sub

Private Sub ClearItemRowOutline(ByVal sheet As Excel.Worksheet, _
                                ByVal firstRow As Long, _
                                ByVal lastRow As Long)
    If firstRow < 1 Or lastRow < firstRow Then
        Err.Raise ERROR_BASE + 71, "PMTool.ClearItemRowOutline", _
                  "The Items outline row range is invalid."
    End If

    Dim passCount As Long
    Dim maxLevel As Long
    Dim rowIndex As Long
    Dim currentLevel As Long
    Dim groupStart As Long
    Do
        maxLevel = 1
        For rowIndex = firstRow To lastRow
            currentLevel = sheet.Rows(rowIndex).OutlineLevel
            If currentLevel > maxLevel Then maxLevel = currentLevel
        Next rowIndex
        If maxLevel = 1 Then Exit Do

        groupStart = 0
        For rowIndex = firstRow To lastRow + 1
            If rowIndex <= lastRow Then
                currentLevel = sheet.Rows(rowIndex).OutlineLevel
            Else
                currentLevel = 0
            End If
            If currentLevel = maxLevel Then
                If groupStart = 0 Then groupStart = rowIndex
            ElseIf groupStart > 0 Then
                sheet.Rows( _
                    CStr(groupStart) & ":" & CStr(rowIndex - 1)).Ungroup
                groupStart = 0
            End If
        Next rowIndex

        passCount = passCount + 1
        If passCount > 7 Then
            Err.Raise ERROR_BASE + 72, "PMTool.ClearItemRowOutline", _
                      "The existing Items outline could not be cleared."
        End If
    Loop

    For rowIndex = firstRow To lastRow
        If sheet.Rows(rowIndex).OutlineLevel <> 1 Then
            Err.Raise ERROR_BASE + 73, "PMTool.ClearItemRowOutline", _
                      "The existing Items outline could not be verified."
        End If
    Next rowIndex
End Sub

Private Function ArrayRowHasEnteredData(ByRef data As Variant, _
                                        ByVal inputIndexes As Variant, _
                                        ByVal rowIndex As Long, _
                                        ByVal tableLabel As String) As Boolean
    If Not IsArray(data) Then
        Err.Raise ERROR_BASE + 91, "PMTool.ArrayRowHasEnteredData", _
                  "The " & tableLabel & " table values are unavailable."
    End If
    If rowIndex < LBound(data, 1) Or rowIndex > UBound(data, 1) Then
        Err.Raise ERROR_BASE + 92, "PMTool.ArrayRowHasEnteredData", _
                  tableLabel & " row " & rowIndex & _
                  " is outside the values array."
    End If
    Dim position As Long
    Dim columnIndex As Long
    Dim value As Variant
    For position = LBound(inputIndexes) To UBound(inputIndexes)
        columnIndex = CLng(inputIndexes(position))
        If columnIndex < LBound(data, 2) Or _
           columnIndex > UBound(data, 2) Then
            Err.Raise ERROR_BASE + 93, "PMTool.ArrayRowHasEnteredData", _
                      "A " & tableLabel & _
                      " input column is outside the values array."
        End If
        value = data(rowIndex, columnIndex)
        If IsError(value) Then
            Err.Raise ERROR_BASE + 81, "PMTool.ArrayRowHasEnteredData", _
                      tableLabel & " row " & rowIndex & _
                      " contains an Excel error in an editable field."
        End If
        If Not IsBlankValue(value) Then
            ArrayRowHasEnteredData = True
            Exit Function
        End If
    Next position
End Function

Private Sub AddUniqueIdentifier(ByVal identifiers As Collection, _
                                ByVal identifier As String, _
                                ByVal fieldName As String, _
                                ByVal sourceName As String)
    If identifiers Is Nothing Then
        Err.Raise ERROR_BASE + 94, sourceName, _
                  "The " & fieldName & " register is unavailable."
    End If
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo addFailure
    identifiers.Add True, "K:" & LCase$(identifier)
    On Error GoTo 0
    Exit Sub

addFailure:
    failureNumber = Err.Number
    failureDescription = Err.description
    On Error GoTo 0
    If failureNumber = 457 Then
        Err.Raise ERROR_BASE + 44, sourceName, _
                  fieldName & " '" & identifier & "' is duplicated."
    End If
    Err.Raise failureNumber, sourceName, _
              fieldName & " '" & identifier & _
              "' could not be registered: " & _
              failureDescription
End Sub

Private Sub AddIdentifierRow(ByVal identifiers As Collection, _
                             ByVal identifier As String, _
                             ByVal rowIndex As Long, _
                             ByVal sourceName As String)
    If identifiers Is Nothing Then
        Err.Raise ERROR_BASE + 110, sourceName, _
                  "The Item ID register is unavailable."
    End If
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo addFailure
    identifiers.Add rowIndex, "K:" & LCase$(identifier)
    On Error GoTo 0
    Exit Sub

addFailure:
    failureNumber = Err.Number
    failureDescription = Err.description
    On Error GoTo 0
    If failureNumber = 457 Then
        Err.Raise ERROR_BASE + 111, sourceName, _
                  "Item ID '" & identifier & "' is duplicated."
    End If
    Err.Raise failureNumber, sourceName, _
              "Item ID '" & identifier & _
              "' could not be registered: " & failureDescription
End Sub

Private Function IdentifierRow(ByVal identifiers As Collection, _
                               ByVal identifier As String, _
                               ByRef found As Boolean, _
                               ByVal sourceName As String) As Long
    If identifiers Is Nothing Then
        Err.Raise ERROR_BASE + 112, sourceName, _
                  "The Item ID register is unavailable."
    End If
    On Error GoTo lookupFailure
    IdentifierRow = CLng(identifiers.Item("K:" & LCase$(identifier)))
    found = True
    On Error GoTo 0
    Exit Function

lookupFailure:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.description
    On Error GoTo 0
    If failureNumber = 5 Then
        found = False
        Exit Function
    End If
    Err.Raise failureNumber, sourceName, _
              "Item ID '" & identifier & _
              "' could not be resolved: " & failureDescription
End Function

Private Function ValidateItemsHierarchy( _
        ByVal table As Excel.ListObject, ByRef data As Variant, _
        ByVal sourceName As String, _
        ByRef identifiers As Collection) As Long
    If table Is Nothing Then
        Err.Raise ERROR_BASE + 113, sourceName, _
                  "The Items table is unavailable."
    End If
    If table.DataBodyRange Is Nothing Then Exit Function
    If Not IsArray(data) Then
        Err.Raise ERROR_BASE + 114, sourceName, _
                  "The Items table values are unavailable."
    End If
    If identifiers Is Nothing Then Set identifiers = New Collection

    Dim idIndex As Long
    Dim typeIndex As Long
    Dim titleIndex As Long
    Dim parentIndex As Long
    Dim keyIndex As Long
    Dim levelIndex As Long
    idIndex = ColOf(table, "ID")
    typeIndex = ColOf(table, "Type")
    titleIndex = ColOf(table, "Title")
    parentIndex = ColOf(table, "Parent")
    keyIndex = ColOf(table, "WbsKey")
    levelIndex = ColOf(table, "Level")
    Dim inputIndexes As Variant
    inputIndexes = Array( _
        idIndex, typeIndex, titleIndex, parentIndex, _
        ColOf(table, "Priority"), ColOf(table, "Start"), _
        ColOf(table, "Status"), ColOf(table, "Due"), _
        ColOf(table, "Delivery Health"), _
        ColOf(table, "Latest Status"), ColOf(table, "Owner"), _
        ColOf(table, "BlockedBy"))

    Dim rowCount As Long
    rowCount = table.ListRows.Count
    Dim populated() As Boolean
    Dim itemIds() As String
    Dim parents() As String
    Dim levels() As Long
    ReDim populated(1 To rowCount)
    ReDim itemIds(1 To rowCount)
    ReDim parents(1 To rowCount)
    ReDim levels(1 To rowCount)

    Dim rowIndex As Long
    Dim itemId As String
    Dim itemTitle As String
    Dim parentId As String
    Dim keyValue As Variant
    Dim hierarchyKey As String
    Dim levelValue As Variant
    Dim level As Long
    For rowIndex = 1 To rowCount
        If ArrayRowHasEnteredData( _
                data, inputIndexes, rowIndex, "Items") Then
            populated(rowIndex) = True
            ValidateItemsHierarchy = ValidateItemsHierarchy + 1
            itemId = ValidatedIdentifierText( _
                data(rowIndex, idIndex), _
                "Items row " & rowIndex & " ID", sourceName)
            itemTitle = ConfiguredText( _
                data(rowIndex, titleIndex), _
                "Item '" & itemId & "' Title", sourceName)
            If Len(itemTitle) = 0 Then
                Err.Raise ERROR_BASE + 115, sourceName, _
                          "Item '" & itemId & "' has no Title."
            End If

            keyValue = data(rowIndex, keyIndex)
            If IsError(keyValue) Then
                Err.Raise ERROR_BASE + 116, sourceName, _
                          "Item '" & itemId & _
                          "' has no valid hierarchy key."
            End If
            hierarchyKey = ConfiguredText( _
                keyValue, "Item '" & itemId & "' WbsKey", sourceName)
            If Len(hierarchyKey) = 0 Then
                Err.Raise ERROR_BASE + 116, sourceName, _
                          "Item '" & itemId & _
                          "' has no valid hierarchy key."
            End If

            levelValue = data(rowIndex, levelIndex)
            If IsError(levelValue) Then
                Err.Raise ERROR_BASE + 117, sourceName, _
                          "Item '" & itemId & "' has an invalid level."
            End If
            If Not IsNumeric(levelValue) Then
                Err.Raise ERROR_BASE + 117, sourceName, _
                          "Item '" & itemId & "' has an invalid level."
            End If
            If CDbl(levelValue) <> Fix(CDbl(levelValue)) Then
                Err.Raise ERROR_BASE + 118, sourceName, _
                          "Item '" & itemId & "' has a fractional level."
            End If
            level = CLng(levelValue)
            If level < 1 Or level > 6 Then
                Err.Raise ERROR_BASE + 119, sourceName, _
                          "Item '" & itemId & "' has level " & level & _
                          "; expected 1 to 6."
            End If

            parentId = ConfiguredText( _
                data(rowIndex, parentIndex), _
                "Item '" & itemId & "' Parent", sourceName)
            If Len(parentId) > 0 Then
                parentId = ValidatedIdentifierText( _
                    data(rowIndex, parentIndex), _
                    "Item '" & itemId & "' Parent", sourceName)
            End If
            If level = 1 And Len(parentId) > 0 Then
                Err.Raise ERROR_BASE + 120, sourceName, _
                          "Level-1 item '" & itemId & _
                          "' cannot have a Parent."
            End If
            If level > 1 And Len(parentId) = 0 Then
                Err.Raise ERROR_BASE + 121, sourceName, _
                          "Item '" & itemId & "' requires a Parent."
            End If

            itemIds(rowIndex) = itemId
            parents(rowIndex) = parentId
            levels(rowIndex) = level
            AddIdentifierRow identifiers, itemId, rowIndex, sourceName
        End If
    Next rowIndex

    Dim parentRow As Long
    Dim chainRow As Long
    Dim depth As Long
    Dim found As Boolean
    Dim reachedRoot As Boolean
    For rowIndex = 1 To rowCount
        If populated(rowIndex) And levels(rowIndex) > 1 Then
            found = False
            parentRow = IdentifierRow( _
                identifiers, parents(rowIndex), found, sourceName)
            If Not found Then
                Err.Raise ERROR_BASE + 122, sourceName, _
                          "Item '" & itemIds(rowIndex) & "' Parent '" & _
                          parents(rowIndex) & "' does not exist."
            End If
            If levels(parentRow) >= levels(rowIndex) Then
                Err.Raise ERROR_BASE + 123, sourceName, _
                          "Item '" & itemIds(rowIndex) & "' at level " & _
                          levels(rowIndex) & " must reference a Parent " & _
                          "at a lower level."
            End If

            chainRow = rowIndex
            reachedRoot = False
            For depth = 1 To 6
                If levels(chainRow) = 1 Then
                    reachedRoot = True
                    Exit For
                End If
                found = False
                parentRow = IdentifierRow( _
                    identifiers, parents(chainRow), found, sourceName)
                If Not found Then Exit For
                If levels(parentRow) >= levels(chainRow) Then Exit For
                chainRow = parentRow
            Next depth
            If Not reachedRoot Then
                Err.Raise ERROR_BASE + 124, sourceName, _
                          "Item '" & itemIds(rowIndex) & _
                          "' does not resolve through valid Parents " & _
                          "to a Level-1 item."
            End If
        End If
    Next rowIndex
End Function

Private Sub AppendDiagnostic(ByRef diagnostics As String, _
                             ByVal label As String, _
                             ByVal description As String)
    If Len(diagnostics) > 0 Then diagnostics = diagnostics & "; "
    diagnostics = diagnostics & label & ": " & description
End Sub

Private Function TryRestoreEvents(ByVal previousValue As Boolean, _
                                  ByRef failureDescription As String) As Boolean
    On Error GoTo restoreFailure
    Application.EnableEvents = previousValue
    TryRestoreEvents = True
    Exit Function

restoreFailure:
    failureDescription = Err.description
End Function

Private Function TryRestoreScreenUpdating( _
        ByVal previousValue As Boolean, _
        ByRef failureDescription As String) As Boolean
    On Error GoTo restoreFailure
    Application.ScreenUpdating = previousValue
    TryRestoreScreenUpdating = True
    Exit Function

restoreFailure:
    failureDescription = Err.description
End Function

Private Function RestoreOrganiseApplicationState( _
        ByVal previousEvents As Boolean, _
        ByVal previousScreenUpdating As Boolean, _
        ByRef diagnostics As String) As Boolean
    Dim screenDescription As String
    Dim eventsDescription As String
    Dim screenRestored As Boolean
    Dim eventsRestored As Boolean
    screenRestored = TryRestoreScreenUpdating( _
        previousScreenUpdating, screenDescription)
    eventsRestored = TryRestoreEvents(previousEvents, eventsDescription)
    If Not screenRestored Then
        AppendDiagnostic diagnostics, _
                         "ScreenUpdating restoration failed", _
                         screenDescription
    End If
    If Not eventsRestored Then
        AppendDiagnostic diagnostics, _
                         "EnableEvents restoration failed", _
                         eventsDescription
    End If
    RestoreOrganiseApplicationState = screenRestored And eventsRestored
End Function

Private Function TryClearItemRowOutline( _
        ByVal sheet As Excel.Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long, _
        ByRef failureDescription As String) As Boolean
    On Error GoTo clearFailure
    ClearItemRowOutline sheet, firstRow, lastRow
    TryClearItemRowOutline = True
    Exit Function

clearFailure:
    failureDescription = Err.description
End Function

Private Function TryResetItemFormatting( _
        ByVal table As Excel.ListObject, _
        ByVal sheet As Excel.Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal titleIndex As Long, _
        ByRef failureDescription As String) As Boolean
    On Error GoTo resetFailure
    sheet.Rows(CStr(firstRow) & ":" & CStr(lastRow)).RowHeight = 24
    If Not table.DataBodyRange Is Nothing Then
        With table.ListColumns(titleIndex).DataBodyRange
            .IndentLevel = 0
            .Font.Bold = False
            .Font.Size = 10
        End With
    End If
    TryResetItemFormatting = True
    Exit Function

resetFailure:
    failureDescription = Err.description
End Function

Private Function FlattenItemsPresentation( _
        ByVal table As Excel.ListObject, _
        ByVal sheet As Excel.Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal titleIndex As Long, _
        ByRef diagnostics As String) As Boolean
    Dim outlineDescription As String
    Dim formatDescription As String
    Dim outlineCleared As Boolean
    Dim formatReset As Boolean
    outlineCleared = TryClearItemRowOutline( _
        sheet, firstRow, lastRow, outlineDescription)
    formatReset = TryResetItemFormatting( _
        table, sheet, firstRow, lastRow, titleIndex, formatDescription)
    If Not outlineCleared Then
        AppendDiagnostic diagnostics, _
                         "flat outline restoration failed", _
                         outlineDescription
    End If
    If Not formatReset Then
        AppendDiagnostic diagnostics, _
                         "flat row formatting restoration failed", _
                         formatDescription
    End If
    FlattenItemsPresentation = outlineCleared And formatReset
End Function

Public Sub OrganiseItems()
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim applicationStateChanged As Boolean
    Dim mutationStarted As Boolean
    Dim organiseStage As String
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo organiseFailure
    organiseStage = "opening the Items table"
    Dim table As Excel.ListObject
    Set table = Tbl("tblItems")
    If table.DataBodyRange Is Nothing Then
        MsgBox "There are no item rows to organise.", _
               vbInformation, "Organise rows"
        Exit Sub
    End If
    If table.ListRows.Count > ITEM_CAPACITY Then
        Err.Raise ERROR_BASE + 42, "PMTool.OrganiseItems", _
                  "Items contains " & table.ListRows.Count & _
                  " rows; the supported capacity is " & ITEM_CAPACITY & "."
    End If

    Dim idIndex As Long
    Dim keyIndex As Long
    Dim levelIndex As Long
    Dim titleIndex As Long
    idIndex = ColOf(table, "ID")
    keyIndex = ColOf(table, "WbsKey")
    levelIndex = ColOf(table, "Level")
    titleIndex = ColOf(table, "Title")
    Dim inputIndexes As Variant
    inputIndexes = Array( _
        idIndex, ColOf(table, "Type"), titleIndex, _
        ColOf(table, "Parent"), ColOf(table, "Priority"), _
        ColOf(table, "Start"), ColOf(table, "Status"), _
        ColOf(table, "Due"), ColOf(table, "Delivery Health"), _
        ColOf(table, "Latest Status"), ColOf(table, "Owner"), _
        ColOf(table, "BlockedBy"))

    Application.EnableEvents = False
    applicationStateChanged = True
    Application.ScreenUpdating = False

    organiseStage = "calculating the Items table"
    table.Range.Calculate

    organiseStage = "validating the Items hierarchy"
    Dim data As Variant
    data = table.DataBodyRange.Value2
    Dim identifiers As New Collection
    Dim rowIndex As Long
    Dim activeCount As Long
    Dim idValue As Variant
    Dim titleValue As Variant
    Dim hasEnteredData As Boolean
    Dim level As Long
    activeCount = ValidateItemsHierarchy( _
        table, data, "PMTool.OrganiseItems", identifiers)

    If activeCount = 0 Then
        Dim emptyRestoreDiagnostics As String
        If Not RestoreOrganiseApplicationState( _
                previousEvents, previousScreenUpdating, _
                emptyRestoreDiagnostics) Then
            Err.Raise ERROR_BASE + 126, "PMTool.OrganiseItems", _
                      emptyRestoreDiagnostics
        End If
        applicationStateChanged = False
        MsgBox "There are no populated item rows to organise.", _
               vbInformation, "Organise rows"
        Exit Sub
    End If

    Dim sheet As Excel.Worksheet
    Set sheet = table.Parent
    Dim outlineFirstRow As Long
    Dim outlineLastRow As Long
    outlineFirstRow = table.DataBodyRange.Row
    outlineLastRow = outlineFirstRow + table.DataBodyRange.Rows.Count - 1

    organiseStage = "clearing the existing row groups"
    mutationStarted = True
    ClearItemRowOutline sheet, outlineFirstRow, outlineLastRow

    organiseStage = "sorting the Items hierarchy"
    With table.Sort
        .SortFields.Clear
        .SortFields.Add Key:=table.ListColumns(keyIndex).DataBodyRange, _
                        Order:=xlAscending
        .header = xlYes
        .Apply
    End With

    organiseStage = "verifying the sorted Items rows"
    data = table.DataBodyRange.Value2
    For rowIndex = 1 To table.ListRows.Count
        hasEnteredData = ArrayRowHasEnteredData( _
            data, inputIndexes, rowIndex, "Items")
        If rowIndex <= activeCount And Not hasEnteredData Then
            Err.Raise ERROR_BASE + 95, "PMTool.OrganiseItems", _
                      "The Items sort left a blank row inside the populated range."
        End If
        If rowIndex > activeCount And hasEnteredData Then
            Err.Raise ERROR_BASE + 96, "PMTool.OrganiseItems", _
                      "The Items sort left entered data below the populated range."
        End If
    Next rowIndex

    organiseStage = "resizing the Items table"
    Dim headerRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    headerRow = table.HeaderRowRange.Row
    firstColumn = table.Range.Column
    lastColumn = firstColumn + table.ListColumns.Count - 1
    table.Resize sheet.Range( _
        sheet.Cells(headerRow, firstColumn), _
        sheet.Cells(headerRow + activeCount, lastColumn))
    If table.ListRows.Count <> activeCount Then
        Err.Raise ERROR_BASE + 97, "PMTool.OrganiseItems", _
                  "The Items table did not resize to the populated row count."
    End If
    data = table.DataBodyRange.Value2

    Dim groupLevel As Long
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim firstDataRow As Long
    firstDataRow = table.DataBodyRange.Row
    organiseStage = "rebuilding the row groups"
    For groupLevel = 2 To 6
        groupStart = 0
        For rowIndex = 1 To table.ListRows.Count + 1
            If rowIndex <= table.ListRows.Count Then
                idValue = data(rowIndex, idIndex)
                titleValue = data(rowIndex, titleIndex)
                If Len(Trim$(CStr(idValue))) = 0 And _
                   Len(Trim$(CStr(titleValue))) = 0 Then
                    level = 0
                Else
                    level = CLng(data(rowIndex, levelIndex))
                End If
            Else
                level = 0
            End If

            If level >= groupLevel Then
                If groupStart = 0 Then groupStart = firstDataRow + rowIndex - 1
            ElseIf groupStart > 0 Then
                groupEnd = firstDataRow + rowIndex - 2
                sheet.Rows( _
                    CStr(groupStart) & ":" & CStr(groupEnd)).Group
                groupStart = 0
            End If
        Next rowIndex
    Next groupLevel

    Dim tableRow As Excel.Range
    Dim titleSheetColumn As Long
    organiseStage = "formatting the hierarchy"
    titleSheetColumn = table.Range.Column + titleIndex - 1
    For rowIndex = 1 To table.ListRows.Count
        Set tableRow = table.DataBodyRange.Rows(rowIndex)
        idValue = data(rowIndex, idIndex)
        titleValue = data(rowIndex, titleIndex)
        If Len(Trim$(CStr(idValue))) = 0 And _
           Len(Trim$(CStr(titleValue))) = 0 Then GoTo nextFormatRow
        level = CLng(data(rowIndex, levelIndex))
        If sheet.Rows(tableRow.Row).OutlineLevel <> level Then
            Err.Raise ERROR_BASE + 54, "PMTool.OrganiseItems", _
                      "The outline level for item '" & _
                      CStr(tableRow.Cells(1, idIndex).value) & _
                      "' could not be verified."
        End If
        ApplyItemLevelPresentation sheet, tableRow.Row, _
                                   titleSheetColumn, level
nextFormatRow:
    Next rowIndex

    organiseStage = "restoring Excel state"
    Dim successRestoreDiagnostics As String
    If Not RestoreOrganiseApplicationState( _
            previousEvents, previousScreenUpdating, _
            successRestoreDiagnostics) Then
        Err.Raise ERROR_BASE + 127, "PMTool.OrganiseItems", _
                  successRestoreDiagnostics
    End If
    applicationStateChanged = False
    MsgBox CStr(activeCount) & " item rows organised.", _
           vbInformation, "Organise rows"
    Exit Sub

organiseFailure:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.description
    Dim cleanupDiagnostics As String
    Dim cleanupSummary As String
    Dim flattenSucceeded As Boolean
    If mutationStarted Then
        flattenSucceeded = FlattenItemsPresentation( _
            table, sheet, outlineFirstRow, outlineLastRow, _
            titleIndex, cleanupDiagnostics)
        If flattenSucceeded Then
            cleanupSummary = _
                "Items was left in a verified flat layout."
        End If
    End If
    If applicationStateChanged Then
        Dim stateDiagnostics As String
        If Not RestoreOrganiseApplicationState( _
                previousEvents, previousScreenUpdating, _
                stateDiagnostics) Then
            AppendDiagnostic cleanupDiagnostics, _
                             "Excel state restoration", _
                             stateDiagnostics
        End If
    End If
    MsgBox "Rows could not be organised." & vbCrLf & vbCrLf & _
           "While " & organiseStage & "." & vbCrLf & _
           "Error " & CStr(failureNumber) & " in " & _
           failureSource & vbCrLf & failureDescription & _
           IIf(Len(cleanupSummary) > 0, _
               vbCrLf & vbCrLf & cleanupSummary, "") & _
           IIf(Len(cleanupDiagnostics) > 0, _
               vbCrLf & vbCrLf & cleanupDiagnostics, ""), _
           IIf(Len(cleanupDiagnostics) > 0, vbCritical, vbExclamation), _
           "Organise rows"
End Sub
