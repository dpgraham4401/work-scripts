Attribute VB_Name = "WklyUpdate"
Sub ppc2pivot()
Attribute ppc2pivot.VB_ProcData.VB_Invoke_Func = " \n14"
'-------------------------------------------------------------------
' ppc2pivot2 Macro
'-------------------------------------------------------------------
    Cells.Select
    Range("H11").Activate
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "PPC-search-export!R1C1:R1048576C6", Version:=6).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=6
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TSDF ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Manifest Tracking Number"), _
        "Count of Manifest Tracking Number", xlCount
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TSDF ID").AutoSort _
        xlDescending, "Count of Manifest Tracking Number"
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = "EPA ID"
    Range("1:1,2:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp
    Sheets("Sheet1").Select
    Sheets("Sheet1").name = "Week"
    Sheets("PPC-search-export").Select
    ActiveWindow.SelectedSheets.Delete
End Sub
