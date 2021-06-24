Attribute VB_Name = "tblFilter"
Sub tbFilter()
'-------------------------------------------------------------------
' filter table and create new sheet from pivot table
'-------------------------------------------------------------------
    Dim pt As PivotTable
    Dim pv As PivotField
    Dim epaId As PivotItem
    Dim idString As String
    
    Set pt = Worksheets("Summary").PivotTables(1)
    Set pv = pt.PivotFields("TSDF ID")
    
    For Each epaId In pv.PivotItems
        If epaId <> "" Then
            Worksheets("MTN suggestions").ListObjects("Table1").Range.AutoFilter _
            Field:=1, Criteria1:="=" & epaId
        Else
            MsgBox ("Value not found where expected")
            Exit Sub
        End If
        addSheet (epaId)
        idString = epaId
        Worksheets("MTN suggestions").ListObjects("Table1").Range.Copy _
        Destination:=Worksheets(idString).Range("A1")
        newBook (idString)
    Next epaId
    
End Sub

Sub addSheet(name)
'-------------------------------------------------------------------
' Add Sheet and name by input
'-------------------------------------------------------------------
    Sheets.Add(After:=Sheets(Sheets.Count)).name = name
End Sub

Sub newBook(sheetName)
    Dim wb As Workbook
    Dim epaId As String
    
    'epaId = Worksheets("Summary").Range("A4")

    ActiveWorkbook.Sheets(sheetName).Copy
    ActiveWorkbook.SaveAs "C:\Users\dgraha01\OneDrive - Environmental Protection Agency (EPA)\e-Manifest - dg\DQ_issues\InvalidGenID\" & sheetName & ".csv", FileFormat:=xlCSV
    ActiveWorkbook.Close

End Sub

Sub filterToSheet()
'-------------------------------------------------------------------
' filter table and create new sheet from pivot table
'-------------------------------------------------------------------
    Dim pt As PivotTable
    Dim pv As PivotField
    Dim epaId As PivotItem
    Dim idString As String
    
    Set pt = Worksheets("Summary").PivotTables(1)
    Set pv = pt.PivotFields("TSDF ID")
    
    For Each epaId In pv.PivotItems
        If epaId <> "" Then
            Worksheets("MTN suggestions").ListObjects("Table1").Range.AutoFilter _
            Field:=1, Criteria1:="=" & epaId
        Else
            MsgBox ("Value not found where expected")
            Exit Sub
        End If
        'addSheet (epaId)
        idString = epaId
        Worksheets("MTN suggestions").ListObjects("Table1").Range.Copy _
        Destination:=Worksheets(idString).Range("A1")
    Next epaId
    
End Sub
