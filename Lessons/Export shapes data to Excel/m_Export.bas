Attribute VB_Name = "m_Export"
Option Explicit

Dim i As Integer


Public Sub ExportShapeDatasToExcel()

Dim ex As Excel.Application
Dim exDoc As Excel.Workbook
Dim exSheet As Excel.Worksheet

    Set ex = New Excel.Application
    Set exDoc = ex.Workbooks.Add
    Set exSheet = exDoc.Worksheets(1)
    
    i = 1
    
Dim shp As Visio.Shape
    
    For Each shp In Visio.Application.ActivePage.Shapes
        FillShpData exSheet, shp, 1
    Next shp
    
    ex.Visible = True
    
End Sub


Private Sub FillShpData(ByRef exSheet As Excel.Worksheet, ByRef shp As Visio.Shape, ByVal colNum As Integer)
'    exSheet.Range("A" & i).Value = shp.Name
    exSheet.Cells(i, colNum).Value = shp.Name
    i = i + 1
    
    If shp.Shapes.Count > 0 Then
        For Each shp In shp.Shapes
            FillShpData exSheet, shp, colNum + 1
        Next shp
    End If
    
    
End Sub
