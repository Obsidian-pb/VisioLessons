VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub DrawFromEcel()
Dim exApp As Excel.Application
Dim i As Integer
Dim exSheet As Excel.Worksheet

    Set exApp = New Excel.Application
    exApp.Visible = False
    
    exApp.Workbooks.Open ("c:\Coord.xls")
    Set exSheet = exApp.ActiveWorkbook.Worksheets(1)
    
    i = 2
    Do While exSheet.Range("A" & i).Value <> ""
        Application.ActivePage.DrawLine exSheet.Range("A" & i).Value, exSheet.Range("B" & i).Value, _
                                        exSheet.Range("C" & i).Value, exSheet.Range("D" & i).Value
        
        i = i + 1
        
    Loop
    
    exApp.Quit
End Sub
