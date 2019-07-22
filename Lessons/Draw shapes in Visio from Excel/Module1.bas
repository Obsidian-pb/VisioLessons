Attribute VB_Name = "Module1"
Option Explicit

Public Sub DrawInVisio()
Dim VisioApp As Visio.Application
Dim visDoc As Visio.Document
Dim pg As Visio.Page
Dim shp As Visio.Shape

    Set VisioApp = New Visio.Application
    
    VisioApp.Visible = True
    
    
    Set visDoc = VisioApp.Documents.AddEx("", visMSMetric)
    
    
    Dim i As Integer
    For i = 2 To 5
        Set shp = visDoc.Pages(1).DrawRectangle(Range("B" & i), Range("C" & i), Range("D" & i), Range("E" & i))
        shp.Characters.Text = i
    Next i
    
End Sub
