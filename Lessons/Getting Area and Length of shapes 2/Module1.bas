Attribute VB_Name = "Module1"
Option Explicit

Const mmInInch = 25.4

Public Sub GetDimensions(ByRef shp As Visio.Shape)
    
'    Debug.Print "Square: " & Round(shp.AreaIU * mmInInch ^ 2, 0)
'    Debug.Print "Length: " & Round(shp.LengthIU * mmInInch, 0)
    
    shp.Cells("Prop.Square").Formula = Round(shp.AreaIU * mmInInch ^ 2, 0)
    shp.Cells("Prop.Perimeter").Formula = Round(shp.LengthIU * mmInInch, 0)
    
End Sub
