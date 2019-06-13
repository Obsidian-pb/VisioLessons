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

Const mmInInch = 25.4

Public Sub GetDimensions()
Dim shp As Visio.Shape
    
    Set shp = Application.ActiveWindow.Selection(1)
    
    Debug.Print "Square: " & Round(shp.AreaIU * mmInInch ^ 2, 0)
    Debug.Print "Length: " & Round(shp.LengthIU * mmInInch, 0)
    
    
End Sub
