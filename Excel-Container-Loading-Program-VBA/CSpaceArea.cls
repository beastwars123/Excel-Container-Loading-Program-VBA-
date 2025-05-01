VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSpaceArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CSpaceArea ÀàÄ£¿é
Public x As Double
Public y As Double
Public z As Double
Public length As Double
Public width As Double
Public height As Double

Public Sub Initialize(Optional ByVal x_ As Double = 0, _
                     Optional ByVal y_ As Double = 0, _
                     Optional ByVal z_ As Double = 0, _
                     Optional ByVal length_ As Double = 0, _
                     Optional ByVal width_ As Double = 0, _
                     Optional ByVal height_ As Double = 0)
    x = x_
    y = y_
    z = z_
    length = length_
    width = width_
    height = height_
End Sub

Public Property Get volume() As Double
    volume = length * width * height
End Property

