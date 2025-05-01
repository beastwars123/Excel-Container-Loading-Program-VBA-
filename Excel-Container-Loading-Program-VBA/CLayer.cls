VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLayer 类模块：表示集装箱内的一个层（二维平面）
Option Explicit

Public z As Double                ' 层底部高度（从集装箱底部开始，单位：cm）
Public availableAreas As Collection ' 该层的可用区域集合，每个区域为数组(x, y, width, height)

Private Sub Class_Initialize()
    ' 初始化时创建可用区域集合
    Set availableAreas = New Collection
End Sub

' 添加可用区域（封装方法，避免直接操作集合）
Public Sub AddAvailableArea(x As Double, y As Double, width As Double, height As Double)
    availableAreas.Add Array(x, y, width, height)
End Sub

