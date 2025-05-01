VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CContainerLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CContainerLoad 类模块：存储集装箱装载结果
Option Explicit

Public Container As CContainer     ' 集装箱实例
Public LoadedBoxes As Collection   ' 已装载的货物实例（CBox对象）
Public UsedHeight As Double        ' 已使用的总高度（cm）
Public UsedWeight As Double        ' 已使用的总重量（kg）

Private Sub Class_Initialize()
    ' 初始化集合
    Set LoadedBoxes = New Collection
End Sub

