VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPackingResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CPackingResult 类模块
Private pContainer As CContainer
Private pPackedBoxes As Collection
Private pRemainingSpaces As Collection
Private pWeightUsage As Double
Private pEfficiency As Double

' 属性定义
Public Property Set Container(value As CContainer)
    Set pContainer = value
End Property
Public Property Get Container() As CContainer
    Set Container = pContainer
End Property

Public Property Set PackedBoxes(value As Collection)
    Set pPackedBoxes = value
End Property
Public Property Get PackedBoxes() As Collection
    Set PackedBoxes = pPackedBoxes
End Property

Public Property Set RemainingSpaces(value As Collection)
    Set pRemainingSpaces = value
End Property
Public Property Get RemainingSpaces() As Collection
    Set RemainingSpaces = pRemainingSpaces
End Property

Public Property Let WeightUsage(value As Double)
    pWeightUsage = value
End Property
Public Property Get WeightUsage() As Double
    WeightUsage = pWeightUsage
End Property

Public Property Let Efficiency(value As Double)
    pEfficiency = value
End Property
Public Property Get Efficiency() As Double
    Efficiency = pEfficiency
End Property

' 初始化方法
Private Sub Class_Initialize()
    Set pPackedBoxes = New Collection
    Set pRemainingSpaces = New Collection
End Sub
