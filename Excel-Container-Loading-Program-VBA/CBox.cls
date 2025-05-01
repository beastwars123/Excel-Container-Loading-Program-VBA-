VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CBox 类模块代码
Option Explicit

Private m_ID As String
Private m_Length As Double
Private m_Width As Double
Private m_Height As Double
Private m_Weight As Double
Private m_Quantity As Long
Private m_Stackable As Boolean
Private m_Rotatable As Boolean
Private m_RotationAxes As String
Private m_Color As Long
Private m_Shape As String
Private m_Fragility As String
Private m_CenterOfGravityX As Double
Private m_CenterOfGravityY As Double
Private m_CenterOfGravityZ As Double
Private m_SpecialHandling As String
Private m_Grouping As String
Private m_Precedence As Long
Private m_VolumeDensity As Double
Private m_MaxStackLayers As Long
Private m_WeightCapacity As Double
Private m_CanInvert As Boolean


' ID 属性
Public Property Get ID() As String
    ID = m_ID
End Property
Public Property Let ID(ByVal strValue As String)
    m_ID = strValue
End Property

' 长度属性
Public Property Get length() As Double
    length = m_Length
End Property
Public Property Let length(ByVal dblValue As Double)
    m_Length = dblValue
End Property

' 宽度属性
Public Property Get width() As Double
    width = m_Width
End Property
Public Property Let width(ByVal dblValue As Double)
    m_Width = dblValue
End Property

' 高度属性
Public Property Get height() As Double
    height = m_Height
End Property
Public Property Let height(ByVal dblValue As Double)
    m_Height = dblValue
End Property

' 重量属性
Public Property Get weight() As Double
    weight = m_Weight
End Property
Public Property Let weight(ByVal dblValue As Double)
    m_Weight = dblValue
End Property

' 数量属性
Public Property Get Quantity() As Long
    Quantity = m_Quantity
End Property
Public Property Let Quantity(ByVal lngValue As Long)
    m_Quantity = lngValue
End Property

' 可堆叠属性
Public Property Get Stackable() As Boolean
    Stackable = m_Stackable
End Property
Public Property Let Stackable(ByVal blnValue As Boolean)
    m_Stackable = blnValue
End Property

' 可旋转属性
Public Property Get Rotatable() As Boolean
    Rotatable = m_Rotatable
End Property
Public Property Let Rotatable(ByVal blnValue As Boolean)
    m_Rotatable = blnValue
End Property

' 旋转轴属性
Public Property Get rotationAxes() As String
    rotationAxes = m_RotationAxes
End Property
Public Property Let rotationAxes(ByVal strValue As String)
    m_RotationAxes = strValue
End Property

' 颜色属性
Public Property Get Color() As Long
    Color = m_Color
End Property
Public Property Let Color(ByVal lngValue As Long)
    m_Color = lngValue
End Property

' 形状属性
Public Property Get Shape() As String
    Shape = m_Shape
End Property
Public Property Let Shape(ByVal strValue As String)
    m_Shape = strValue
End Property

' 脆弱性属性
Public Property Get Fragility() As String
    Fragility = m_Fragility
End Property
Public Property Let Fragility(ByVal strValue As String)
    m_Fragility = strValue
End Property

' 重心位置 X 坐标属性
Public Property Get CenterOfGravityX() As Double
    CenterOfGravityX = m_CenterOfGravityX
End Property
Public Property Let CenterOfGravityX(ByVal dblValue As Double)
    m_CenterOfGravityX = dblValue
End Property

' 重心位置 Y 坐标属性
Public Property Get CenterOfGravityY() As Double
    CenterOfGravityY = m_CenterOfGravityY
End Property
Public Property Let CenterOfGravityY(ByVal dblValue As Double)
    m_CenterOfGravityY = dblValue
End Property

' 重心位置 Z 坐标属性
Public Property Get CenterOfGravityZ() As Double
    CenterOfGravityZ = m_CenterOfGravityZ
End Property
Public Property Let CenterOfGravityZ(ByVal dblValue As Double)
    m_CenterOfGravityZ = dblValue
End Property

' 特殊处理要求属性
Public Property Get SpecialHandling() As String
    SpecialHandling = m_SpecialHandling
End Property
Public Property Let SpecialHandling(ByVal strValue As String)
    m_SpecialHandling = strValue
End Property

' 分组信息属性
Public Property Get Grouping() As String
    Grouping = m_Grouping
End Property
Public Property Let Grouping(ByVal strValue As String)
    m_Grouping = strValue
End Property

' 装箱优先级属性
Public Property Get Precedence() As Long
    Precedence = m_Precedence
End Property
Public Property Let Precedence(ByVal lngValue As Long)
    m_Precedence = lngValue
End Property

' 体积密度属性
Public Property Get VolumeDensity() As Double
    VolumeDensity = m_VolumeDensity
End Property
Public Property Let VolumeDensity(ByVal dblValue As Double)
    m_VolumeDensity = dblValue
End Property

' 获取尺寸字符串函数
Public Function GetDimensions() As String
    GetDimensions = m_Length & "x" & m_Width & "x" & m_Height
End Function

' 获取体积函数
Public Function GetVolume() As Double
    GetVolume = m_Length * m_Width * m_Height
End Function

' 最大叠层数属性
Public Property Get MaxStackLayers() As Long
    MaxStackLayers = m_MaxStackLayers
End Property
Public Property Let MaxStackLayers(ByVal lngValue As Long)
    m_MaxStackLayers = lngValue
End Property

' 承重能力属性
Public Property Get WeightCapacity() As Double
    WeightCapacity = m_WeightCapacity
End Property
Public Property Let WeightCapacity(ByVal dblValue As Double)
    m_WeightCapacity = dblValue
End Property

' 允许倒置属性
Public Property Get CanInvert() As Boolean
    CanInvert = m_CanInvert
End Property
Public Property Let CanInvert(ByVal blnValue As Boolean)
    m_CanInvert = blnValue
End Property

