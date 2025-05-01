VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CContainer 类模块代码
Option Explicit

Private m_Name As String
Private m_InnerLength As Double
Private m_InnerWidth As Double
Private m_InnerHeight As Double
Private m_MaxLoad As Double

' 集装箱类型属性
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal strValue As String)
    m_Name = strValue
End Property

' 内部长度属性
Public Property Get InnerLength() As Double
    InnerLength = m_InnerLength
End Property

Public Property Let InnerLength(ByVal dblValue As Double)
    m_InnerLength = dblValue
End Property

' 内部宽度属性
Public Property Get InnerWidth() As Double
    InnerWidth = m_InnerWidth
End Property

Public Property Let InnerWidth(ByVal dblValue As Double)
    m_InnerWidth = dblValue
End Property

' 内部高度属性
Public Property Get InnerHeight() As Double
    InnerHeight = m_InnerHeight
End Property

Public Property Let InnerHeight(ByVal dblValue As Double)
    m_InnerHeight = dblValue
End Property

' 最大载重属性
Public Property Get MaxLoad() As Double
    MaxLoad = m_MaxLoad
End Property

Public Property Let MaxLoad(ByVal dblValue As Double)
    m_MaxLoad = dblValue
End Property

' 体积属性
Public Property Get volume() As Double
    volume = m_InnerLength * m_InnerWidth * m_InnerHeight
End Property

' 显示名称属性
Public Property Get DisplayName() As String
    DisplayName = m_Name & " (" & Format(m_InnerLength, "0") & "x" & _
                Format(m_InnerWidth, "0") & "x" & _
                Format(m_InnerHeight, "0") & ")"
End Property
