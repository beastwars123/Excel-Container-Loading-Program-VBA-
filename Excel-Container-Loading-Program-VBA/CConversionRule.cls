VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConversionRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 单位转换规则类，用于封装单个转换规则的所有属性
' 必须与 InitializeConverter 中读取的数据源列一一对应

'#Region "类属性" - 转换规则的核心字段
Public fromUnit As String
' 源单位（如 "M"、"KG"），从数据源 FromUnit 列读取，统一转为大写并去空格
' 注意：VBA 区分大小写，需与 ConvertUnit 传入的单位格式完全一致（通常全大写）

Public toUnit As String
' 目标单位（如 "MM"、"G"），从数据源 ToUnit 列读取，统一转为大写并去空格

Public Factor As Double
' 转换因子（线性转换公式中的乘法因子）：
' - 对于 LINEAR 类型：目标值 = 原值 × Factor
' - 对于 INVERSE 类型：目标值 = Factor ÷ 原值
' 数据源中对应 Factor 列，默认值为 1（通过 NzVBA 处理空值）

Public Offset As Double
' 转换偏移量（仅 LINEARWITHOFFSET 类型使用）：
' 目标值 = 原值 × Factor + Offset
' 数据源中对应 Offset 列，默认值为 0（通过 NzVBA 处理空值）

Public FormulaType As String
' 转换公式类型（必须为以下三种之一）：
' - "LINEAR"：简单线性转换（目标值 = 原值 × Factor）
' - "LINEARWITHOFFSET"：带偏移的线性转换（目标值 = 原值 × Factor + Offset）
' - "INVERSE"：反向转换（目标值 = Factor ÷ 原值，原值不能为 0）
' 数据源中对应 FormulaType 列，默认值为 "LINEAR"（通过 UCase 统一为大写）

Public ValidationHash As String
' 校验哈希值，用于跟踪规则来源（生成格式："R行号_时间戳"）
' 主要用于调试和数据验证，确保规则与数据源行号对应
'#End Region


'#Region "类初始化" - 确保属性初始值正确
Private Sub Class_Initialize()
    ' 初始化时设置默认值（防止未赋值时出现空值错误）
    Factor = 1
    Offset = 0
    FormulaType = "LINEAR"
End Sub

'#End Region


'#Region "类终止" - 释放资源（可选，简单类可省略）
Private Sub Class_Terminate()
    ' 清理资源（如有对象引用，需在此释放）
    ' 本类无复杂对象，可留空
End Sub
'#End Region
