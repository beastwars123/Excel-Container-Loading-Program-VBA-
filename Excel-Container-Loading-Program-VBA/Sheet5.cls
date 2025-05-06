VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
End
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ==================== 模块级定义 ====================
' 单位系统常量
Private Const UNIT_METRIC As String = "Metric"    ' 公制单位标识
Private Const UNIT_IMPERIAL As String = "Imperial" ' 英制单位标识

' 单位符号常量
Private Const LENGTH_UNIT_METRIC As String = "cm" ' 公制长度单位
Private Const LENGTH_UNIT_IMPERIAL As String = "in" ' 英制长度单位
Private Const WEIGHT_UNIT_METRIC As String = "kg" ' 公制重量单位
Private Const WEIGHT_UNIT_IMPERIAL As String = "lb" ' 英制重量单位

' 维度类型枚举
Private Enum DimensionType
    length    ' 长度维度
    weight    ' 重量维度
End Enum

' 记录前一个单位系统的状态
Private previousUnitSystem As String

' ==================== 工作表事件处理 ====================
' 处理单元格选择变化事件
' @param target 用户当前选择的单元格范围
Private Sub Worksheet_SelectionChange(ByVal target As Range)
    ' 仅在选择H4单元格时记录当前单位系统
    If target.Address <> "$H$4" Then
        Application.EnableEvents = True
        Exit Sub
    End If
    previousUnitSystem = Me.Range("H4").value
End Sub

' 主工作表变更事件处理程序
' @param target 发生变更的单元格范围
Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo ErrorHandler
    ' 禁用事件和屏幕刷新以提升性能
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call UnitConverter.InitializeConverter  ' 初始化单位转换器

    ' 定义关键单元格范围
    Dim unitCell As Range
    Set unitCell = Me.Range("H4")  ' 单位系统选择单元格
    Dim containerCells As Range
    Set containerCells = Me.ListObjects("CTNR_Use").ListColumns("CTNR").DataBodyRange ' 集装箱使用数据范围
    
    ' 当变更涉及单位系统或集装箱数据时
    If Not Intersect(target, Union(unitCell, containerCells)) Is Nothing Then
        Dim isMetric As Boolean
        isMetric = (Me.Range("H4").value = UNIT_METRIC)  ' 判断当前单位系统

        ' 处理单位系统变更
        If target.Address = "$H$4" Then
            ' 单位未实际变化时直接退出
            If Me.Range("H4").value = previousUnitSystem Then
                Range("I4").Select
                GoTo ExitHandler
            End If
            ' 转换货物规格数据
            ConvertData Me.ListObjects("Cargo_Spec"), isMetric, False
            Range("I4").Select
        End If

        ' 转换集装箱使用数据
        ConvertContainerData Me.ListObjects("CTNR_Use"), isMetric

        ' 应用新的数字格式
        ApplyNumberFormatting isMetric
    End If

ExitHandler:
    ' 恢复系统设置
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume ExitHandler
End Sub

' ==================== 数据转换核心方法 ====================
' 转换集装箱使用表数据
' @param table 要处理的ListObject表格对象
' @param isMetric 当前是否为公制单位系统
Private Sub ConvertContainerData(ByVal table As ListObject, ByVal isMetric As Boolean)
    Dim containerSpec As Range
    Set containerSpec = Me.Range("CTNR_Spec")  ' 获取集装箱规格数据源
    
    ' 遍历表格每一行
    Dim i As Long
    For i = 1 To table.ListRows.Count
        Dim containerName As String
        containerName = table.DataBodyRange.Cells(i, 1).value

        ' 处理空容器名称的情况
        If containerName = "" Then
            ClearContainerDimensions table, i
            GoTo NextRow
        End If

        ' 查找对应的规格数据
        Dim specRow As Range
        Set specRow = containerSpec.Find(containerName, LookAt:=xlWhole, MatchCase:=False)
        
        If Not specRow Is Nothing Then
            ' 转换各个维度数据
            ConvertDimension table.ListColumns("Length").DataBodyRange.Cells(i), specRow.Cells(1, 2), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Width").DataBodyRange.Cells(i), specRow.Cells(1, 3), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Height").DataBodyRange.Cells(i), specRow.Cells(1, 4), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Payload").DataBodyRange.Cells(i), specRow.Cells(1, 5), isMetric, DimensionType.weight, True
        Else
            ClearContainerDimensions table, i
        End If
NextRow:
    Next i
End Sub

' ==================== 通用功能方法 ====================
' 转换指定表格的数据
' @param table 要转换的ListObject表格
' @param isMetric 当前单位系统是否为公制
' @param isFromSpec 标记数据是否来自规格表（影响转换方向）
Private Sub ConvertData(ByVal table As ListObject, ByVal isMetric As Boolean, ByVal isFromSpec As Boolean)
    ' 定义需要转换的列
    Dim lengthCols As Variant
    Dim weightCols As Variant

    If table.Name = "Cargo_Spec" Then
        lengthCols = Array("Length", "Width", "Height")  ' 需要转换的长度相关列
        weightCols = Array("Weight")                     ' 需要转换的重量相关列
    End If

    ' 处理长度列转换
    Dim colName As Variant
    For Each colName In lengthCols
        Dim cell As Range
        For Each cell In table.ListColumns(colName).DataBodyRange
            If cell.value <> "" And IsNumeric(cell.value) Then
                ConvertDimension cell, cell, isMetric, DimensionType.length, isFromSpec
            End If
        Next cell
    Next colName

    ' 处理重量列转换
    For Each colName In weightCols
        For Each cell In table.ListColumns(colName).DataBodyRange
            If cell.value <> "" And IsNumeric(cell.value) Then
                ConvertDimension cell, cell, isMetric, DimensionType.weight, isFromSpec
            End If
        Next cell
    Next colName
End Sub

' ==================== 格式设置方法 ====================
' 应用数字格式到指定表格
' @param isMetric 当前是否为公制单位系统
Private Sub ApplyNumberFormatting(ByVal isMetric As Boolean)
    ' 定义格式字符串
    Dim lengthFormat As String
    Dim weightFormat As String

    lengthFormat = IIf(isMetric, "0.00 ""cm""", "0.00 ""in""")  ' 长度数值格式
    weightFormat = IIf(isMetric, "0.00 ""kg""", "0.00 ""lb""")  ' 重量数值格式

    ' 应用格式到不同表格
    ApplyFormatting Me.ListObjects("CTNR_Use"), Array("Length", "Width", "Height"), lengthFormat
    ApplyFormatting Me.ListObjects("CTNR_Use"), Array("Payload"), weightFormat
    ApplyFormatting Me.ListObjects("Cargo_Spec"), Array("Length", "Width", "Height"), lengthFormat
    ApplyFormatting Me.ListObjects("Cargo_Spec"), Array("Weight"), weightFormat
End Sub

' ==================== 工具方法 ====================
' 清空指定行的集装箱维度数据
' @param table 目标表格对象
' @param rowIndex 要清空的行索引
Private Sub ClearContainerDimensions(ByVal table As ListObject, ByVal rowIndex As Long)
    With table
        .ListColumns("Length").DataBodyRange.Cells(rowIndex).value = ""
        .ListColumns("Width").DataBodyRange.Cells(rowIndex).value = ""
        .ListColumns("Height").DataBodyRange.Cells(rowIndex).value = ""
        .ListColumns("Payload").DataBodyRange.Cells(rowIndex).value = ""
    End With
End Sub

' 应用格式到指定表格列
' @param table 目标表格对象
' @param colNames 要设置格式的列名称数组
' @param formatStr 要应用的格式字符串
Private Sub ApplyFormatting(ByVal table As ListObject, ByVal colNames As Variant, ByVal formatStr As String)
    Dim colName As Variant
    For Each colName In colNames
        With table.ListColumns(colName).DataBodyRange
            .NumberFormat = formatStr      ' 设置数字格式
            .HorizontalAlignment = xlRight ' 右对齐
        End With
    Next colName
End Sub

' ==================== 核心转换逻辑 ====================
' 执行单位维度转换
' @param targetCell 要写入转换结果的单元格
' @param sourceCell 包含原始值的单元格
' @param isMetric 目标单位是否为公制
' @param dimType 维度类型（长度/重量）
' @param isFromSpec 标记转换方向：True表示从规格数据转换，False表示用户输入数据转换
Private Sub ConvertDimension(ByVal targetCell As Range,
                             ByVal sourceCell As Range,
                             ByVal isMetric As Boolean,
                             ByVal dimType As DimensionType,
                             ByVal isFromSpec As Boolean)
    On Error GoTo ErrorHandler
    ' 处理空值或非数值输入
    If sourceCell Is Nothing Or IsEmpty(sourceCell) Or Not IsNumeric(sourceCell.value) Then
        Exit Sub
    End If

    Dim originalValue As Double
    originalValue = CDbl(sourceCell.value)  ' 获取原始数值

    ' 确定转换方向
    Dim fromUnit As String
    Dim toUnit As String

    Select Case dimType
        Case DimensionType.length
            ' 长度单位转换逻辑
            fromUnit = IIf(isFromSpec, LENGTH_UNIT_METRIC, IIf(isMetric, LENGTH_UNIT_IMPERIAL, LENGTH_UNIT_METRIC))
        Case DimensionType.weight
            ' 重量单位转换逻辑
            fromUnit = IIf(isFromSpec, WEIGHT_UNIT_METRIC, IIf(isMetric, WEIGHT_UNIT_IMPERIAL, WEIGHT_UNIT_METRIC))
    End Select

    toUnit = IIf(isMetric, IIf(dimType = length, LENGTH_UNIT_METRIC, WEIGHT_UNIT_METRIC), IIf(dimType = length, LENGTH_UNIT_IMPERIAL, WEIGHT_UNIT_IMPERIAL))

    ' 执行单位转换并写入结果
    targetCell.value = UnitConverter.ConvertUnit(originalValue, fromUnit, toUnit)
    Exit Sub

ErrorHandler:
    targetCell.value = CVErr(xlErrValue)  ' 标记错误值
    Debug.Print "Dimension conversion error: " & Err.Description
    Resume Next
End Sub

' ==================== 错误处理 ====================
' 统一错误处理程序
' @param errorNumber 错误代码
' @param errorDescription 错误描述
Private Sub HandleError(ByVal errorNumber As Long, ByVal errorDescription As String)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Operation failed: " & errorDescription & vbCrLf &
           "Error code: " & errorNumber, vbCritical
End Sub