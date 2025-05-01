VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 模块级常量和枚举定义
Private Const UNIT_METRIC As String = "Metric"
Private Const UNIT_IMPERIAL As String = "Imperial"
Private Const LENGTH_UNIT_METRIC As String = "cm"
Private Const LENGTH_UNIT_IMPERIAL As String = "in"
Private Const WEIGHT_UNIT_METRIC As String = "kg"
Private Const WEIGHT_UNIT_IMPERIAL As String = "lb"

Private Enum DimensionType
    length
    weight
End Enum

' 模块级变量声明
Private previousUnitSystem As String

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If target.Address <> "$H$4" Then
        Application.EnableEvents = True
        Exit Sub
    End If
    previousUnitSystem = Me.Range("H4").value
End Sub

' Worksheet change event handler
Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call UnitConverter.InitializeConverter
    
    Dim unitCell As Range
    Set unitCell = Me.Range("H4")
    Dim containerCells As Range
    Set containerCells = Me.ListObjects("CTNR_Use").ListColumns("CTNR").DataBodyRange
    
    If Not Intersect(target, Union(unitCell, containerCells)) Is Nothing Then
        Dim isMetric As Boolean
        isMetric = (Me.Range("H4").value = UNIT_METRIC)
        
        If target.Address = "$H$4" Then
            If Me.Range("H4").value = previousUnitSystem Then
                Range("I4").Select
                GoTo ExitHandler
            End If
            ' 仅在单位变更时处理 Cargo_Spec
            ConvertData Me.ListObjects("Cargo_Spec"), isMetric, False
            Range("I4").Select
        End If
        
        ' 处理 CTNR_Use
        ConvertContainerData Me.ListObjects("CTNR_Use"), isMetric
        
        ApplyNumberFormatting isMetric
    End If

ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume ExitHandler
End Sub

' 处理 CTNR_Use 数据转换
Private Sub ConvertContainerData(ByVal table As ListObject, ByVal isMetric As Boolean)
    Dim containerSpec As Range
    Set containerSpec = Me.Range("CTNR_Spec")
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
        Dim containerName As String
        containerName = table.DataBodyRange.Cells(i, 1).value
        
        ' 处理 CTNR 删除情况
        If containerName = "" Then
            table.ListColumns("Length").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Width").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Height").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Payload").DataBodyRange.Cells(i).value = ""
            GoTo NextRow
        End If
        
        Dim specRow As Range
        Set specRow = containerSpec.Find(containerName, LookAt:=xlWhole, MatchCase:=False)
        
        If Not specRow Is Nothing Then
            ConvertDimension table.ListColumns("Length").DataBodyRange.Cells(i), specRow.Cells(1, 2), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Width").DataBodyRange.Cells(i), specRow.Cells(1, 3), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Height").DataBodyRange.Cells(i), specRow.Cells(1, 4), isMetric, DimensionType.length, True
            ConvertDimension table.ListColumns("Payload").DataBodyRange.Cells(i), specRow.Cells(1, 5), isMetric, DimensionType.weight, True
        Else
            table.ListColumns("Length").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Width").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Height").DataBodyRange.Cells(i).value = ""
            table.ListColumns("Payload").DataBodyRange.Cells(i).value = ""
        End If
NextRow:
    Next i
End Sub

' 通用的数据转换函数
Private Sub ConvertData(ByVal table As ListObject, ByVal isMetric As Boolean, ByVal isFromSpec As Boolean)
    Dim lengthCols As Variant
    Dim weightCols As Variant
    Dim cell As Range
    
    If table.Name = "Cargo_Spec" Then
        lengthCols = Array("Length", "Width", "Height")
        weightCols = Array("Weight")
    End If
    
    ' 处理长度列
    For Each colName In lengthCols
        For Each cell In table.ListColumns(colName).DataBodyRange
            If cell.value <> "" And IsNumeric(cell.value) Then
                ConvertDimension cell, cell, isMetric, DimensionType.length, isFromSpec
            End If
        Next cell
    Next colName
    
    ' 处理重量列
    For Each colName In weightCols
        For Each cell In table.ListColumns(colName).DataBodyRange
            If cell.value <> "" And IsNumeric(cell.value) Then
                ConvertDimension cell, cell, isMetric, DimensionType.weight, isFromSpec
            End If
        Next cell
    Next colName
End Sub

' Apply number formatting based on unit system
Private Sub ApplyNumberFormatting(ByVal isMetric As Boolean)
    Dim lengthFormat As String
    Dim weightFormat As String
    
    lengthFormat = IIf(isMetric, "0.00 ""cm""", "0.00 ""in""")
    weightFormat = IIf(isMetric, "0.00 ""kg""", "0.00 ""lb""")
    
    ApplyFormatting Me.ListObjects("CTNR_Use"), Array("Length", "Width", "Height"), lengthFormat
    ApplyFormatting Me.ListObjects("CTNR_Use"), Array("Payload"), weightFormat
    ApplyFormatting Me.ListObjects("Cargo_Spec"), Array("Length", "Width", "Height"), lengthFormat
    ApplyFormatting Me.ListObjects("Cargo_Spec"), Array("Weight"), weightFormat
End Sub

' 通用的格式设置函数
Private Sub ApplyFormatting(ByVal table As ListObject, ByVal colNames As Variant, ByVal formatStr As String)
    Dim colName As Variant
    For Each colName In colNames
        With table.ListColumns(colName).DataBodyRange
            .NumberFormat = formatStr
            .HorizontalAlignment = xlRight
        End With
    Next colName
End Sub

' Convert individual dimension with error handling
Private Sub ConvertDimension(ByVal targetCell As Range, ByVal sourceCell As Range, ByVal isMetric As Boolean, ByVal dimType As DimensionType, ByVal isFromSpec As Boolean)
    On Error GoTo ErrorHandler
    If sourceCell Is Nothing Or IsEmpty(sourceCell) Or Not IsNumeric(sourceCell.value) Then
        ' 如果不需要将无数据处设为 #N/A，可将以下行注释掉
        ' targetCell.Value = CVErr(xlErrNA)
        Exit Sub
    End If
    
    Dim originalValue As Double
    originalValue = CDbl(sourceCell.value)
    
    Dim fromUnit As String
    Dim toUnit As String
    
    Select Case dimType
        Case DimensionType.length
            fromUnit = IIf(isFromSpec, LENGTH_UNIT_METRIC, IIf(isMetric, LENGTH_UNIT_IMPERIAL, LENGTH_UNIT_METRIC))
        Case DimensionType.weight
            fromUnit = IIf(isFromSpec, WEIGHT_UNIT_METRIC, IIf(isMetric, WEIGHT_UNIT_IMPERIAL, WEIGHT_UNIT_METRIC))
    End Select
    
    toUnit = IIf(isMetric, IIf(dimType = length, LENGTH_UNIT_METRIC, WEIGHT_UNIT_METRIC), IIf(dimType = length, LENGTH_UNIT_IMPERIAL, WEIGHT_UNIT_IMPERIAL))
    
    targetCell.value = UnitConverter.ConvertUnit(originalValue, fromUnit, toUnit)
    Exit Sub

ErrorHandler:
    targetCell.value = CVErr(xlErrValue)
    Debug.Print "Dimension conversion error: " & Err.Description
    Resume Next
End Sub

' 自定义错误处理函数
Private Sub HandleError(ByVal errorNumber As Long, ByVal errorDescription As String)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Operation failed: " & errorDescription & vbCrLf & _
           "Error code: " & errorNumber, vbCritical
End Sub

