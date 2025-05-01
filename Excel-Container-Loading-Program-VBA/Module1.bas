Attribute VB_Name = "Module1"
Option Explicit

' 模块顶部定义表格名常量（关键！）
Private Const TABLE_CTNR_USE As String = "CTNR_Use"
Private Const TABLE_CARGO_SPEC As String = "Cargo_Spec"

Public Function GetSelectedContainers(ws As Worksheet) As Collection
    Dim containers As New Collection
    Dim row As ListRow
    Dim tbl As ListObject
    
    ' 检查表格是否存在
    If Not DoesTableExist(ws, TABLE_CTNR_USE) Then
        Debug.Print "表格 " & TABLE_CTNR_USE & " 不存在于工作表 " & ws.Name
        Exit Function
    End If
    
    Set tbl = ws.ListObjects(TABLE_CTNR_USE)
    Dim ctnrCol As ListColumn
    Set ctnrCol = tbl.ListColumns("CTNR")
    
    Dim c As CContainer
    For Each row In tbl.ListRows
        Set c = New CContainer
        Dim ctnrValue As Variant
        ctnrValue = ctnrCol.DataBodyRange.Cells(row.Index, 1).value
        ' 检查 CTNR 值是否为空或错误值
        If Not IsEmpty(ctnrValue) And Not IsError(ctnrValue) And Trim(ctnrValue) <> "" Then
            c.Name = CStr(ctnrValue)
            c.InnerLength = tbl.ListColumns("Length").DataBodyRange.Cells(row.Index, 1).value
            c.InnerWidth = tbl.ListColumns("Width").DataBodyRange.Cells(row.Index, 1).value
            c.InnerHeight = tbl.ListColumns("Height").DataBodyRange.Cells(row.Index, 1).value
            c.MaxLoad = tbl.ListColumns("Payload").DataBodyRange.Cells(row.Index, 1).value
            containers.Add c
        End If
    Next row
    
    Set GetSelectedContainers = containers
End Function

' 自动填充Color列（30到50之间的随机颜色索引）
Private Sub FillRandomColor(row As ListRow, columnIndex As Long)
    Randomize ' 初始化随机种子
    Dim randomColor As Long
    randomColor = Int((50 - 34 + 1) * Rnd + 34) ' 生成34到50的随机整数
    row.Range.Cells(1, columnIndex).Interior.ColorIndex = randomColor
End Sub

' 自动填充 Cargo_Spec 表格的 Color 列和 VolumeDensity 列
Public Sub AutoFillCargoSpec(ws As Worksheet)
    Dim cargoTbl As ListObject
    Dim unitCell As Range
    Dim unitSystem As String ' "Metric" 或 "Imperial"
    
    ' 检查表格和单位单元格是否存在
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Then Exit Sub
    Set cargoTbl = ws.ListObjects(TABLE_CARGO_SPEC)
    Dim colorColIndex As Long
    colorColIndex = cargoTbl.ListColumns("Color").Index ' 获取Color列的索引
    
    Set unitCell = ws.Range("H4") ' 假设单位单元格为 H4，可根据实际调整
    unitSystem = UCase(Trim(Coalesce(unitCell, "METRIC")))
    
    Dim row As ListRow
    For Each row In cargoTbl.ListRows
        Dim cargoNameCell As Range, lengthCell As Range
        Set cargoNameCell = row.Range.Cells(1, 1) ' CargoName列
        Set lengthCell = row.Range.Cells(1, 2) ' Length列
        
        If Not IsEmpty(cargoNameCell.value) And Not IsEmpty(lengthCell.value) Then
            ' 填充随机颜色（30到50的ColorIndex）
            FillRandomColor row, colorColIndex
            
            ' 计算VolumeDensity（保留原有逻辑）
            Dim length As Double, width As Double, height As Double, weight As Double
            length = GetNumericValue(lengthCell, 0)
            width = GetNumericValue(row.Range.Cells(1, 3), 0) ' Width列
            height = GetNumericValue(row.Range.Cells(1, 4), 0) ' Height列
            weight = GetNumericValue(row.Range.Cells(1, 5), 0) ' Weight列
            
            Dim volume As Double
            Select Case unitSystem
                Case "METRIC"
                    volume = (length * width * height) / 1000000 ' cm3转m3
                Case "IMPERIAL"
                    volume = (length * width * height) / 1728 ' in3转ft3
            End Select
            
            If volume > 0 Then
                row.Range.Cells(1, cargoTbl.ListColumns("VolumeDensity").Index).value = weight / volume
            Else
                row.Range.Cells(1, cargoTbl.ListColumns("VolumeDensity").Index).value = 0
            End If
        End If
    Next row
End Sub


' 读取货物数据（处理新字段的空值）
Public Function ReadBoxDataFromSheet(ws As Worksheet) As Collection
    Dim boxes As New Collection
    Dim row As ListRow
    Dim cargoTbl As ListObject
    
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Then
        Debug.Print "表格 " & TABLE_CARGO_SPEC & " 不存在于工作表 " & ws.Name
        Exit Function
    End If
    
    Set cargoTbl = ws.ListObjects(TABLE_CARGO_SPEC)
    Dim cargoNameCol As ListColumn
    Set cargoNameCol = cargoTbl.ListColumns("CargoName")
    
    ' 检查 cargoNameCol 的 DataBodyRange 是否为空
    If cargoNameCol.DataBodyRange Is Nothing Then
        Debug.Print "CargoName 列的数据体范围为空"
        Exit Function
    End If
    
    For Each row In cargoTbl.ListRows
        Dim currentBox As New CBox
        Dim cargoNameValue As Variant
        ' 获取单元格的值并处理空值
        cargoNameValue = Coalesce(cargoNameCol.DataBodyRange.Cells(row.Index, 1), "")
        
        If Trim(cargoNameValue) = "" Then
            Set currentBox = Nothing
            GoTo SkipCurrentRow ' 跳过货物名称为空的行
        End If
        
        With currentBox
            .ID = cargoNameValue
            .length = GetNumericValue(cargoTbl.ListColumns("Length").DataBodyRange.Cells(row.Index, 1), 0#)
            .width = GetNumericValue(cargoTbl.ListColumns("Width").DataBodyRange.Cells(row.Index, 1), 0#)
            .height = GetNumericValue(cargoTbl.ListColumns("Height").DataBodyRange.Cells(row.Index, 1), 0#)
            .weight = GetNumericValue(cargoTbl.ListColumns("Weight").DataBodyRange.Cells(row.Index, 1), 0#)
            .Quantity = CLng(Coalesce(cargoTbl.ListColumns("Quantity").DataBodyRange.Cells(row.Index, 1), "0"))
            .Stackable = IsYes(cargoTbl.ListColumns("Stackable").DataBodyRange.Cells(row.Index, 1))
            .Rotatable = IsYes(cargoTbl.ListColumns("Rotatable").DataBodyRange.Cells(row.Index, 1))
            .rotationAxes = UCase(Trim(Coalesce(cargoTbl.ListColumns("RotationAxes").DataBodyRange.Cells(row.Index, 1), "XYZ")))
            
            ' 处理新增字段的空值（字符串类型默认空字符串，数值类型默认0）
            .Shape = Trim(Coalesce(cargoTbl.ListColumns("Shape").DataBodyRange.Cells(row.Index, 1), ""))
            .Fragility = Trim(Coalesce(cargoTbl.ListColumns("Fragility").DataBodyRange.Cells(row.Index, 1), ""))
            
            .MaxStackLayers = CLng(Coalesce(cargoTbl.ListColumns("MaxStackLayers").DataBodyRange.Cells(row.Index, 1), "1"))
            .WeightCapacity = GetNumericValue(cargoTbl.ListColumns("WeightCapacity").DataBodyRange.Cells(row.Index, 1), 0#)
            .CanInvert = IsYes(cargoTbl.ListColumns("CanInvert").DataBodyRange.Cells(row.Index, 1)) ' 使用现有IsYes函数处理布尔值


            ' 重心位置：空值默认0，确保在0~1之间（实际业务可增加范围检查）
            .CenterOfGravityX = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityX").DataBodyRange.Cells(row.Index, 1), 0#)
            .CenterOfGravityY = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityY").DataBodyRange.Cells(row.Index, 1), 0#)
            .CenterOfGravityZ = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityZ").DataBodyRange.Cells(row.Index, 1), 0#)
            
            .SpecialHandling = Trim(Coalesce(cargoTbl.ListColumns("SpecialHandling").DataBodyRange.Cells(row.Index, 1), ""))
            .Grouping = Trim(Coalesce(cargoTbl.ListColumns("Grouping").DataBodyRange.Cells(row.Index, 1), ""))
            
            ' 优先级：空值默认0（0表示无优先级，数值越小优先级越高）
            .Precedence = CLng(Coalesce(cargoTbl.ListColumns("Precedence").DataBodyRange.Cells(row.Index, 1), "0"))
            
            ' 体积密度：空值默认0，避免除以0错误（需确保体积不为0，后续逻辑处理）
            .VolumeDensity = GetNumericValue(cargoTbl.ListColumns("VolumeDensity").DataBodyRange.Cells(row.Index, 1), 0#)
            
            ' 颜色字段（保留原有逻辑）
            Dim colorCol As ListColumn
            Set colorCol = cargoTbl.ListColumns("Color")
            On Error Resume Next
            .Color = colorCol.DataBodyRange.Cells(row.Index, 1).Interior.Color
            On Error GoTo 0
        End With
        
        ' 生成多个实例（处理Quantity为0或非整数的情况）
        Dim instanceCount As Long, i As Long
        instanceCount = Max(1, currentBox.Quantity)
        For i = 1 To instanceCount
            Dim newBox As New CBox
            CopyBoxProperties currentBox, newBox
            boxes.Add newBox
            Set newBox = Nothing
        Next i
        
SkipCurrentRow:
        Set currentBox = Nothing
    Next row
    
    Set ReadBoxDataFromSheet = boxes
End Function


' 辅助函数：处理Yes/No转换（增强健壮性）
Private Function IsYes(value As Variant) As Boolean
    value = Coalesce(value, "No")
    IsYes = LCase(Trim(value)) = "yes"
End Function
' 辅助函数：安全获取数值（支持默认值）
Private Function GetNumericValue(cell As Range, Optional defaultValue As Double = 0#) As Double
    If cell Is Nothing Then
        GetNumericValue = defaultValue
        Exit Function
    End If
    Dim cellValue As Variant
    cellValue = cell.value
    If IsEmpty(cellValue) Or IsError(cellValue) Then
        GetNumericValue = defaultValue
    ElseIf IsNumeric(cellValue) Then
        GetNumericValue = CDbl(cellValue)
    Else
        GetNumericValue = defaultValue
    End If
End Function

' 辅助函数：处理空值或错误值（返回字符串）
Private Function Coalesce(cell As Variant, defaultValue As String) As String
    If cell Is Nothing Then
        Coalesce = defaultValue
        Exit Function
    End If
    Dim cellValue As Variant
    cellValue = cell.value
    If IsEmpty(cellValue) Or IsError(cellValue) Then
        Coalesce = defaultValue
    Else
        Coalesce = CStr(cellValue)
    End If
End Function



' 初始化结果输出区域
Public Sub InitResultArea(ws As Worksheet)
    With ws.Range("U3")
        If .CurrentRegion.Count > 1 Then .CurrentRegion.Clear ' 安全清空旧数据
        .Resize(1, 8).value = Array("集装箱类型", "已用数量", "剩余空间(L)", "剩余空间(W)", "剩余空间(H)", "载重使用率", "装箱策略", "货物列表")
        .EntireColumn.AutoFit
    End With
End Sub

' 辅助函数：检查表格是否存在
Private Function DoesTableExist(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    Dim dummy As ListObject: Set dummy = ws.ListObjects(tableName)
    DoesTableExist = Not dummy Is Nothing
End Function




' 辅助函数：复制CBox属性（深拷贝）
Private Sub CopyBoxProperties(source As CBox, target As CBox)
    With source
        target.ID = .ID
        target.length = .length
        target.width = .width
        target.height = .height
        target.weight = .weight
        target.Quantity = 1 ' 每个实例数量固定为1，总数量由外层循环控制
        target.Stackable = .Stackable
        target.Rotatable = .Rotatable
        target.rotationAxes = .rotationAxes
        target.Color = .Color
        target.MaxStackLayers = .MaxStackLayers
        target.WeightCapacity = .WeightCapacity
        target.CanInvert = .CanInvert

    End With
End Sub

' 辅助函数：获取非负整数（处理Quantity可能的错误值）
Private Function Max(ByVal a As Long, ByVal b As Long) As Long
    Max = IIf(a >= b, a, b)
End Function



Sub TestDataLoading()
    Dim ws As Worksheet
    Set ws = Worksheets("Stuffing") ' 确保工作表名称正确
    
    Dim containers As Collection
    Dim boxes As Collection
    
    ' 自动填充Color和VolumeDensity列
    AutoFillCargoSpec ws
    ' 初始化结果区域
    InitResultArea ws
    
    ' 获取用户选择的集装箱
    Set containers = GetSelectedContainers(ws)
    ' 获取货物数据
    Set boxes = ReadBoxDataFromSheet(ws)
    
    ' ------------------------
    ' 输出集装箱详细信息
    ' ------------------------
    Debug.Print "===== 集装箱数据列表 ====="
    If containers.Count = 0 Then
        Debug.Print "未获取到有效集装箱数据"
    Else
        Dim c As CContainer ' 声明循环变量 c 为 CContainer 类型
        For Each c In containers
            Debug.Print "? 集装箱类型: " & c.Name
            Debug.Print "   内部尺寸 (LxWxH): " & c.InnerLength & " x " & c.InnerWidth & " x " & c.InnerHeight
            Debug.Print "   最大载重: " & c.MaxLoad & " kg"
            Debug.Print "------------------------"
        Next c
    End If
    
    ' ------------------------
    ' 输出货物详细信息（包含新增属性）
    ' ------------------------
    Debug.Print vbNewLine & "===== 货物数据列表 ====="
    If boxes.Count = 0 Then
        Debug.Print "未获取到有效货物数据"
    Else
        Dim box As CBox ' 声明循环变量 box 为 CBox 类型
        For Each box In boxes
            Debug.Print "? 货物ID: " & box.ID
            Debug.Print "   尺寸 (LxWxH): " & box.length & " x " & box.width & " x " & box.height
            Debug.Print "   重量: " & box.weight & " kg"
            Debug.Print "   数量: " & box.Quantity & " 件"
            Debug.Print "   可堆叠: " & IIf(box.Stackable, "是", "否")
            Debug.Print "   可旋转: " & IIf(box.Rotatable, "是", "否")
            Debug.Print "   旋转轴: " & box.rotationAxes
            Debug.Print "   颜色 (RGB): " & box.Color
            ' 新增属性打印
            Debug.Print "   可堆叠层数: " & box.MaxStackLayers
            Debug.Print "   承重能力: " & box.WeightCapacity & " kg"
            Debug.Print "   允许倒置: " & IIf(box.CanInvert, "是", "否")

            Debug.Print "   形状: " & box.Shape
            Debug.Print "   脆弱性: " & box.Fragility
            Debug.Print "   重心位置 (X,Y,Z): " & box.CenterOfGravityX & ", " & box.CenterOfGravityY & ", " & box.CenterOfGravityZ
            Debug.Print "   特殊处理: " & box.SpecialHandling
            Debug.Print "   分组信息: " & box.Grouping
            Debug.Print "   装箱优先级: " & box.Precedence
            Debug.Print "   体积密度: " & box.VolumeDensity & " kg/m3" ' 假设单位为 kg/m3，可根据实际调整
            Debug.Print "------------------------"
        Next box
    End If
    
    ' 显示统计结果
    Debug.Print vbNewLine & "===== 统计结果 ====="
    Debug.Print "已选择集装箱种类: " & containers.Count
    Debug.Print "待装货物总数: " & boxes.Count & " 件" ' 注意：boxes.Count是实例总数（Quantity展开后）
    
    ' 示例输出到U3（保留原有表格初始化逻辑）
    With ws.Range("U3")
        .Offset(1, 0).Resize(containers.Count, 1).value = "待计算"
        .Offset(1, 0).Resize(containers.Count, 8).Interior.Color = RGB(240, 240, 240)
    End With
End Sub
