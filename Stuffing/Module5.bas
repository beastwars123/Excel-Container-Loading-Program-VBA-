Attribute VB_Name = "Module5"
Option Explicit
Private Const TABLE_CTNR_USE As String = "CTNR_Use"
Private Const TABLE_CARGO_SPEC As String = "Cargo_Spec"

' 主装箱算法：FFD + 三维分层底左填充
Public Sub PackContainersFFD()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Stuffing")
    Dim containers As Collection: Set containers = GetSelectedContainers(ws)
    Dim boxes As Collection: Set boxes = ReadBoxDataFromSheet(ws)
    
    ' 按体积降序排序（FFD核心步骤）
    Dim sortedBoxes As Collection: Set sortedBoxes = SortBoxes(boxes, VolumeDesc)
    
    Dim containerLoads As New Collection ' 存储所有集装箱的装载结果
    
    For Each ctnr In containers
        Dim load As New CContainerLoad
        Set load.Container = ctnr
        Dim layers As New Collection ' 存储该集装箱的所有层
        
        ' 初始化第一层：z=0，可用区域为集装箱底面尺寸
        Dim firstLayer As New CLayer
        firstLayer.z = 0
        firstLayer.AddAvailableArea 0, 0, ctnr.InnerWidth, ctnr.InnerHeight
        layers.Add firstLayer
        
        ' 放置每个货物实例（展开Quantity）
        Dim box As CBox
        For Each box In sortedBoxes
            Dim placed As Boolean: placed = False
            Dim currentLayer As CLayer
            
            ' 首次适应：尝试在现有层中找到合适的区域
            For Each currentLayer In layers
                If PlaceBoxInLayer(box, currentLayer, load) Then
                    placed = True
                    Exit For
                End If
            Next
            
            ' 若现有层无法放置，且未超过集装箱高度，创建新层
            If Not placed And (load.UsedHeight + box.height <= ctnr.InnerHeight) Then
                Dim newLayer As New CLayer
                newLayer.z = load.UsedHeight
                newLayer.AddAvailableArea 0, 0, ctnr.InnerWidth, ctnr.InnerHeight
                layers.Add newLayer
                If PlaceBoxInLayer(box, newLayer, load) Then
                    placed = True
                End If
            End If
            
            ' 记录未放置的货物（可选：错误处理）
            If Not placed Then Debug.Print "货物 " & box.ID & " 无法装入集装箱 " & ctnr.Name
        Next box
        
        containerLoads.Add load
    Next ctnr
    
    ' 保存结果供可视化模块使用（全局变量或类属性）
    ThisWorkbook.Names.Add "GlobalLoadedContainers", containerLoads
End Sub

' 在层中放置货物并返回是否成功
Private Function PlaceBoxInLayer(box As CBox, layer As CLayer, load As CContainerLoad) As Boolean
    Dim area As Variant
    For Each area In layer.availableAreas
        ' 检查货物是否能放入当前区域（x,y方向尺寸匹配，z方向不超过集装箱长度）
        If box.width <= area(2) And box.height <= area(3) And (layer.z + box.length <= load.Container.InnerLength) Then
            ' 放置货物，分割剩余区域
            PlaceBoxInArea box, area, layer, load
            PlaceBoxInLayer = True
            Exit Function
        End If
    Next
End Function

' 在区域中放置货物并分割剩余空间（底左填充）
Private Sub PlaceBoxInArea(box As CBox, area As Variant, layer As CLayer, load As CContainerLoad)
    Dim x As Double: x = area(0)
    Dim y As Double: y = area(1)
    Dim areaWidth As Double: areaWidth = area(2)
    Dim areaHeight As Double: areaHeight = area(3)
    
    ' 记录货物位置和方向（简化为原始方向，可扩展旋转逻辑）
    Dim position As New Dictionary
    position.Add "ID", box.ID
    position.Add "x", x
    position.Add "y", y
    position.Add "z", layer.z
    position.Add "length", box.length
    position.Add "width", box.width
    position.Add "height", box.height
    
    ' 添加到装载结果
    load.LoadedBoxes.Add position
    load.UsedHeight = Max(load.UsedHeight, layer.z + box.length) ' z轴为长度方向（集装箱深度）
    load.UsedWeight = load.UsedWeight + box.weight
    
    ' 分割剩余区域：生成右侧和上方的可用区域
    layer.availableAreas.Remove (layer.availableAreas.IndexOf(area)) ' 移除当前使用的区域
    
    ' 右侧区域：x + 货物宽度 到 区域宽度
    If box.width < areaWidth Then
        layer.AddAvailableArea x + box.width, y, areaWidth - box.width, areaHeight
    End
    
    ' 上方区域：y + 货物高度 到 区域高度
    If box.height < areaHeight Then
        layer.AddAvailableArea x, y + box.height, areaWidth, areaHeight - box.height
    End
End Sub

' 辅助函数：获取最大值（处理Double类型）
Private Function Max(a As Double, b As Double) As Double
    If a >= b Then Max = a Else Max = b
End Function

' 其他现有函数（GetSelectedContainers, ReadBoxDataFromSheet, SortBoxes 等保持不变）
' 确保SortBoxes函数正确返回按体积降序排序的货物集合

Sub TestRealDataPacking()
    On Error GoTo TestErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Stuffing") ' 假设数据在"Stuffing"工作表
    
    ' 验证表格存在
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Or _
       Not DoesTableExist(ws, TABLE_CTNR_USE) Then
        MsgBox "请先在表格中输入货物（Cargo_Spec）和集装箱（CTNR_Use）数据", vbExclamation
        Exit Sub
    End If
    
    ' 执行装箱算法
    PackContainersFFD
    
    ' 获取全局装载结果（假设在主模块中已正确存储）
    Dim globalLoads As Collection
    Set globalLoads = ThisWorkbook.Names("GlobalLoadedContainers").RefersToRange.value
    
    ' 输出结果到立即窗口
    Debug.Print "======================"
    Debug.Print "= 实际数据装箱结果报告 ="
    Debug.Print "======================"
    Debug.Print "读取表格数据成功，开始验证装箱结果..."
    
    Dim load As CContainerLoad
    For Each load In globalLoads
        With load.Container
            Debug.Print vbNewLine & "【集装箱信息】"
            Debug.Print "型号: " & .Name
            Debug.Print "内部尺寸 (LxWxH): " & .InnerLength & " x " & .InnerWidth & " x " & .InnerHeight & " cm"
            Debug.Print "最大载重: " & .MaxLoad & " kg"
        End With
        
        Debug.Print "【装载摘要】"
        Debug.Print "已用高度: " & load.UsedHeight & " cm (" & Format(load.UsedHeight / load.Container.InnerHeight, "0.00%") & ")"
        Debug.Print "已用重量: " & load.UsedWeight & " kg (" & Format(load.UsedWeight / load.Container.MaxLoad, "0.00%") & ")"
        Debug.Print "装载货物数量: " & load.LoadedBoxes.Count & " 件（原始Quantity展开后）"
        
        ' 输出每个货物的装载位置（前5个示例，避免输出过长）
        Debug.Print vbNewLine & "【前5个装载货物详情】"
        Dim i As Long
        Dim position As New Dictionary
        For i = 1 To Min(5, load.LoadedBoxes.Count)
            Set position = load.LoadedBoxes(i)
            Debug.Print "货物ID: " & position("ID")
            Debug.Print "摆放位置 (X,Y,Z): " & position("x") & " x " & position("y") & " x " & position("z")
            Debug.Print "占用尺寸 (LxWxH): " & position("length") & " x " & position("width") & " x " & position("height")
            Debug.Print "------------------------"
        Next i
        
        ' 验证关键约束（可根据业务添加更多验证）
        If load.UsedWeight > load.Container.MaxLoad Then
            Debug.Print "?? 警告：载重超过集装箱最大载重！"
        End If
        If load.UsedHeight > load.Container.InnerHeight Then
            Debug.Print "?? 警告：使用高度超过集装箱内部高度！"
        End If
    Next load
    
    Exit Sub

TestErrorHandler:
    Debug.Print "测试过程中出现错误：" & Err.Description
    Resume Next
End Sub

' 辅助函数：获取较小值（用于限制输出数量）
Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

Private Function DoesTableExist(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    Dim dummy As ListObject: Set dummy = ws.ListObjects(tableName)
    DoesTableExist = Not dummy Is Nothing
End Function
