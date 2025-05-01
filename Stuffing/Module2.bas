Attribute VB_Name = "Module2"
' 模块2: 货物排序算法
Option Explicit

' 排序策略枚举
Public Enum SortStrategy
    VolumeDesc        ' 体积降序
    WeightDesc        ' 重量降序
    StackableFirst    ' 可堆叠优先
    AreaHeightDesc    ' 底面积×高度降序
    VolumeDensityDesc ' 体积密度降序
    PrecedenceAsc     ' 装箱优先级升序
End Enum

' 主排序函数
Public Function SortBoxes(boxes As Collection, strategy As SortStrategy) As Collection
    On Error GoTo ErrorHandler ' 错误处理

    Dim boxArray() As CBox
    Dim i As Long
    
    ' 转换为数组
    ReDim boxArray(1 To boxes.Count)
    For i = 1 To boxes.Count
        Set boxArray(i) = boxes(i)
    Next
    
    ' 执行排序
    Select Case strategy
        Case VolumeDesc
            QuickSort boxArray, 1, boxes.Count, "VolumeCompare"
        Case WeightDesc
            QuickSort boxArray, 1, boxes.Count, "WeightCompare"
        Case StackableFirst
            QuickSort boxArray, 1, boxes.Count, "StackableCompare"
        Case AreaHeightDesc
            QuickSort boxArray, 1, boxes.Count, "AreaHeightCompare"
        Case VolumeDensityDesc
            QuickSort boxArray, 1, boxes.Count, "VolumeDensityCompare"
        Case PrecedenceAsc
            QuickSort boxArray, 1, boxes.Count, "PrecedenceCompare"
    End Select
    
    ' 转换回集合
    Dim sortedBoxes As New Collection
    For i = 1 To UBound(boxArray)
        sortedBoxes.Add boxArray(i)
    Next
    
    Set SortBoxes = sortedBoxes
    Exit Function

ErrorHandler:
    MsgBox "排序过程中出现错误: " & Err.Description
    Set SortBoxes = New Collection
End Function

' 快速排序算法
Private Sub QuickSort(arr() As CBox, ByVal low As Long, ByVal high As Long, compMethod As String)
    If low < high Then
        Dim pivot As Long
        pivot = Partition(arr, low, high, compMethod)
        QuickSort arr, low, pivot - 1, compMethod
        QuickSort arr, pivot + 1, high, compMethod
    End If
End Sub

Private Function Partition(arr() As CBox, ByVal low As Long, ByVal high As Long, compMethod As String) As Long
    Dim pivot As CBox
    Dim i As Long, j As Long
    Set pivot = arr(high)
    i = low - 1
    
    For j = low To high - 1
        If CompareBoxes(arr(j), pivot, compMethod) Then
            i = i + 1
            Swap arr, i, j
        End If
    Next
    Swap arr, i + 1, high
    Partition = i + 1
End Function

' 比较方法集合
Private Function CompareBoxes(a As CBox, b As CBox, method As String) As Boolean
    On Error GoTo ErrorHandler ' 错误处理

    Select Case method
        Case "VolumeCompare"
            CompareBoxes = (CalcEffectiveVolume(a) > CalcEffectiveVolume(b))
        Case "WeightCompare"
            CompareBoxes = (a.weight > b.weight)
        Case "StackableCompare"
            If a.Stackable <> b.Stackable Then
                CompareBoxes = b.Stackable  ' 不可堆叠的在前
            Else
                CompareBoxes = (CalcEffectiveVolume(a) > CalcEffectiveVolume(b))
            End If
        Case "AreaHeightCompare"
            CompareBoxes = (CalcAreaHeight(a) > CalcAreaHeight(b))
        Case "VolumeDensityCompare"
            CompareBoxes = (a.VolumeDensity > b.VolumeDensity)
        Case "PrecedenceCompare"
            CompareBoxes = (a.Precedence < b.Precedence)
    End Select
    Exit Function

ErrorHandler:
    MsgBox "比较过程中出现错误: " & Err.Description
    CompareBoxes = False
End Function

' 计算有效体积（考虑最佳旋转方向）
Private Function CalcEffectiveVolume(box As CBox) As Double
    Dim validOrientations As Collection
    Set validOrientations = GetValidOrientations(box)
    
    Dim maxVolume As Double
    Dim orientation As Variant
    For Each orientation In validOrientations
        maxVolume = Application.WorksheetFunction.Max(maxVolume, orientation(0) * orientation(1) * orientation(2))
    Next
    
    CalcEffectiveVolume = maxVolume
End Function

' 计算底面积×高度（用于分层策略）
Private Function CalcAreaHeight(box As CBox) As Double
    Dim validOrientations As Collection
    Set validOrientations = GetValidOrientations(box)
    
    Dim baseArea As Double
    Dim height As Double ' 修正变量名，与属性名区分
    Dim bestOrientation As Variant
    Dim orientation As Variant
    
    For Each orientation In validOrientations
        If orientation(2) < height Or height = 0 Then ' 修正为寻找高度最小的方向
            baseArea = orientation(0) * orientation(1)
            height = orientation(2)
            bestOrientation = orientation
        End If
    Next
    
    CalcAreaHeight = baseArea * height
End Function

' 获取所有有效摆放方向（考虑旋转限制）
Public Function GetValidOrientations(box As CBox) As Collection
    Dim orientations As New Collection
    Dim rotationAxes As String: rotationAxes = box.rotationAxes ' 修正变量名，与属性名区分
    
    ' 原始方向
    orientations.Add Array(box.length, box.width, box.height)
    
    ' 生成所有有效旋转组合
    If InStr(rotationAxes, "X") > 0 Then
        orientations.Add Array(box.length, box.height, box.width)   ' X轴旋转
    End If
    If InStr(rotationAxes, "Y") > 0 Then
        orientations.Add Array(box.height, box.width, box.length)  ' Y轴旋转
    End If
    If InStr(rotationAxes, "Z") > 0 Then
        orientations.Add Array(box.width, box.length, box.height)   ' Z轴旋转
    End If
    
    Set GetValidOrientations = orientations
End Function

' 辅助方法：交换数组元素
Private Sub Swap(arr() As CBox, i As Long, j As Long)
    Dim temp As CBox
    Set temp = arr(i)
    Set arr(i) = arr(j)
    Set arr(j) = temp
End Sub

Sub TestSorting()
    On Error GoTo ErrorHandler ' 错误处理

    Dim ws As Worksheet
    Set ws = Worksheets("Stuffing")
    
    Dim boxes As Collection
    Set boxes = ReadBoxDataFromSheet(ws)
    
    ' 测试不同排序策略
    Dim sortedBoxes As Collection
    
    ' 策略1：按体积降序
    Set sortedBoxes = SortBoxes(boxes, VolumeDesc)
    Debug.Print "=== 按体积排序 ==="
    PrintBoxList sortedBoxes
    
    ' 策略2：按重量降序
    Set sortedBoxes = SortBoxes(boxes, WeightDesc)
    Debug.Print "=== 按重量排序 ==="
    PrintBoxList sortedBoxes
    
    ' 策略3：分层策略排序
    Set sortedBoxes = SortBoxes(boxes, AreaHeightDesc)
    Debug.Print "=== 分层策略排序 ==="
    PrintBoxList sortedBoxes
    
    ' 新增策略：按体积密度降序
    Set sortedBoxes = SortBoxes(boxes, VolumeDensityDesc)
    Debug.Print "=== 按体积密度排序 ==="
    PrintBoxList sortedBoxes
    
    ' 新增策略：按装箱优先级升序
    Set sortedBoxes = SortBoxes(boxes, PrecedenceAsc)
    Debug.Print "=== 按装箱优先级排序 ==="
    PrintBoxList sortedBoxes

    Exit Sub

ErrorHandler:
    MsgBox "测试排序过程中出现错误: " & Err.Description
End Sub

' 辅助输出方法
Private Sub PrintBoxList(boxes As Collection)
    Dim box As CBox
    For Each box In boxes
        Debug.Print "ID:" & box.ID & _
                  " Dim:" & box.GetDimensions & _
                  " Vol:" & box.GetVolume & _
                  " W:" & box.weight & _
                  " S:" & box.Stackable & _
                  " R:" & box.rotationAxes & _
                  " Shape:" & box.Shape & _
                  " Fragility:" & box.Fragility & _
                  " COG(X,Y,Z):" & box.CenterOfGravityX & "," & box.CenterOfGravityY & "," & box.CenterOfGravityZ & _
                  " SpecialHandling:" & box.SpecialHandling & _
                  " Grouping:" & box.Grouping & _
                  " Precedence:" & box.Precedence & _
                  " VolumeDensity:" & box.VolumeDensity
    Next
End Sub

