Attribute VB_Name = "Module4"
' 模块4: 可视化模块
Option Explicit

' 颜色配置
Private Const CONTAINER_COLOR As Long = &HD0CECE
Private Const COLOR_PALETTE As String = "FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,800080"

' 主可视化函数
Public Sub GeneratePackingDiagram(ws As Worksheet, results As Collection)
    ' 清空现有图形
    DeleteOldShapes ws
    
    ' 设置绘图参数
    Dim startLeft As Double: startLeft = 20
    Dim startTop As Double: startTop = 100
    Dim scaleFactor As Double: scaleFactor = 0.2 ' 0.2 points/mm
    
    Dim result As Variant
    For Each result In results
        ' 绘制单个集装箱视图
        DrawContainerView ws, result, startLeft, startTop, scaleFactor
        
        ' 更新起始位置
        startLeft = startLeft + GetContainerWidth(result, scaleFactor) + 50
        If startLeft > 1000 Then
            startLeft = 20
            startTop = startTop + GetContainerHeight(result, scaleFactor) + 50
        End If
    Next
    
    ' 添加图例
    AddColorLegend ws, startTop
End Sub

' 绘制单个集装箱视图
Private Sub DrawContainerView(ws As Worksheet, result As PackingResult, left As Double, top As Double, scaleFactor As Double)
    ' 绘制集装箱轮廓
    With ws.Shapes.AddShape(msoShapeRectangle, left, top, result.Container.InnerLength * scaleFactor, result.Container.InnerHeight * scaleFactor)
        .Name = "Cont_" & result.Container.Name
        .Fill.ForeColor.RGB = CONTAINER_COLOR
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    ' 绘制所有货物
    Dim boxInfo As Object
    For Each boxInfo In result.PackedBoxes
        DrawBox3D ws, boxInfo, left, top, scaleFactor, result.Container
    Next
    
    ' 添加标签
    AddContainerLabel ws, result, left, top, scaleFactor
End Sub

' 绘制单个货物的3D投影
Private Sub DrawBox3D(ws As Worksheet, boxInfo As Object, contLeft As Double, contTop As Double, scaleFactor As Double, cont As CContainer)
    Dim box As CBox
    Set box = boxInfo("Box")
    
    Dim pos() As Variant
    pos = boxInfo("Position")
    Dim orientation() As Double
    orientation = boxInfo("Orientation")
    
    ' 计算投影坐标
    Dim projLeft As Double: projLeft = contLeft + pos(0) * scaleFactor
    Dim projTop As Double: projTop = contTop + pos(2) * scaleFactor ' Z轴投影到垂直方向
    
    ' 绘制3D效果
    With ws.Shapes.AddShape(msoShapeRectangle, projLeft, projTop, _
                           orientation(0) * scaleFactor, _
                           orientation(2) * scaleFactor)
        .Name = "Box_" & box.ID
        .Fill.ForeColor.RGB = GetSizeColor(box.length, box.width, box.height)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Rotation = 5 ' 创建3D投影效果
    End With
    
    ' 添加尺寸标签
    If orientation(0) * scaleFactor > 30 And orientation(2) * scaleFactor > 15 Then
        With ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            projLeft + 2, projTop + 2, 50, 15)
            .TextFrame.Characters.Text = box.ID & " " & _
                Format(orientation(0), "0") & "x" & _
                Format(orientation(1), "0") & "x" & _
                Format(orientation(2), "0")
            .TextFrame.Characters.Font.Size = 8
        End With
    End If
End Sub

' 根据尺寸获取颜色编码
Private Function GetSizeColor(L As Double, W As Double, H As Double) As Long
    Dim sizeKey As String
    sizeKey = Format(L, "000") & Format(W, "000") & Format(H, "000")
    
    Dim colors() As String
    colors = Split(COLOR_PALETTE, ",")
    
    ' 简单哈希算法生成颜色索引
    Dim hash As Long
    hash = Val(Right(sizeKey, 3)) Mod (UBound(colors) + 1)
    
    GetSizeColor = Val("&H" & colors(hash))
End Function

' 添加集装箱标签
Private Sub AddContainerLabel(ws As Worksheet, result As PackingResult, _
                            left As Double, top As Double, scaleFactor As Double)
    With ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        left, top - 20, result.Container.InnerLength * scaleFactor, 20)
        .TextFrame.Characters.Text = result.Container.Name & _
            " (利用率: " & Format(result.Efficiency * 100, "0.0") & "%)"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
End Sub

' 添加颜色图例
Private Sub AddColorLegend(ws As Worksheet, topPos As Double)
    Dim colors() As String
    colors = Split(COLOR_PALETTE, ",")
    
    Dim i As Long
    For i = 0 To UBound(colors)
        With ws.Shapes.AddShape(msoShapeRectangle, 20 + i * 60, topPos, 50, 20)
            .Fill.ForeColor.RGB = Val("&H" & colors(i))
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        With ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            20 + i * 60, topPos + 25, 50, 15)
            .TextFrame.Characters.Text = "尺寸组" & i + 1
            .TextFrame.Characters.Font.Size = 8
            .TextFrame.HorizontalAlignment = xlHAlignCenter
        End With
    Next
End Sub

' 辅助函数
Private Sub DeleteOldShapes(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type <> msoFormControl Then
            shp.Delete
        End If
    Next
End Sub

Private Function GetContainerWidth(result As PackingResult, scaleFactor As Double) As Double
    GetContainerWidth = result.Container.InnerLength * scaleFactor
End Function

Private Function GetContainerHeight(result As PackingResult, scaleFactor As Double) As Double
    GetContainerHeight = result.Container.InnerHeight * scaleFactor
End Function

