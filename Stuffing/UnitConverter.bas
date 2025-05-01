Attribute VB_Name = "UnitConverter"
Option Explicit
Option Private Module '模块内过程默认私有，公共过程需显式声明为 Public

'#Region "类型定义"
Private Type TUnitConverter
    Initialized As Boolean
    ConversionRules As Object 'Scripting.Dictionary
    ConversionMatrix As Variant
    LastUpdate As Date
End Type
'#End Region

'#Region "模块级变量"
Private this As TUnitConverter '唯一的转换器实例
'#End Region

Public Sub InitializeConverter(Optional forceReload As Boolean = False)
    If Not forceReload And this.Initialized Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    '清空现有规则
    Set this.ConversionRules = CreateObject("Scripting.Dictionary")
    this.ConversionRules.CompareMode = 1 '文本比较模式
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UnitConversions")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1001, "UnitConverter", "未找到 'UnitConversions' 工作表。"
    End If
    
    '批量读取数据提升性能
    Dim dataRange As Range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).row
    If lastRow < 5 Then Exit Sub '无数据时退出
    
    Set dataRange = ws.Range("Q5:V" & lastRow)
    this.ConversionMatrix = dataRange.value
    
    Dim i As Long
    Dim currentRow As Long '用于保存当前行号
    For i = LBound(this.ConversionMatrix, 1) To UBound(this.ConversionMatrix, 1)
        currentRow = i
        Dim rule As New CConversionRule '确保每次创建新的规则对象实例
        With rule
            .fromUnit = UCase(Trim(NzVBA(this.ConversionMatrix(i, 1), "")))
            .toUnit = UCase(Trim(NzVBA(this.ConversionMatrix(i, 3), "")))
            .Factor = Val(NzVBA(this.ConversionMatrix(i, 2), 1))
            .Offset = Val(NzVBA(this.ConversionMatrix(i, 4), 0))
            .FormulaType = UCase(NzVBA(this.ConversionMatrix(i, 5), "LINEAR"))
            .ValidationHash = GetValidationHash(i + 4) '行号从第5行开始
        End With
        
        '数据验证
        On Error Resume Next
        ValidateRule rule, i + 4
        On Error GoTo ErrorHandler
        
        '仅注册正向规则（移除双向注册逻辑）
        RegisterRule rule
    Next i
    
    this.Initialized = True
    this.LastUpdate = Now
    
    Exit Sub
    
ErrorHandler:
    Err.Raise vbObjectError + 1000, "UnitConverter", _
        "初始化失败: " & Err.Description & " (行号:" & currentRow + 4 & ")"
End Sub

Public Function ConvertUnit( _
    ByVal value As Double, _
    ByVal fromUnit As String, _
    ByVal toUnit As String) As Variant
    
    '输入参数验证
    If IsEmpty(value) Or IsNull(value) Then
        ConvertUnit = CVErr(xlErrValue)
        Exit Function
    End If
    If Trim(fromUnit) = "" Or Trim(toUnit) = "" Then
        ConvertUnit = CVErr(xlErrValue)
        Exit Function
    End If
    
    If Not this.Initialized Then InitializeConverter
    
    '处理 fromUnit 和 toUnit 相同的情况
    If UCase(Trim(fromUnit)) = UCase(Trim(toUnit)) Then
        ConvertUnit = value
        Exit Function
    End If
    
    Dim key As String
    key = BuildKey(UCase(Trim(fromUnit)), UCase(Trim(toUnit)))
'    Debug.Print "尝试查找的键: " & key '输出尝试查找的键，方便调试
    
    If Not this.ConversionRules.Exists(key) Then
        ConvertUnit = CVErr(xlErrNA)
'        Debug.Print "未找到匹配的键: " & key '输出未找到匹配键的信息
        Exit Function
    End If
    
    Dim rule As CConversionRule
    Set rule = this.ConversionRules(key)
    
'    Debug.Print "找到的规则 - 键: " & key
'    Debug.Print "找到的规则 - fromUnit: " & rule.fromUnit
'    Debug.Print "找到的规则 - toUnit: " & rule.toUnit
'    Debug.Print "找到的规则 - Factor: " & rule.Factor
    
    Select Case rule.FormulaType
        Case "LINEAR"
            ConvertUnit = value * rule.Factor
        Case "LINEARWITHOFFSET"
            ConvertUnit = (value * rule.Factor) + rule.Offset
        Case "INVERSE"
            If value = 0 Then
                ConvertUnit = CVErr(xlErrDiv0)
            Else
                ConvertUnit = rule.Factor / value
            End If
        Case Else
            ConvertUnit = CVErr(xlErrValue)
    End Select
End Function

'#Region "私有方法"
Private Sub ValidateRule(ByRef rule As CConversionRule, ByVal sourceRow As Long)
    Const validTypes As String = "LINEAR,LINEARWITHOFFSET,INVERSE"
    
    With rule
        If Len(rule.fromUnit) * Len(rule.toUnit) = 0 Then
            Err.Raise vbObjectError + 1001, , "单位名称不能为空 (行号:" & sourceRow & ")"
        End If
        
        If rule.Factor = 0 And rule.FormulaType <> "INVERSE" Then
            Err.Raise vbObjectError + 1002, , "转换因子不能为零 (行号:" & sourceRow & ")"
        End If
        
        If InStr(1, validTypes, rule.FormulaType, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1003, , _
                "无效的公式类型: " & rule.FormulaType & " (行号:" & sourceRow & ")"
        End If
    End With
End Sub

Private Sub RegisterRule(ByRef rule As CConversionRule)
    '仅注册正向规则（删除反向规则生成逻辑）
    Dim key As String
    key = BuildKey(rule.fromUnit, rule.toUnit)
'    Debug.Print "注册的键: " & key '输出注册的键，方便调试
    If Not this.ConversionRules.Exists(key) Then
        Dim newRule As New CConversionRule '创建新的规则对象实例
        With newRule
            .fromUnit = rule.fromUnit
            .toUnit = rule.toUnit
            .Factor = rule.Factor
            .Offset = rule.Offset
            .FormulaType = rule.FormulaType
            .ValidationHash = rule.ValidationHash
        End With
        this.ConversionRules.Add key, newRule '将新的规则对象存入字典
    End If
End Sub

Private Function BuildKey(ByVal fromUnit As String, ByVal toUnit As String) As String
    '确保键的构建方式一致，去除多余字符
    BuildKey = Replace(Trim(fromUnit), " ", "") & "→" & Replace(Trim(toUnit), " ", "")
End Function

Private Function GetValidationHash(ByVal rowNumber As Long) As String
    '生成包含行列信息的校验哈希
    GetValidationHash = "R" & rowNumber & "_" & Format(Now, "yyyymmddhhnnss")
End Function

'自定义 Nz 函数
Private Function NzVBA(ByVal value As Variant, ByVal defaultValue As Variant) As Variant
    If IsNull(value) Then
        NzVBA = defaultValue
    Else
        NzVBA = value
    End If
End Function

