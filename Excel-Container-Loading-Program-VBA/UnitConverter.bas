Attribute VB_Name = "UnitConverter"
Option Explicit
Option Private Module 'ģ���ڹ���Ĭ��˽�У�������������ʽ����Ϊ Public

'#Region "���Ͷ���"
Private Type TUnitConverter
    Initialized As Boolean
    ConversionRules As Object 'Scripting.Dictionary
    ConversionMatrix As Variant
    LastUpdate As Date
End Type
'#End Region

'#Region "ģ�鼶����"
Private this As TUnitConverter 'Ψһ��ת����ʵ��
'#End Region

Public Sub InitializeConverter(Optional forceReload As Boolean = False)
    If Not forceReload And this.Initialized Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    '������й���
    Set this.ConversionRules = CreateObject("Scripting.Dictionary")
    this.ConversionRules.CompareMode = 1 '�ı��Ƚ�ģʽ
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UnitConversions")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1001, "UnitConverter", "δ�ҵ� 'UnitConversions' ������"
    End If
    
    '������ȡ������������
    Dim dataRange As Range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).row
    If lastRow < 5 Then Exit Sub '������ʱ�˳�
    
    Set dataRange = ws.Range("Q5:V" & lastRow)
    this.ConversionMatrix = dataRange.value
    
    Dim i As Long
    Dim currentRow As Long '���ڱ��浱ǰ�к�
    For i = LBound(this.ConversionMatrix, 1) To UBound(this.ConversionMatrix, 1)
        currentRow = i
        Dim rule As New CConversionRule 'ȷ��ÿ�δ����µĹ������ʵ��
        With rule
            .fromUnit = UCase(Trim(NzVBA(this.ConversionMatrix(i, 1), "")))
            .toUnit = UCase(Trim(NzVBA(this.ConversionMatrix(i, 3), "")))
            .Factor = Val(NzVBA(this.ConversionMatrix(i, 2), 1))
            .Offset = Val(NzVBA(this.ConversionMatrix(i, 4), 0))
            .FormulaType = UCase(NzVBA(this.ConversionMatrix(i, 5), "LINEAR"))
            .ValidationHash = GetValidationHash(i + 4) '�кŴӵ�5�п�ʼ
        End With
        
        '������֤
        On Error Resume Next
        ValidateRule rule, i + 4
        On Error GoTo ErrorHandler
        
        '��ע����������Ƴ�˫��ע���߼���
        RegisterRule rule
    Next i
    
    this.Initialized = True
    this.LastUpdate = Now
    
    Exit Sub
    
ErrorHandler:
    Err.Raise vbObjectError + 1000, "UnitConverter", _
        "��ʼ��ʧ��: " & Err.Description & " (�к�:" & currentRow + 4 & ")"
End Sub

Public Function ConvertUnit( _
    ByVal value As Double, _
    ByVal fromUnit As String, _
    ByVal toUnit As String) As Variant
    
    '���������֤
    If IsEmpty(value) Or IsNull(value) Then
        ConvertUnit = CVErr(xlErrValue)
        Exit Function
    End If
    If Trim(fromUnit) = "" Or Trim(toUnit) = "" Then
        ConvertUnit = CVErr(xlErrValue)
        Exit Function
    End If
    
    If Not this.Initialized Then InitializeConverter
    
    '���� fromUnit �� toUnit ��ͬ�����
    If UCase(Trim(fromUnit)) = UCase(Trim(toUnit)) Then
        ConvertUnit = value
        Exit Function
    End If
    
    Dim key As String
    key = BuildKey(UCase(Trim(fromUnit)), UCase(Trim(toUnit)))
'    Debug.Print "���Բ��ҵļ�: " & key '������Բ��ҵļ����������
    
    If Not this.ConversionRules.Exists(key) Then
        ConvertUnit = CVErr(xlErrNA)
'        Debug.Print "δ�ҵ�ƥ��ļ�: " & key '���δ�ҵ�ƥ�������Ϣ
        Exit Function
    End If
    
    Dim rule As CConversionRule
    Set rule = this.ConversionRules(key)
    
'    Debug.Print "�ҵ��Ĺ��� - ��: " & key
'    Debug.Print "�ҵ��Ĺ��� - fromUnit: " & rule.fromUnit
'    Debug.Print "�ҵ��Ĺ��� - toUnit: " & rule.toUnit
'    Debug.Print "�ҵ��Ĺ��� - Factor: " & rule.Factor
    
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

'#Region "˽�з���"
Private Sub ValidateRule(ByRef rule As CConversionRule, ByVal sourceRow As Long)
    Const validTypes As String = "LINEAR,LINEARWITHOFFSET,INVERSE"
    
    With rule
        If Len(rule.fromUnit) * Len(rule.toUnit) = 0 Then
            Err.Raise vbObjectError + 1001, , "��λ���Ʋ���Ϊ�� (�к�:" & sourceRow & ")"
        End If
        
        If rule.Factor = 0 And rule.FormulaType <> "INVERSE" Then
            Err.Raise vbObjectError + 1002, , "ת�����Ӳ���Ϊ�� (�к�:" & sourceRow & ")"
        End If
        
        If InStr(1, validTypes, rule.FormulaType, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1003, , _
                "��Ч�Ĺ�ʽ����: " & rule.FormulaType & " (�к�:" & sourceRow & ")"
        End If
    End With
End Sub

Private Sub RegisterRule(ByRef rule As CConversionRule)
    '��ע���������ɾ��������������߼���
    Dim key As String
    key = BuildKey(rule.fromUnit, rule.toUnit)
'    Debug.Print "ע��ļ�: " & key '���ע��ļ����������
    If Not this.ConversionRules.Exists(key) Then
        Dim newRule As New CConversionRule '�����µĹ������ʵ��
        With newRule
            .fromUnit = rule.fromUnit
            .toUnit = rule.toUnit
            .Factor = rule.Factor
            .Offset = rule.Offset
            .FormulaType = rule.FormulaType
            .ValidationHash = rule.ValidationHash
        End With
        this.ConversionRules.Add key, newRule '���µĹ����������ֵ�
    End If
End Sub

Private Function BuildKey(ByVal fromUnit As String, ByVal toUnit As String) As String
    'ȷ�����Ĺ�����ʽһ�£�ȥ�������ַ�
    BuildKey = Replace(Trim(fromUnit), " ", "") & "��" & Replace(Trim(toUnit), " ", "")
End Function

Private Function GetValidationHash(ByVal rowNumber As Long) As String
    '���ɰ���������Ϣ��У���ϣ
    GetValidationHash = "R" & rowNumber & "_" & Format(Now, "yyyymmddhhnnss")
End Function

'�Զ��� Nz ����
Private Function NzVBA(ByVal value As Variant, ByVal defaultValue As Variant) As Variant
    If IsNull(value) Then
        NzVBA = defaultValue
    Else
        NzVBA = value
    End If
End Function

