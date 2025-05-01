Attribute VB_Name = "Module1"
Option Explicit

' ģ�鶥�����������������ؼ�����
Private Const TABLE_CTNR_USE As String = "CTNR_Use"
Private Const TABLE_CARGO_SPEC As String = "Cargo_Spec"

Public Function GetSelectedContainers(ws As Worksheet) As Collection
    Dim containers As New Collection
    Dim row As ListRow
    Dim tbl As ListObject
    
    ' ������Ƿ����
    If Not DoesTableExist(ws, TABLE_CTNR_USE) Then
        Debug.Print "��� " & TABLE_CTNR_USE & " �������ڹ����� " & ws.Name
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
        ' ��� CTNR ֵ�Ƿ�Ϊ�ջ����ֵ
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

' �Զ����Color�У�30��50֮��������ɫ������
Private Sub FillRandomColor(row As ListRow, columnIndex As Long)
    Randomize ' ��ʼ���������
    Dim randomColor As Long
    randomColor = Int((50 - 34 + 1) * Rnd + 34) ' ����34��50���������
    row.Range.Cells(1, columnIndex).Interior.ColorIndex = randomColor
End Sub

' �Զ���� Cargo_Spec ���� Color �к� VolumeDensity ��
Public Sub AutoFillCargoSpec(ws As Worksheet)
    Dim cargoTbl As ListObject
    Dim unitCell As Range
    Dim unitSystem As String ' "Metric" �� "Imperial"
    
    ' �����͵�λ��Ԫ���Ƿ����
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Then Exit Sub
    Set cargoTbl = ws.ListObjects(TABLE_CARGO_SPEC)
    Dim colorColIndex As Long
    colorColIndex = cargoTbl.ListColumns("Color").Index ' ��ȡColor�е�����
    
    Set unitCell = ws.Range("H4") ' ���赥λ��Ԫ��Ϊ H4���ɸ���ʵ�ʵ���
    unitSystem = UCase(Trim(Coalesce(unitCell, "METRIC")))
    
    Dim row As ListRow
    For Each row In cargoTbl.ListRows
        Dim cargoNameCell As Range, lengthCell As Range
        Set cargoNameCell = row.Range.Cells(1, 1) ' CargoName��
        Set lengthCell = row.Range.Cells(1, 2) ' Length��
        
        If Not IsEmpty(cargoNameCell.value) And Not IsEmpty(lengthCell.value) Then
            ' ��������ɫ��30��50��ColorIndex��
            FillRandomColor row, colorColIndex
            
            ' ����VolumeDensity������ԭ���߼���
            Dim length As Double, width As Double, height As Double, weight As Double
            length = GetNumericValue(lengthCell, 0)
            width = GetNumericValue(row.Range.Cells(1, 3), 0) ' Width��
            height = GetNumericValue(row.Range.Cells(1, 4), 0) ' Height��
            weight = GetNumericValue(row.Range.Cells(1, 5), 0) ' Weight��
            
            Dim volume As Double
            Select Case unitSystem
                Case "METRIC"
                    volume = (length * width * height) / 1000000 ' cm3תm3
                Case "IMPERIAL"
                    volume = (length * width * height) / 1728 ' in3תft3
            End Select
            
            If volume > 0 Then
                row.Range.Cells(1, cargoTbl.ListColumns("VolumeDensity").Index).value = weight / volume
            Else
                row.Range.Cells(1, cargoTbl.ListColumns("VolumeDensity").Index).value = 0
            End If
        End If
    Next row
End Sub


' ��ȡ�������ݣ��������ֶεĿ�ֵ��
Public Function ReadBoxDataFromSheet(ws As Worksheet) As Collection
    Dim boxes As New Collection
    Dim row As ListRow
    Dim cargoTbl As ListObject
    
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Then
        Debug.Print "��� " & TABLE_CARGO_SPEC & " �������ڹ����� " & ws.Name
        Exit Function
    End If
    
    Set cargoTbl = ws.ListObjects(TABLE_CARGO_SPEC)
    Dim cargoNameCol As ListColumn
    Set cargoNameCol = cargoTbl.ListColumns("CargoName")
    
    ' ��� cargoNameCol �� DataBodyRange �Ƿ�Ϊ��
    If cargoNameCol.DataBodyRange Is Nothing Then
        Debug.Print "CargoName �е������巶ΧΪ��"
        Exit Function
    End If
    
    For Each row In cargoTbl.ListRows
        Dim currentBox As New CBox
        Dim cargoNameValue As Variant
        ' ��ȡ��Ԫ���ֵ�������ֵ
        cargoNameValue = Coalesce(cargoNameCol.DataBodyRange.Cells(row.Index, 1), "")
        
        If Trim(cargoNameValue) = "" Then
            Set currentBox = Nothing
            GoTo SkipCurrentRow ' ������������Ϊ�յ���
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
            
            ' ���������ֶεĿ�ֵ���ַ�������Ĭ�Ͽ��ַ�������ֵ����Ĭ��0��
            .Shape = Trim(Coalesce(cargoTbl.ListColumns("Shape").DataBodyRange.Cells(row.Index, 1), ""))
            .Fragility = Trim(Coalesce(cargoTbl.ListColumns("Fragility").DataBodyRange.Cells(row.Index, 1), ""))
            
            .MaxStackLayers = CLng(Coalesce(cargoTbl.ListColumns("MaxStackLayers").DataBodyRange.Cells(row.Index, 1), "1"))
            .WeightCapacity = GetNumericValue(cargoTbl.ListColumns("WeightCapacity").DataBodyRange.Cells(row.Index, 1), 0#)
            .CanInvert = IsYes(cargoTbl.ListColumns("CanInvert").DataBodyRange.Cells(row.Index, 1)) ' ʹ������IsYes����������ֵ


            ' ����λ�ã���ֵĬ��0��ȷ����0~1֮�䣨ʵ��ҵ������ӷ�Χ��飩
            .CenterOfGravityX = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityX").DataBodyRange.Cells(row.Index, 1), 0#)
            .CenterOfGravityY = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityY").DataBodyRange.Cells(row.Index, 1), 0#)
            .CenterOfGravityZ = GetNumericValue(cargoTbl.ListColumns("CenterOfGravityZ").DataBodyRange.Cells(row.Index, 1), 0#)
            
            .SpecialHandling = Trim(Coalesce(cargoTbl.ListColumns("SpecialHandling").DataBodyRange.Cells(row.Index, 1), ""))
            .Grouping = Trim(Coalesce(cargoTbl.ListColumns("Grouping").DataBodyRange.Cells(row.Index, 1), ""))
            
            ' ���ȼ�����ֵĬ��0��0��ʾ�����ȼ�����ֵԽС���ȼ�Խ�ߣ�
            .Precedence = CLng(Coalesce(cargoTbl.ListColumns("Precedence").DataBodyRange.Cells(row.Index, 1), "0"))
            
            ' ����ܶȣ���ֵĬ��0���������0������ȷ�������Ϊ0�������߼�����
            .VolumeDensity = GetNumericValue(cargoTbl.ListColumns("VolumeDensity").DataBodyRange.Cells(row.Index, 1), 0#)
            
            ' ��ɫ�ֶΣ�����ԭ���߼���
            Dim colorCol As ListColumn
            Set colorCol = cargoTbl.ListColumns("Color")
            On Error Resume Next
            .Color = colorCol.DataBodyRange.Cells(row.Index, 1).Interior.Color
            On Error GoTo 0
        End With
        
        ' ���ɶ��ʵ��������QuantityΪ0��������������
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


' ��������������Yes/Noת������ǿ��׳�ԣ�
Private Function IsYes(value As Variant) As Boolean
    value = Coalesce(value, "No")
    IsYes = LCase(Trim(value)) = "yes"
End Function
' ������������ȫ��ȡ��ֵ��֧��Ĭ��ֵ��
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

' ���������������ֵ�����ֵ�������ַ�����
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



' ��ʼ������������
Public Sub InitResultArea(ws As Worksheet)
    With ws.Range("U3")
        If .CurrentRegion.Count > 1 Then .CurrentRegion.Clear ' ��ȫ��վ�����
        .Resize(1, 8).value = Array("��װ������", "��������", "ʣ��ռ�(L)", "ʣ��ռ�(W)", "ʣ��ռ�(H)", "����ʹ����", "װ�����", "�����б�")
        .EntireColumn.AutoFit
    End With
End Sub

' ����������������Ƿ����
Private Function DoesTableExist(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    Dim dummy As ListObject: Set dummy = ws.ListObjects(tableName)
    DoesTableExist = Not dummy Is Nothing
End Function




' ��������������CBox���ԣ������
Private Sub CopyBoxProperties(source As CBox, target As CBox)
    With source
        target.ID = .ID
        target.length = .length
        target.width = .width
        target.height = .height
        target.weight = .weight
        target.Quantity = 1 ' ÿ��ʵ�������̶�Ϊ1�������������ѭ������
        target.Stackable = .Stackable
        target.Rotatable = .Rotatable
        target.rotationAxes = .rotationAxes
        target.Color = .Color
        target.MaxStackLayers = .MaxStackLayers
        target.WeightCapacity = .WeightCapacity
        target.CanInvert = .CanInvert

    End With
End Sub

' ������������ȡ�Ǹ�����������Quantity���ܵĴ���ֵ��
Private Function Max(ByVal a As Long, ByVal b As Long) As Long
    Max = IIf(a >= b, a, b)
End Function



Sub TestDataLoading()
    Dim ws As Worksheet
    Set ws = Worksheets("Stuffing") ' ȷ��������������ȷ
    
    Dim containers As Collection
    Dim boxes As Collection
    
    ' �Զ����Color��VolumeDensity��
    AutoFillCargoSpec ws
    ' ��ʼ���������
    InitResultArea ws
    
    ' ��ȡ�û�ѡ��ļ�װ��
    Set containers = GetSelectedContainers(ws)
    ' ��ȡ��������
    Set boxes = ReadBoxDataFromSheet(ws)
    
    ' ------------------------
    ' �����װ����ϸ��Ϣ
    ' ------------------------
    Debug.Print "===== ��װ�������б� ====="
    If containers.Count = 0 Then
        Debug.Print "δ��ȡ����Ч��װ������"
    Else
        Dim c As CContainer ' ����ѭ������ c Ϊ CContainer ����
        For Each c In containers
            Debug.Print "? ��װ������: " & c.Name
            Debug.Print "   �ڲ��ߴ� (LxWxH): " & c.InnerLength & " x " & c.InnerWidth & " x " & c.InnerHeight
            Debug.Print "   �������: " & c.MaxLoad & " kg"
            Debug.Print "------------------------"
        Next c
    End If
    
    ' ------------------------
    ' ���������ϸ��Ϣ�������������ԣ�
    ' ------------------------
    Debug.Print vbNewLine & "===== ���������б� ====="
    If boxes.Count = 0 Then
        Debug.Print "δ��ȡ����Ч��������"
    Else
        Dim box As CBox ' ����ѭ������ box Ϊ CBox ����
        For Each box In boxes
            Debug.Print "? ����ID: " & box.ID
            Debug.Print "   �ߴ� (LxWxH): " & box.length & " x " & box.width & " x " & box.height
            Debug.Print "   ����: " & box.weight & " kg"
            Debug.Print "   ����: " & box.Quantity & " ��"
            Debug.Print "   �ɶѵ�: " & IIf(box.Stackable, "��", "��")
            Debug.Print "   ����ת: " & IIf(box.Rotatable, "��", "��")
            Debug.Print "   ��ת��: " & box.rotationAxes
            Debug.Print "   ��ɫ (RGB): " & box.Color
            ' �������Դ�ӡ
            Debug.Print "   �ɶѵ�����: " & box.MaxStackLayers
            Debug.Print "   ��������: " & box.WeightCapacity & " kg"
            Debug.Print "   ������: " & IIf(box.CanInvert, "��", "��")

            Debug.Print "   ��״: " & box.Shape
            Debug.Print "   ������: " & box.Fragility
            Debug.Print "   ����λ�� (X,Y,Z): " & box.CenterOfGravityX & ", " & box.CenterOfGravityY & ", " & box.CenterOfGravityZ
            Debug.Print "   ���⴦��: " & box.SpecialHandling
            Debug.Print "   ������Ϣ: " & box.Grouping
            Debug.Print "   װ�����ȼ�: " & box.Precedence
            Debug.Print "   ����ܶ�: " & box.VolumeDensity & " kg/m3" ' ���赥λΪ kg/m3���ɸ���ʵ�ʵ���
            Debug.Print "------------------------"
        Next box
    End If
    
    ' ��ʾͳ�ƽ��
    Debug.Print vbNewLine & "===== ͳ�ƽ�� ====="
    Debug.Print "��ѡ��װ������: " & containers.Count
    Debug.Print "��װ��������: " & boxes.Count & " ��" ' ע�⣺boxes.Count��ʵ��������Quantityչ����
    
    ' ʾ�������U3������ԭ�б���ʼ���߼���
    With ws.Range("U3")
        .Offset(1, 0).Resize(containers.Count, 1).value = "������"
        .Offset(1, 0).Resize(containers.Count, 8).Interior.Color = RGB(240, 240, 240)
    End With
End Sub
