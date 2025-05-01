Attribute VB_Name = "Module2"
' ģ��2: ���������㷨
Option Explicit

' �������ö��
Public Enum SortStrategy
    VolumeDesc        ' �������
    WeightDesc        ' ��������
    StackableFirst    ' �ɶѵ�����
    AreaHeightDesc    ' ��������߶Ƚ���
    VolumeDensityDesc ' ����ܶȽ���
    PrecedenceAsc     ' װ�����ȼ�����
End Enum

' ��������
Public Function SortBoxes(boxes As Collection, strategy As SortStrategy) As Collection
    On Error GoTo ErrorHandler ' ������

    Dim boxArray() As CBox
    Dim i As Long
    
    ' ת��Ϊ����
    ReDim boxArray(1 To boxes.Count)
    For i = 1 To boxes.Count
        Set boxArray(i) = boxes(i)
    Next
    
    ' ִ������
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
    
    ' ת���ؼ���
    Dim sortedBoxes As New Collection
    For i = 1 To UBound(boxArray)
        sortedBoxes.Add boxArray(i)
    Next
    
    Set SortBoxes = sortedBoxes
    Exit Function

ErrorHandler:
    MsgBox "��������г��ִ���: " & Err.Description
    Set SortBoxes = New Collection
End Function

' ���������㷨
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

' �ȽϷ�������
Private Function CompareBoxes(a As CBox, b As CBox, method As String) As Boolean
    On Error GoTo ErrorHandler ' ������

    Select Case method
        Case "VolumeCompare"
            CompareBoxes = (CalcEffectiveVolume(a) > CalcEffectiveVolume(b))
        Case "WeightCompare"
            CompareBoxes = (a.weight > b.weight)
        Case "StackableCompare"
            If a.Stackable <> b.Stackable Then
                CompareBoxes = b.Stackable  ' ���ɶѵ�����ǰ
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
    MsgBox "�ȽϹ����г��ִ���: " & Err.Description
    CompareBoxes = False
End Function

' ������Ч��������������ת����
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

' �����������߶ȣ����ڷֲ���ԣ�
Private Function CalcAreaHeight(box As CBox) As Double
    Dim validOrientations As Collection
    Set validOrientations = GetValidOrientations(box)
    
    Dim baseArea As Double
    Dim height As Double ' ������������������������
    Dim bestOrientation As Variant
    Dim orientation As Variant
    
    For Each orientation In validOrientations
        If orientation(2) < height Or height = 0 Then ' ����ΪѰ�Ҹ߶���С�ķ���
            baseArea = orientation(0) * orientation(1)
            height = orientation(2)
            bestOrientation = orientation
        End If
    Next
    
    CalcAreaHeight = baseArea * height
End Function

' ��ȡ������Ч�ڷŷ��򣨿�����ת���ƣ�
Public Function GetValidOrientations(box As CBox) As Collection
    Dim orientations As New Collection
    Dim rotationAxes As String: rotationAxes = box.rotationAxes ' ������������������������
    
    ' ԭʼ����
    orientations.Add Array(box.length, box.width, box.height)
    
    ' ����������Ч��ת���
    If InStr(rotationAxes, "X") > 0 Then
        orientations.Add Array(box.length, box.height, box.width)   ' X����ת
    End If
    If InStr(rotationAxes, "Y") > 0 Then
        orientations.Add Array(box.height, box.width, box.length)  ' Y����ת
    End If
    If InStr(rotationAxes, "Z") > 0 Then
        orientations.Add Array(box.width, box.length, box.height)   ' Z����ת
    End If
    
    Set GetValidOrientations = orientations
End Function

' ������������������Ԫ��
Private Sub Swap(arr() As CBox, i As Long, j As Long)
    Dim temp As CBox
    Set temp = arr(i)
    Set arr(i) = arr(j)
    Set arr(j) = temp
End Sub

Sub TestSorting()
    On Error GoTo ErrorHandler ' ������

    Dim ws As Worksheet
    Set ws = Worksheets("Stuffing")
    
    Dim boxes As Collection
    Set boxes = ReadBoxDataFromSheet(ws)
    
    ' ���Բ�ͬ�������
    Dim sortedBoxes As Collection
    
    ' ����1�����������
    Set sortedBoxes = SortBoxes(boxes, VolumeDesc)
    Debug.Print "=== ��������� ==="
    PrintBoxList sortedBoxes
    
    ' ����2������������
    Set sortedBoxes = SortBoxes(boxes, WeightDesc)
    Debug.Print "=== ���������� ==="
    PrintBoxList sortedBoxes
    
    ' ����3���ֲ��������
    Set sortedBoxes = SortBoxes(boxes, AreaHeightDesc)
    Debug.Print "=== �ֲ�������� ==="
    PrintBoxList sortedBoxes
    
    ' �������ԣ�������ܶȽ���
    Set sortedBoxes = SortBoxes(boxes, VolumeDensityDesc)
    Debug.Print "=== ������ܶ����� ==="
    PrintBoxList sortedBoxes
    
    ' �������ԣ���װ�����ȼ�����
    Set sortedBoxes = SortBoxes(boxes, PrecedenceAsc)
    Debug.Print "=== ��װ�����ȼ����� ==="
    PrintBoxList sortedBoxes

    Exit Sub

ErrorHandler:
    MsgBox "������������г��ִ���: " & Err.Description
End Sub

' �����������
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

