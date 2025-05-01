Attribute VB_Name = "Module3"
'============================================
' ģ������: PackingAlgorithm
' ����: ��άװ������㷨ʵ��
' ������ģ��: CBox, CContainer, CPackingResult, CSpaceArea
'============================================
Option Explicit

' �������壨ͳһ�ֵ������
Private Const KEY_POSITION As String = "position"
Private Const KEY_DIMS As String = "dims"
Private Const KEY_SPACE_INDEX As String = "spaceindex"

' ������װ��װ�����
Public Function PackSingleContainer( _
    cont As CContainer, _
    ByRef boxes As Collection) As CPackingResult
    
    Dim result As New CPackingResult
    Set result.Container = cont
    
    ' ��ʼ��ʣ��ռ�
    Dim initialSpace As New CSpaceArea
    initialSpace.Initialize 0, 0, 0, cont.InnerLength, cont.InnerWidth, cont.InnerHeight
    result.RemainingSpaces.Add initialSpace
    
    ' ���ֲ��������
    Dim sortedBoxes As Collection
    Set sortedBoxes = SortBoxes(boxes, AreaHeightDesc)
    
    Dim box As CBox
    For Each box In sortedBoxes
        If result.WeightUsage + box.weight > cont.MaxLoad Then Exit For
        
        Dim bestPlacement As Object
        Set bestPlacement = FindBestPlacement(box, result.RemainingSpaces)
        
        If Not bestPlacement Is Nothing Then
            ' ��¼װ����Ϣ
            Dim packedInfo As Object
            Set packedInfo = CreatePackingInfo(box, bestPlacement)
            result.PackedBoxes.Add packedInfo
            
            ' ��������
            result.WeightUsage = result.WeightUsage + box.weight
            
            ' �ָ�ʣ��ռ�
            UpdateSpaces result.RemainingSpaces, bestPlacement
        End If
    Next
    
    ' ����ռ�������
    result.Efficiency = CalculateEfficiency(result)
    Set PackSingleContainer = result
End Function

' ����װ����Ϣ�ֵ�
Private Function CreatePackingInfo(box As CBox, placement As Object) As Object
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    
    ' ��Ӽ���������֤
    If Not placement.Exists(KEY_POSITION) Then Err.Raise 438, , "ȱʧλ����Ϣ"
    If Not placement.Exists(KEY_DIMS) Then Err.Raise 438, , "ȱʧ������Ϣ"
    
    info.Add "Box", box
    info.Add KEY_POSITION, placement(KEY_POSITION)
    info.Add "Orientation", placement(KEY_DIMS)
    
    Set CreatePackingInfo = info
End Function

' ������ѷ���λ��
Private Function FindBestPlacement(box As CBox, spaces As Collection) As Object
    Dim bestScore As Double: bestScore = 0
    Dim bestPlacement As Object: Set bestPlacement = Nothing
    
    Dim i As Long
    For i = 1 To spaces.Count
        Dim space As CSpaceArea
        Set space = spaces(i)
        
        Dim orientations As Collection
        Set orientations = GetValidOrientations(box)
        
        Dim o As Variant
        For Each o In orientations
            If IsValidPlacement(space, o) Then
                Dim score As Double
                score = CalculatePlacementScore(space, o)
                
                If score > bestScore Then
                    bestScore = score
                    Set bestPlacement = CreatePlacement(space, o, i)
                End If
            End If
        Next
    Next
    
    Set FindBestPlacement = bestPlacement
End Function

' ��֤���ÿ�����
Private Function IsValidPlacement(space As CSpaceArea, dims As Variant) As Boolean
    IsValidPlacement = (dims(0) <= space.length) And _
                      (dims(1) <= space.width) And _
                      (dims(2) <= space.height)
End Function

' ����������֣��߶����ȣ�
Private Function CalculatePlacementScore(space As CSpaceArea, dims As Variant) As Double
    CalculatePlacementScore = dims(2) / space.height  ' �߶�������
End Function

' ����������Ϣ
Private Function CreatePlacement(space As CSpaceArea, dims As Variant, spaceIndex As Long) As Object
    Dim placement As Object
    Set placement = CreateObject("Scripting.Dictionary")
    
    placement.Add KEY_POSITION, Array(space.x, space.y, space.z)
    placement.Add KEY_DIMS, dims
    placement.Add KEY_SPACE_INDEX, spaceIndex
    
    Set CreatePlacement = placement
End Function

' ����ʣ��ռ䣨��ά�ռ�ָ
Private Sub UpdateSpaces(spaces As Collection, placement As Object)
    Dim spaceIndex As Long: spaceIndex = placement(KEY_SPACE_INDEX)
    Dim dims As Variant: dims = placement(KEY_DIMS)
    
    Dim originalSpace As CSpaceArea
    Set originalSpace = spaces(spaceIndex)
    spaces.Remove spaceIndex
    
    ' �ҷ�ʣ��ռ�
    If originalSpace.length > dims(0) Then
        Dim rightSpace As New CSpaceArea
        rightSpace.Initialize originalSpace.x + dims(0), _
                            originalSpace.y, _
                            originalSpace.z, _
                            originalSpace.length - dims(0), _
                            originalSpace.width, _
                            originalSpace.height
        spaces.Add rightSpace
    End If
    
    ' �Ϸ�ʣ��ռ�
    If originalSpace.width > dims(1) Then
        Dim upperSpace As New CSpaceArea
        upperSpace.Initialize originalSpace.x, _
                            originalSpace.y + dims(1), _
                            originalSpace.z, _
                            dims(0), _
                            originalSpace.width - dims(1), _
                            originalSpace.height
        spaces.Add upperSpace
    End If
    
    ' ǰ��ʣ��ռ�
    If originalSpace.height > dims(2) Then
        Dim frontSpace As New CSpaceArea
        frontSpace.Initialize originalSpace.x, _
                            originalSpace.y, _
                            originalSpace.z + dims(2), _
                            dims(0), _
                            dims(1), _
                            originalSpace.height - dims(2)
        spaces.Add frontSpace
    End If
End Sub

' ����ռ�������
Private Function CalculateEfficiency(result As CPackingResult) As Double
    Dim totalVol As Double
    Dim contVol As Double
    
    With result.Container
        contVol = .InnerLength * .InnerWidth * .InnerHeight
    End With
    
    Dim item As Object
    For Each item In result.PackedBoxes
        Dim box As CBox
        Set box = item("Box")
        totalVol = totalVol + box.length * box.width * box.height
    Next
    
    CalculateEfficiency = totalVol / contVol
End Function

Public Function PackBoxes(containers As Collection, boxes As Collection) As Collection
    Dim results As New Collection
    Dim remainingBoxes As Collection
    
    ' ������Ｏ�ϣ������޸�ԭʼ���ݣ�
    Set remainingBoxes = CloneBoxCollection(boxes)
    
    ' ����װ��������������ʹ�ô�װ�䣩
    Dim sortedContainers As Collection
    Set sortedContainers = SortContainersByVolume(containers)
    
    Dim cont As CContainer
    For Each cont In sortedContainers
        Do While remainingBoxes.Count > 0
            Dim result As CPackingResult
            Set result = PackSingleContainer(cont, remainingBoxes)
            
            If result.PackedBoxes.Count > 0 Then
                results.Add result
                UpdateRemainingBoxes remainingBoxes, result.PackedBoxes
            Else
                Exit Do
            End If
        Loop
        If remainingBoxes.Count = 0 Then Exit For
    Next
    
    Set PackBoxes = results
End Function

'============================================
' ��������ʵ��
'============================================

' ������Ｏ��
Private Function CloneBoxCollection(src As Collection) As Collection
    Dim dest As New Collection
    Dim box As CBox
    For Each box In src
        dest.Add box
    Next
    Set CloneBoxCollection = dest
End Function

' ����װ���ݻ����򣨴Ӵ�С��
Private Function SortContainersByVolume(containers As Collection) As Collection
    Dim arr() As CContainer
    ReDim arr(1 To containers.Count)
    
    ' ת��Ϊ����
    Dim i As Long, j As Long
    For i = 1 To containers.Count
        Set arr(i) = containers(i)
    Next
    
    ' ��ð������
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i).volume < arr(j).volume Then
                Dim temp As CContainer
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next
    Next
    
    ' ת���ؼ���
    Dim sorted As New Collection
    For i = 1 To UBound(arr)
        sorted.Add arr(i)
    Next
    Set SortContainersByVolume = sorted
End Function

' ����δװ�ػ��Ｏ��
Private Sub UpdateRemainingBoxes(remaining As Collection, packed As Collection)
    Dim packedItem As Object
    For Each packedItem In packed
        Dim targetBox As CBox
        Set targetBox = packedItem("Box")
        
        ' ����������Ҫ�Ƴ��Ļ���
        Dim i As Long
        For i = remaining.Count To 1 Step -1
            If remaining(i).ID = targetBox.ID Then
                remaining.Remove i
                Exit For
            End If
        Next
    Next
End Sub

'============================================
' CContainer�������Volume����
'============================================
' ��CContainer��ģ������ӣ�
Public Property Get volume() As Double
    volume = InnerLength * InnerWidth * InnerHeight
End Property

Sub RunFullProcess()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stuffing")
    
    ' ��������
    Dim containers As Collection
    Dim boxes As Collection
    Set containers = GetSelectedContainers(ws)  ' �ӹ������ȡѡ��ļ�װ��
    Set boxes = ReadBoxDataFromSheet(ws)        ' �ӹ������ȡ��������
    
    ' ִ��װ�����
    Dim results As Collection
    Set results = PackBoxes(containers, boxes)
    
    ' ������
    OutputResultsToSheet ws, results
End Sub

' ��������������ʾ����
Private Sub OutputResultsToSheet(ws As Worksheet, results As Collection)
    ws.Range("K3:R1000").ClearContents
    
    Dim rowOffset As Long: rowOffset = 0
    Dim result As CPackingResult
    For Each result In results
        With ws.Range("K3").Offset(rowOffset)
            .Offset(0, 0).value = result.Container.Name
            .Offset(0, 1).value = result.PackedBoxes.Count
            .Offset(0, 2).value = Format(result.Efficiency, "0.00%")
            .Offset(0, 3).value = result.WeightUsage & "/" & result.Container.MaxLoad
        End With
        rowOffset = rowOffset + 1
    Next
End Sub

Private Sub OutputResults(ws As Worksheet, results As Collection)
    ' [�������߼��������ʵ������ʵ��]
End Sub

