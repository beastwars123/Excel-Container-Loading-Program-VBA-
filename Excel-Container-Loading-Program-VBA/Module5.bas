Attribute VB_Name = "Module5"
Option Explicit
Private Const TABLE_CTNR_USE As String = "CTNR_Use"
Private Const TABLE_CARGO_SPEC As String = "Cargo_Spec"

' ��װ���㷨��FFD + ��ά�ֲ�������
Public Sub PackContainersFFD()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Stuffing")
    Dim containers As Collection: Set containers = GetSelectedContainers(ws)
    Dim boxes As Collection: Set boxes = ReadBoxDataFromSheet(ws)
    
    ' �������������FFD���Ĳ��裩
    Dim sortedBoxes As Collection: Set sortedBoxes = SortBoxes(boxes, VolumeDesc)
    
    Dim containerLoads As New Collection ' �洢���м�װ���װ�ؽ��
    
    For Each ctnr In containers
        Dim load As New CContainerLoad
        Set load.Container = ctnr
        Dim layers As New Collection ' �洢�ü�װ������в�
        
        ' ��ʼ����һ�㣺z=0����������Ϊ��װ�����ߴ�
        Dim firstLayer As New CLayer
        firstLayer.z = 0
        firstLayer.AddAvailableArea 0, 0, ctnr.InnerWidth, ctnr.InnerHeight
        layers.Add firstLayer
        
        ' ����ÿ������ʵ����չ��Quantity��
        Dim box As CBox
        For Each box In sortedBoxes
            Dim placed As Boolean: placed = False
            Dim currentLayer As CLayer
            
            ' �״���Ӧ�����������в����ҵ����ʵ�����
            For Each currentLayer In layers
                If PlaceBoxInLayer(box, currentLayer, load) Then
                    placed = True
                    Exit For
                End If
            Next
            
            ' �����в��޷����ã���δ������װ��߶ȣ������²�
            If Not placed And (load.UsedHeight + box.height <= ctnr.InnerHeight) Then
                Dim newLayer As New CLayer
                newLayer.z = load.UsedHeight
                newLayer.AddAvailableArea 0, 0, ctnr.InnerWidth, ctnr.InnerHeight
                layers.Add newLayer
                If PlaceBoxInLayer(box, newLayer, load) Then
                    placed = True
                End If
            End If
            
            ' ��¼δ���õĻ����ѡ��������
            If Not placed Then Debug.Print "���� " & box.ID & " �޷�װ�뼯װ�� " & ctnr.Name
        Next box
        
        containerLoads.Add load
    Next ctnr
    
    ' �����������ӻ�ģ��ʹ�ã�ȫ�ֱ����������ԣ�
    ThisWorkbook.Names.Add "GlobalLoadedContainers", containerLoads
End Sub

' �ڲ��з��û��ﲢ�����Ƿ�ɹ�
Private Function PlaceBoxInLayer(box As CBox, layer As CLayer, load As CContainerLoad) As Boolean
    Dim area As Variant
    For Each area In layer.availableAreas
        ' �������Ƿ��ܷ��뵱ǰ����x,y����ߴ�ƥ�䣬z���򲻳�����װ�䳤�ȣ�
        If box.width <= area(2) And box.height <= area(3) And (layer.z + box.length <= load.Container.InnerLength) Then
            ' ���û���ָ�ʣ������
            PlaceBoxInArea box, area, layer, load
            PlaceBoxInLayer = True
            Exit Function
        End If
    Next
End Function

' �������з��û��ﲢ�ָ�ʣ��ռ䣨������䣩
Private Sub PlaceBoxInArea(box As CBox, area As Variant, layer As CLayer, load As CContainerLoad)
    Dim x As Double: x = area(0)
    Dim y As Double: y = area(1)
    Dim areaWidth As Double: areaWidth = area(2)
    Dim areaHeight As Double: areaHeight = area(3)
    
    ' ��¼����λ�úͷ��򣨼�Ϊԭʼ���򣬿���չ��ת�߼���
    Dim position As New Dictionary
    position.Add "ID", box.ID
    position.Add "x", x
    position.Add "y", y
    position.Add "z", layer.z
    position.Add "length", box.length
    position.Add "width", box.width
    position.Add "height", box.height
    
    ' ��ӵ�װ�ؽ��
    load.LoadedBoxes.Add position
    load.UsedHeight = Max(load.UsedHeight, layer.z + box.length) ' z��Ϊ���ȷ��򣨼�װ����ȣ�
    load.UsedWeight = load.UsedWeight + box.weight
    
    ' �ָ�ʣ�����������Ҳ���Ϸ��Ŀ�������
    layer.availableAreas.Remove (layer.availableAreas.IndexOf(area)) ' �Ƴ���ǰʹ�õ�����
    
    ' �Ҳ�����x + ������ �� ������
    If box.width < areaWidth Then
        layer.AddAvailableArea x + box.width, y, areaWidth - box.width, areaHeight
    End
    
    ' �Ϸ�����y + ����߶� �� ����߶�
    If box.height < areaHeight Then
        layer.AddAvailableArea x, y + box.height, areaWidth, areaHeight - box.height
    End
End Sub

' ������������ȡ���ֵ������Double���ͣ�
Private Function Max(a As Double, b As Double) As Double
    If a >= b Then Max = a Else Max = b
End Function

' �������к�����GetSelectedContainers, ReadBoxDataFromSheet, SortBoxes �ȱ��ֲ��䣩
' ȷ��SortBoxes������ȷ���ذ������������Ļ��Ｏ��

Sub TestRealDataPacking()
    On Error GoTo TestErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Stuffing") ' ����������"Stuffing"������
    
    ' ��֤������
    If Not DoesTableExist(ws, TABLE_CARGO_SPEC) Or _
       Not DoesTableExist(ws, TABLE_CTNR_USE) Then
        MsgBox "�����ڱ����������Cargo_Spec���ͼ�װ�䣨CTNR_Use������", vbExclamation
        Exit Sub
    End If
    
    ' ִ��װ���㷨
    PackContainersFFD
    
    ' ��ȡȫ��װ�ؽ������������ģ��������ȷ�洢��
    Dim globalLoads As Collection
    Set globalLoads = ThisWorkbook.Names("GlobalLoadedContainers").RefersToRange.value
    
    ' ����������������
    Debug.Print "======================"
    Debug.Print "= ʵ������װ�������� ="
    Debug.Print "======================"
    Debug.Print "��ȡ������ݳɹ�����ʼ��֤װ����..."
    
    Dim load As CContainerLoad
    For Each load In globalLoads
        With load.Container
            Debug.Print vbNewLine & "����װ����Ϣ��"
            Debug.Print "�ͺ�: " & .Name
            Debug.Print "�ڲ��ߴ� (LxWxH): " & .InnerLength & " x " & .InnerWidth & " x " & .InnerHeight & " cm"
            Debug.Print "�������: " & .MaxLoad & " kg"
        End With
        
        Debug.Print "��װ��ժҪ��"
        Debug.Print "���ø߶�: " & load.UsedHeight & " cm (" & Format(load.UsedHeight / load.Container.InnerHeight, "0.00%") & ")"
        Debug.Print "��������: " & load.UsedWeight & " kg (" & Format(load.UsedWeight / load.Container.MaxLoad, "0.00%") & ")"
        Debug.Print "װ�ػ�������: " & load.LoadedBoxes.Count & " ����ԭʼQuantityչ����"
        
        ' ���ÿ�������װ��λ�ã�ǰ5��ʾ�����������������
        Debug.Print vbNewLine & "��ǰ5��װ�ػ������顿"
        Dim i As Long
        Dim position As New Dictionary
        For i = 1 To Min(5, load.LoadedBoxes.Count)
            Set position = load.LoadedBoxes(i)
            Debug.Print "����ID: " & position("ID")
            Debug.Print "�ڷ�λ�� (X,Y,Z): " & position("x") & " x " & position("y") & " x " & position("z")
            Debug.Print "ռ�óߴ� (LxWxH): " & position("length") & " x " & position("width") & " x " & position("height")
            Debug.Print "------------------------"
        Next i
        
        ' ��֤�ؼ�Լ�����ɸ���ҵ����Ӹ�����֤��
        If load.UsedWeight > load.Container.MaxLoad Then
            Debug.Print "?? ���棺���س�����װ��������أ�"
        End If
        If load.UsedHeight > load.Container.InnerHeight Then
            Debug.Print "?? ���棺ʹ�ø߶ȳ�����װ���ڲ��߶ȣ�"
        End If
    Next load
    
    Exit Sub

TestErrorHandler:
    Debug.Print "���Թ����г��ִ���" & Err.Description
    Resume Next
End Sub

' ������������ȡ��Сֵ�������������������
Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

Private Function DoesTableExist(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    Dim dummy As ListObject: Set dummy = ws.ListObjects(tableName)
    DoesTableExist = Not dummy Is Nothing
End Function
