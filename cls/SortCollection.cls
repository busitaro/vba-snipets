Option Explicit

' �L�[�R���N�V���� �i�L�[�͕������z�肵�A�d�������e����j
Private m_sortKeyList As Collection
' Value�R���N�V����
Private m_valueList As Collection

Public Property Get keys() As Collection
    Set keys = m_sortKeyList
End Property

Public Property Get values() As Collection
    Set values = m_valueList
End Property

Public Property Get Count() As Long
    Count = m_valueList.Count
End Property

' *************************************************
' * �R���X�g���N�^
' *
' *************************************************
Private Sub Class_Initialize()
    Set m_sortKeyList = New Collection
    Set m_valueList = New Collection
End Sub

' *************************************************
' * �l��ǉ�
' * 
' * Params
' * ------
' * value: 
' *     �ǉ�����l
' * key:
' *     �\�[�g�L�[
' *************************************************
Public Function add(value, key)
    m_sortKeyList.add key
    m_valueList.add value
End Function

' *************************************************
' * ���ׂĂ�key�̍Đݒ肷��
' *
' * Params
' * ------
' * keyList: Collection
' *     �Đݒ�̃��X�g
' *
' *************************************************
Public Function resetKey(keyList As Collection)
    Set m_sortKeyList = keyList
End Function

' *************************************************
' * �\�[�g���s��
' * (����ȃ\�[�g�ŁA��r�I�����ȃ}�[�W�\�[�g���̗p)
' *
' * Return
' * ------
' * �\�[�g��̃��X�g
' *************************************************
Public Function sort() As Collection
    Dim keyArray As Variant
    Dim valArray As Variant
    Dim sortTraceArray As Variant
    Dim cnt As Integer
    Dim isValueObject As Boolean

    ' collection����̏ꍇ�I��
    If m_valueList.Count = 0 Then Exit Function
    
    ' value��Object�^���ǂ�������
    If TypeName(m_valueList.Item(1)) = "Object" _
        Or TypeName(m_valueList.Item(1)) = "Collection" Then
        isValueObject = True
    Else
        isValueObject = False
    End If
    
    ' collection����x�z��ɕϊ�
    ReDim keyArray(m_valueList.Count - 1)
    ReDim valArray(m_valueList.Count - 1)
    ReDim sortTraceArray(m_valueList.Count - 1)
    cnt = 0

    For cnt = 1 To m_sortKeyList.Count
        keyArray(cnt - 1) = m_sortKeyList.Item(cnt)
        If isValueObject Then
            Set valArray(cnt - 1) = m_valueList.Item(cnt)
        Else
            valArray(cnt - 1) = m_valueList.Item(cnt)
        End If
        sortTraceArray(cnt - 1) = cnt
    Next

    ' �}�[�W�\�[�g�����{
    Call MergeSort(valArray, keyArray, sortTraceArray, LBound(valArray), UBound(valArray), isValueObject)

    ' �z���collection�Ɋi�[
    Set m_sortKeyList = New Collection
    Set m_valueList = New Collection
    Set sort = New Collection
    For cnt = LBound(valArray) To UBound(valArray)
        m_sortKeyList.add keyArray(cnt)
        m_valueList.add valArray(cnt)
        sort.add sortTraceArray(cnt)
    Next cnt
End Function

' *************************************************
' * �}�[�W�\�[�g�̏���
' *
' * Return
' * ------
' * valArray: Variant
' *     �\�[�g���s���z��
' * keyArray: Variant
' *     ���̔z��
' * sortTraceArray: Variant
' *     �\�[�g�O�̏��Ԃ�ێ�����z��
' * top: Long
' *     �\�[�g�J�n�ʒu
' * bottom: Long
' *     �\�[�g�I���ʒu
' * isValueObject: Boolean
' *     True: �l���I�u�W�F�N�g�^ / False: �l����I�u�W�F�N�g�^
' *************************************************
Private Function MergeSort(valArray As Variant, keyArray As Variant, sortTraceArray As Variant, top As Long, bottom As Long, isValueObject As Boolean)
    Dim middle As Long
    Dim rCnt As Long
    Dim lCnt As Long
    Dim cnt As Integer
    Dim compare As Integer
    Dim cpyValArray As Variant
    Dim cpyKeyArray As Variant
    Dim cpyTraceArray As Variant
    
    ' �����_�̌���
    middle = Int((top + bottom) / 2)
    If top <> bottom Then
        Call MergeSort(valArray, keyArray, sortTraceArray, top, middle, isValueObject)
        Call MergeSort(valArray, keyArray, sortTraceArray, middle + 1, bottom, isValueObject)
    End If

    ' �\�[�g��z��̃}�[�W
    rCnt = top
    lCnt = middle + 1
    cnt = top
    
    cpyKeyArray = keyArray
    cpyValArray = valArray
    cpyTraceArray = sortTraceArray
    Do While (rCnt <= middle) And (lCnt <= bottom)
        ' ��r
        If VarType(cpyKeyArray(rCnt)) = vbLong And VarType(cpyKeyArray(lCnt)) = vbLong Then
            ' ���l�̏ꍇ
            If cpyKeyArray(rCnt) < cpyKeyArray(lCnt) Then
                compare = -1
            Else
                compare = 1
            End If
        Else
            ' ������̏ꍇ
            compare = StrComp(cpyKeyArray(rCnt), cpyKeyArray(lCnt), vbBinaryCompare)
        End If
        
        If compare = -1 Then
            keyArray(cnt) = cpyKeyArray(rCnt)
            If isValueObject Then
                Set valArray(cnt) = cpyValArray(rCnt)
            Else
                valArray(cnt) = cpyValArray(rCnt)
            End If
            sortTraceArray(cnt) = cpyTraceArray(rCnt)
            cnt = cnt + 1
            rCnt = rCnt + 1
        Else
            keyArray(cnt) = cpyKeyArray(lCnt)
            If isValueObject Then
                Set valArray(cnt) = cpyValArray(lCnt)
            Else
                valArray(cnt) = cpyValArray(lCnt)
            End If
            sortTraceArray(cnt) = cpyTraceArray(lCnt)
            cnt = cnt + 1
            lCnt = lCnt + 1
        End If
    Loop
    ' ���z��̎c��
    Do While rCnt <= middle
        keyArray(cnt) = cpyKeyArray(rCnt)
        If isValueObject Then
            Set valArray(cnt) = cpyValArray(rCnt)
        Else
            valArray(cnt) = cpyValArray(rCnt)
        End If
        sortTraceArray(cnt) = cpyTraceArray(rCnt)
        cnt = cnt + 1
        rCnt = rCnt + 1
    Loop
    ' �E�z��̎c��
    Do While lCnt <= middle
        keyArray(cnt) = cpyKeyArray(lCnt)
        If isValueObject Then
            Set valArray(cnt) = cpyValArray(lCnt)
        Else
            valArray(cnt) = cpyValArray(lCnt)
        End If
        sortTraceArray(cnt) = cpyTraceArray(lCnt)
        cnt = cnt + 1
        lCnt = lCnt + 1
    Loop
End Function

' *************************************************
' * SortCollection���}�[�W����
' *
' * Params
' * ------
' * sCollection: �}�[�W�Ώ�Collection
' *
' *************************************************
Public Function merge(sCollection As SortCollection)
    Dim cnt As Long
    
    For cnt = 1 To sCollection.keys.Count
        m_sortKeyList.add sCollection.keys.Item(cnt)
        m_valueList.add sCollection.values.Item(cnt)
    Next cnt
End Function