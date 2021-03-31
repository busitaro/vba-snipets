Option Explicit

' キーコレクション （キーは文字列を想定し、重複も許容する）
Private m_sortKeyList As Collection
' Valueコレクション
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
' * コンストラクタ
' *
' *************************************************
Private Sub Class_Initialize()
    Set m_sortKeyList = New Collection
    Set m_valueList = New Collection
End Sub

' *************************************************
' * 値を追加
' * 
' * Params
' * ------
' * value: 
' *     追加する値
' * key:
' *     ソートキー
' *************************************************
Public Function add(value, key)
    m_sortKeyList.add key
    m_valueList.add value
End Function

' *************************************************
' * すべてのkeyの再設定する
' *
' * Params
' * ------
' * keyList: Collection
' *     再設定のリスト
' *
' *************************************************
Public Function resetKey(keyList As Collection)
    Set m_sortKeyList = keyList
End Function

' *************************************************
' * ソートを行う
' * (安定なソートで、比較的高速なマージソートを採用)
' *
' * Return
' * ------
' * ソート後のリスト
' *************************************************
Public Function sort() As Collection
    Dim keyArray As Variant
    Dim valArray As Variant
    Dim sortTraceArray As Variant
    Dim cnt As Integer
    Dim isValueObject As Boolean

    ' collectionが空の場合終了
    If m_valueList.Count = 0 Then Exit Function
    
    ' valueがObject型かどうか判定
    If TypeName(m_valueList.Item(1)) = "Object" _
        Or TypeName(m_valueList.Item(1)) = "Collection" Then
        isValueObject = True
    Else
        isValueObject = False
    End If
    
    ' collectionを一度配列に変換
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

    ' マージソートを実施
    Call MergeSort(valArray, keyArray, sortTraceArray, LBound(valArray), UBound(valArray), isValueObject)

    ' 配列をcollectionに格納
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
' * マージソートの処理
' *
' * Return
' * ------
' * valArray: Variant
' *     ソートを行う配列
' * keyArray: Variant
' *     鍵の配列
' * sortTraceArray: Variant
' *     ソート前の順番を保持する配列
' * top: Long
' *     ソート開始位置
' * bottom: Long
' *     ソート終了位置
' * isValueObject: Boolean
' *     True: 値がオブジェクト型 / False: 値が非オブジェクト型
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
    
    ' 分割点の決定
    middle = Int((top + bottom) / 2)
    If top <> bottom Then
        Call MergeSort(valArray, keyArray, sortTraceArray, top, middle, isValueObject)
        Call MergeSort(valArray, keyArray, sortTraceArray, middle + 1, bottom, isValueObject)
    End If

    ' ソート後配列のマージ
    rCnt = top
    lCnt = middle + 1
    cnt = top
    
    cpyKeyArray = keyArray
    cpyValArray = valArray
    cpyTraceArray = sortTraceArray
    Do While (rCnt <= middle) And (lCnt <= bottom)
        ' 比較
        If VarType(cpyKeyArray(rCnt)) = vbLong And VarType(cpyKeyArray(lCnt)) = vbLong Then
            ' 数値の場合
            If cpyKeyArray(rCnt) < cpyKeyArray(lCnt) Then
                compare = -1
            Else
                compare = 1
            End If
        Else
            ' 文字列の場合
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
    ' 左配列の残り
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
    ' 右配列の残り
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
' * SortCollectionをマージする
' *
' * Params
' * ------
' * sCollection: マージ対象Collection
' *
' *************************************************
Public Function merge(sCollection As SortCollection)
    Dim cnt As Long
    
    For cnt = 1 To sCollection.keys.Count
        m_sortKeyList.add sCollection.keys.Item(cnt)
        m_valueList.add sCollection.values.Item(cnt)
    Next cnt
End Function
