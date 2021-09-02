Option Explicit

' ***************************************
' * 文字列中のsearchRegStrをreplaceStrに置換する
' * searchRegStrには正規表現を指定可能
' * (WorkSheet FunctionのSUBSTITUTEの正規表現版)
' *
' * Params
' * ------
' * searchRegStr: String
' *     検索対象文字列(正規表現)
' * replaceStr: String
' *     置換文字列
' * index: Integer
' *     置換対象を絞る場合、前から何番目を対象とするか指定(省略した場合すべて置換)
' *
' * Memo
' * ------
' * 性能を考慮し、正規表現での検索を一度にする為、再帰関数としない
' ***************************************
Public Function SUBSTITUTE_REG(targetStr As String, _
                                searchRegStr As String, _
                                replaceStr As String, _
                                Optional index As Integer = 0) As String

    ' パラメータ バリデーション
    If index < 0 Then
        SUBSTITUTE_REG = CVErr(xlErrValue)
        Exit Function
    End If

    ' 正規表現オブジェクトの設定
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = searchRegStr
        .Global = True
    End With

    ' 正規表現マッチ箇所をリストへ格納
    ' MEMO: re.Executeの結果順序が補償されている必要がある
    Dim match As Variant        ' IMatch2
    Dim replaceIndexList As New Collection

    For Each match In re.Execute(targetStr)
        Call replaceIndexList.Add( _
            new_Collection(match.FirstIndex, match.Length) _
        )
    Next match

    ' 置換対象を絞込
    If index > replaceIndexList.count Then
        SUBSTITUTE_REG = CVErr(xlErrValue)
        Exit Function
    End If
    If index > 0 Then
        Set replaceIndexList = new_Collection(replaceIndexList.Item(index))
    End If

    ' 置換の実施
    Dim replace As Collection
    Dim idx As Integer
    For idx = replaceIndexList.count To 1 Step -1
        Set replace = replaceIndexList.Item(idx)
        targetStr = _
            Mid(targetStr, 1, replace.Item(1)) & _
            replaceStr & _
            Mid(targetStr, replace.Item(1) + replace.Item(2) + 1)
    Next idx

    SUBSTITUTE_REG = targetStr
End Function
