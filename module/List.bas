Option Explicit

' ***************************************
' * 二つのcollectionリストを結合する
' *
' * Params
' * ------
' * beforeList: Collection
' *     結合対象(前)
' * afterList: Collection
' *     結合対象(後)
' ***************************************
Public Function concatList(beforeList As Collection, afterList As Collection)
    Dim conList As New Collection
    Dim val As Variant

    For Each val In beforeList
        conList.Add (val)
    Next

    For Each val In afterList
        conList.Add (val)
    Next

    Set concatList = conList
End Function
