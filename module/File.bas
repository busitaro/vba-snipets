Option Explicit

' ***************************************
' * 指定フォルダのファイル一覧を子フォルダ含め再帰的に取得する
' *
' * Params
' * ------
' * path: String
' *     対象フォルダ
' * fullPath: Boolean (Optional)
' *     フルパスで取得するか、ファイル名のみか
' *
' * Return
' * ------
' * Collection:
' *     取得結果のリスト
' ***************************************
Public Function getFile(path As String, Optional fullPath As Boolean = True) As Collection
    ' ファイルリスト
    Dim fileList As New Collection

    Dim f, d As Object

    With CreateObject("Scripting.FileSystemObject").GetFolder(path)
        ' 通常のファイル
        For Each f In .Files
            If fullPath Then
                fileList.Add (f.path)
            Else
                fileList.Add (f.Name)
            End If
        Next f

        ' フォルダを再帰的に検索
        For Each d In .SubFolders
            Set fileList = concatList(fileList, getFile(CStr(d), fullPath))
        Next d
    End With
    
    Set getFile = fileList
End Function
