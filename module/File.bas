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

' ***************************************
' * 2次元リスト(行、列で入れ子になったCollection)をファイル出力する
' *
' * Params
' * ------
' * filePath: String
' *     出力ファイルパス
' * valueList: Collection
' *     出力するCollection(行、列で入れ子)
' * delimiter: String (Optional)
' *     値の区切り文字
' ***************************************
Public Function output2DimListToFile(filePath As String, valueList As Collection, Optional delimiter As String = ",")
    Dim row As Collection
    Dim value As Variant
    Dim buffer As String

    Dim fileNum As Integer

    ' ファイルのオープン
    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' ファイルの書き込み
    For Each row In valueList
        buffer = ""
        For Each value In row
            buffer = buffer & CStr(value) & delimiter
        Next value
        ' 末尾のdelimiterを削除
        buffer = Left(buffer, Len(buffer) - Len(delimiter))
        Print #fileNum, buffer
    Next row
    
    ' ファイルのクローズ
    Close #fileNum
End Function
