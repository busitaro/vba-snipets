Option Explicit

' クォート文字
Private m_quote As String
' ヘッダ行を持つかのフラグ
Private m_hasHeader As Boolean
' ファイル名
Private m_fileNum As Integer
' 行の末尾かのフラグ
Private m_endLine As Boolean
' ファイルの末尾まで読み終わったかのフラグ
Private m_endOfFile As Boolean
' ファイルを開いているかのフラグ
Private m_fileIsOpen As Boolean

' *************************************************
' * コンストラクタ
' *
' *************************************************
Private Sub Class_Initialize()
    m_fileIsOpen = False
End Sub

' *************************************************
' * デストラクタ
' *
' *************************************************
Private Sub Class_Terminate()
    CloseFile
End Sub

' *************************************************
' * ファイル内容を読込む
' *
' * Params
' * ------
' * filename: String
' *     読込み対象ファイルのパス
' * hasHeader: Boolean
' *     1行目をヘッダとして扱うかのフラグ
' * quote: String
' *     囲み文字
' *
' * Return
' * ------
' * ファイル内容: Collection
' *     各要素はヘッダもしくは連番をキーにした辞書オブジェクト
' *
' *************************************************
Public Function readFile(filename As String, Optional hasHeader As Boolean = False, Optional quote As String = """") As Collection
    ' メンバ設定
    m_quote = quote
    m_hasHeader = hasHeader

    ' ファイルオープン
    Call openFile(filename)

    ' 内容を読込
    Set readFile = readAllValues()
End Function

' *************************************************
' * ファイルをオープンする
' *
' * Params
' * ------
' * filename: String
' *     読込み対象ファイルのパス
' *
' *************************************************
Private Function openFile(filename As String)
    m_fileNum = FreeFile
    Open filename For Input Access Read As m_fileNum
    m_fileIsOpen = True
End Function

' *************************************************
' * ファイルをクローズする
' *
' * Params
' * ------
' * filename: String
' *     読込み対象ファイルのパス
' *
' *************************************************
Function CloseFile()
    If m_fileIsOpen Then
        Close m_fileNum
        m_fileIsOpen = False
    End If
End Function

' *************************************************
' * ファイルの全ての内容を読み込む
' *
' * Params
' * ------
' * filename: String
' *     読込み対象ファイルのパス
' *
' * Return
' * ------
' * ファイル内容: Collection
' *
' *************************************************
Private Function readAllValues() As Collection
    Set readAllValues = New Collection
    Dim values As Object
    Dim lineValues As Collection
    Dim columns As Collection
    Dim idx As Integer
    
    If m_hasHeader Then
        Set columns = readLine()
    End If

    Do Until EOF(m_fileNum)
        ' 1行ずつ読込み、辞書にする
        Set values = CreateObject("Scripting.Dictionary")
        Set lineValues = readLine()

        If m_hasHeader Then
            ' ヘッダがある場合、ヘッダの各項目名をキーとした辞書とする
            For idx = 1 To columns.Count
                Call values.Add(columns.item(idx), lineValues.item(idx))
            Next
        Else
            ' ヘッダがない場合、連番をキーとする
            For idx = 1 To lineValues.Count
                Call values.Add(idx, lineValues.item(idx))
            Next
        End If
        Call readAllValues.Add(values)
    Loop
End Function

' *************************************************
' * ファイルから1行読込む
' *
' * Return
' * ------
' * 行内容: Collection
' *
' *************************************************
Private Function readLine() As Collection
    Set readLine = New Collection
    m_endLine = False

    Do Until m_endLine
        readLine.Add (readValue())
    Loop
End Function

' *************************************************
' * ファイルから1項目読込む
' *
' * Return
' * ------
' * 項目内容: String
' *
' *************************************************
Private Function readValue() As String
    Dim buffer As String
    buffer = ""

    Do Until isCloseValue(buffer)
        buffer = buffer & Input(1, m_fileNum)
    Loop

    readValue = extractValue(buffer)
End Function

' *************************************************
' * 項目の終わりまで読込まれているか判定する
' *
' * Return
' * ------
' * 読込まれているかのフラグ: Boolean
' *
' *************************************************
Private Function isCloseValue(buffer) As Boolean
    Select Case Right(buffer, 1)
        Case ",", vbLf
            isCloseValue = True
        Case Else
            isCloseValue = False
    End Select
End Function

' *************************************************
' * 文字列から項目値部分のみを抽出する
' *
' * Return
' * ------
' * 抽出した値: String
' *
' *************************************************
Private Function extractValue(target As String) As String
    If Left(target, 1) = m_quote Then
        ' クォート囲み有り
        If Right(target, 2) = m_quote & "," Then
            extractValue = Mid(target, 2, Len(target) - 3)
        ElseIf Right(target, 2) = m_quote & vbLf Then
            extractValue = Mid(target, 2, Len(target) - 3)
            m_endLine = True
        ElseIf Right(target, 3) = m_quote & vbCr & vbLf Then
            extractValue = Mid(target, 2, Len(target) - 4)
            m_endLine = True
        Else
            Err.Raise 9999, Description:="ファイルフォーマット不正"
        End If
    Else
        ' クォート囲み無し
        If Right(target, 1) = "," Then
            extractValue = Mid(target, 1, Len(target) - 1)
        ElseIf Right(target, 1) = vbLf Then
            extractValue = Mid(target, 1, Len(target) - 1)
            m_endLine = True
        ElseIf Right(target, 2) = vbCr & vbLf Then
            extractValue = Mid(target, 1, Len(target) - 2)
            m_endLine = True
        Else
            Err.Raise 9999, Description:="ファイルフォーマット不正"
        End If
    End If
End Function
