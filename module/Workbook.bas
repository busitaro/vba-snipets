Option Explicit

' ***************************************
' * エクセルワークブックを開く
' *
' * Params
' * ------
' * filePath: String
' *     オープン対象ファイルへのパス
' *
' * Return
' * ------
' * Workbook:
' *     オープンしたワークブックオブジェクト
' ***************************************
Public Function openWorkbook(filePath As String) As Workbook
    Const errmsg = "ファイルが存在しません。"

    Dim orgWb As Workbook
    Set orgWb = ActiveWorkbook

    If Dir(filePath) = "" Then
        Err.Raise errCode_Stop, errmsg
        Exit Function
    End If

    Set openWorkbook = Workbooks.Open(filePath)
    ' 元々アクティブだったファイルを、再度アクティブにする
    orgWb.Activate
End Function

' ***************************************
' * 新しいシートを作成
' *
' * Params
' * ------
' * sheetName: String
' *     作成するシート名
' * forceFlg: Boolean (Optional)
' *     強制フラグ(True => 既存同名シートが有った場合削除)
' *
' * Return
' * ------
' * Worksheet:
' *     作成したシートへのオブジェクト
' ***************************************
Function makeNewSheet(sheetName As String, Optional forceFlg As Boolean = False) As Worksheet
    Dim orgSt As Worksheet      ' 元々アクティブだったワークシート
    Dim retSt As Worksheet      ' 追加したワークシート

    Set orgSt = ActiveSheet

    If isExistsSheet(sheetName) And forceFlg Then
        Application.DisplayAlerts = False
        Call Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add
    ActiveSheet.Name = sheetName
    Set retSt = ActiveSheet
    orgSt.Activate

    Set makeNewSheet = retSt
End Function

' ***************************************
' * シートの存在確認
' *
' * Params
' * ------
' * sheetName: String
' *     チェック対象シート名
' *
' * Return
' * ------
' * Boolean:
' *     True => 存在する / False => 存在しない
' ***************************************
Function isExistsSheet(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim isExists As Boolean

    isExists = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            isExists = True
            Exit For
        End If
    Next ws

    isExistsSheet = isExists
End Function

' ***************************************
' * 格子状の罫線を引く
' *
' * Params
' * ------
' * rgTopLeft: Range
' *     罫線を引く左上のセル
' * rgBottomRight: Range
' *     罫線を引く右下のセル
' ***************************************
Function drawGridLine(rgTopLeft As Range, rgBottomRight As Range)
    With Range(rgTopLeft, rgBottomRight)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Function
