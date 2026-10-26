Attribute VB_Name = "modListDiff"
Option Explicit

'==============================================================
' ListDiff - 2つの表（CSV）の差分比較マクロ
'
' 「変更前」「変更後」2つのCSVファイルを比較し、
' ・変更後にだけある行（追加）
' ・変更前にだけある行（削除）
' ・両方にあるが内容が変わっている行（変更）
' を自動で判定し、一覧シートにまとめます。
'
' 作者: 天電輝合同会社 (TENDENKO)
' ライセンス: MIT License
'
' ---- 使い方 ----
' 1. 「変更前ファイルを開く」で、比較元のCSVファイルを選ぶ
' 2. 「変更後ファイルを開く」で、比較先のCSVファイルを選ぶ
' 3. 「比較実行」を押す
' 4. 「差分一覧」シートに、追加・削除・変更が一覧表示される
'
' 一度選んだファイルはブック内部に記録されるので、同じファイル
' のまま何度も比較し直したい場合は、選び直す必要はありません
' （比較実行だけで再実行できます）。
'
' ---- 比較の考え方 ----
' 先頭列（1列目）を「キー」とみなし、同じキーを持つ行同士を
' 比較します。キー以外の列の内容が1つでも違えば「変更」と
' 判定します。
'
' ---- カスタマイズポイント（下の設定セクションを書き換えるだけ） ----
' ・INPUT_ENCODING    : 読み込むCSVの文字コード
' ・HAS_HEADER        : CSVの1行目が見出し行かどうか
' ・DELIMITER         : CSVの区切り文字
' ・OUTPUT_SHEET_NAME : 差分を書き出すシート名
' ・KEY_COLUMN_INDEX  : 「同じ行」と判定するためのキー列（1=先頭列）
'
' ---- 注意事項 ----
' ・2つのファイルは、同じ列構成（列の並び・意味）であることを
'   前提にしています。
' ・同じファイル内にキーが重複する行がある場合、最初の行のみを
'   比較対象とします。
'==============================================================

'---------------- 設定セクション（ここだけ書き換えればOK） ----------------

' 読み込むCSVファイルの文字コード（既定：Shift-JIS）
Private Const INPUT_ENCODING As String = "Shift-JIS"

' CSVの1行目が見出し行かどうか
Private Const HAS_HEADER As Boolean = True

' CSVの区切り文字
Private Const DELIMITER As String = ","

' 差分を書き出すシート名
Private Const OUTPUT_SHEET_NAME As String = "差分一覧"

' 「同じ行」と判定するためのキー列（1=先頭列）
Private Const KEY_COLUMN_INDEX As Long = 1

' 変更された行で「変更前→変更後」を表すときの区切り記号
Private Const DIFF_MARK As String = " → "

' 選択したファイルのパスを記録しておく、ブック内部のプロパティ名
Private Const PROP_NAME_OLD As String = "ListDiff_変更前ファイル"
Private Const PROP_NAME_NEW As String = "ListDiff_変更後ファイル"

'---------------------------------------------------------------------

'==============================================================
' Main_変更前ファイルを開く
'   比較元（変更前）のCSVファイルを選び、ブック内部に記録する。
'==============================================================
Sub Main_変更前ファイルを開く()

    Dim filePath As String
    filePath = PickFile("変更前のCSVファイルを選択してください")

    If filePath = "" Then
        MsgBox "ファイルが選択されませんでした。", vbInformation
        Exit Sub
    End If

    SetSettingValue PROP_NAME_OLD, filePath

    MsgBox "変更前ファイルを設定しました。" & vbCrLf & filePath, vbInformation, "完了"

End Sub

'==============================================================
' Main_変更後ファイルを開く
'   比較先（変更後）のCSVファイルを選び、ブック内部に記録する。
'==============================================================
Sub Main_変更後ファイルを開く()

    Dim filePath As String
    filePath = PickFile("変更後のCSVファイルを選択してください")

    If filePath = "" Then
        MsgBox "ファイルが選択されませんでした。", vbInformation
        Exit Sub
    End If

    SetSettingValue PROP_NAME_NEW, filePath

    MsgBox "変更後ファイルを設定しました。" & vbCrLf & filePath, vbInformation, "完了"

End Sub

'==============================================================
' Main_比較実行
'   記録された変更前・変更後ファイルを比較する。
'==============================================================
Sub Main_比較実行()

    On Error GoTo ErrHandler

    Dim oldFile As String, newFile As String
    oldFile = GetSettingValue(PROP_NAME_OLD)
    newFile = GetSettingValue(PROP_NAME_NEW)

    If oldFile = "" Then
        MsgBox "先に「変更前ファイルを開く」で、比較元のファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    If newFile = "" Then
        MsgBox "先に「変更後ファイルを開く」で、比較先のファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    If Dir(oldFile) = "" Then
        MsgBox "変更前ファイルが見つかりません。ファイルが移動・削除されていないか確認し、" & vbCrLf & _
               "もう一度「変更前ファイルを開く」で選び直してください。" & vbCrLf & vbCrLf & oldFile, vbExclamation
        Exit Sub
    End If

    If Dir(newFile) = "" Then
        MsgBox "変更後ファイルが見つかりません。ファイルが移動・削除されていないか確認し、" & vbCrLf & _
               "もう一度「変更後ファイルを開く」で選び直してください。" & vbCrLf & vbCrLf & newFile, vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "ファイルを読み込んでいます..."
    DoEvents

    ' --- 読み込み ---
    Dim oldLines() As String, newLines() As String
    Dim oldLineCount As Long, newLineCount As Long
    ReadCSVLines oldFile, oldLines, oldLineCount
    ReadCSVLines newFile, newLines, newLineCount

    Dim startRow As Long
    startRow = IIf(HAS_HEADER, 1, 0)

    If oldLineCount <= startRow Or newLineCount <= startRow Then
        Application.StatusBar = False

        Dim noDataMsg As String
        If oldLineCount <= startRow And newLineCount <= startRow Then
            noDataMsg = "変更前・変更後のどちらのファイルにも、比較対象となるデータ行がありませんでした。"
        ElseIf oldLineCount <= startRow Then
            noDataMsg = "変更前ファイルに、比較対象となるデータ行がありませんでした。"
        Else
            noDataMsg = "変更後ファイルに、比較対象となるデータ行がありませんでした。"
        End If

        WriteNoDataSheet noDataMsg, oldFile, newFile

        MsgBox noDataMsg & vbCrLf & vbCrLf & "詳細は「" & OUTPUT_SHEET_NAME & "」シートをご確認ください。", vbExclamation
        Exit Sub
    End If

    ' --- 見出し行の決定 ---
    Dim headers() As String
    If HAS_HEADER Then
        headers = SplitCSVLine(oldLines(0))
    Else
        Dim sampleCols() As String
        sampleCols = SplitCSVLine(oldLines(0))
        ReDim headers(LBound(sampleCols) To UBound(sampleCols))
        Dim hi As Long
        For hi = LBound(sampleCols) To UBound(sampleCols)
            headers(hi) = "列" & (hi + 1)
        Next hi
    End If

    Application.StatusBar = "差分を比較しています..."
    DoEvents

    ' --- Dictionary化：キー → 列データ配列 ---
    Dim oldMap As Object, newMap As Object
    Set oldMap = CreateObject("Scripting.Dictionary")
    Set newMap = CreateObject("Scripting.Dictionary")

    Dim oldSkipped As Long, newSkipped As Long
    oldSkipped = 0
    newSkipped = 0

    Dim oldDupKeys As Object, newDupKeys As Object
    Set oldDupKeys = CreateObject("Scripting.Dictionary")
    Set newDupKeys = CreateObject("Scripting.Dictionary")

    BuildRowMap oldLines, oldLineCount, startRow, oldMap, oldSkipped, oldDupKeys
    BuildRowMap newLines, newLineCount, startRow, newMap, newSkipped, newDupKeys

    ' --- 差分判定 ---
    Dim resultRows As New Collection   ' 各要素: Array(判定, 列値の配列)
    Dim addedCount As Long, deletedCount As Long, changedCount As Long
    addedCount = 0
    deletedCount = 0
    changedCount = 0

    Dim k As Variant
    For Each k In oldMap.Keys
        If Not newMap.Exists(k) Then
            resultRows.Add Array("削除", oldMap(k))
            deletedCount = deletedCount + 1
        End If
    Next k

    Dim oldCols() As String, newCols() As String
    Dim mergedCols() As String
    Dim isChanged As Boolean
    Dim ci As Long

    For Each k In newMap.Keys
        If Not oldMap.Exists(k) Then
            resultRows.Add Array("追加", newMap(k))
            addedCount = addedCount + 1
        Else
            oldCols = oldMap(k)
            newCols = newMap(k)
            isChanged = False

            ReDim mergedCols(LBound(newCols) To UBound(newCols))
            For ci = LBound(newCols) To UBound(newCols)
                If ci >= LBound(oldCols) And ci <= UBound(oldCols) Then
                    If oldCols(ci) <> newCols(ci) Then
                        mergedCols(ci) = oldCols(ci) & DIFF_MARK & newCols(ci)
                        isChanged = True
                    Else
                        mergedCols(ci) = newCols(ci)
                    End If
                Else
                    mergedCols(ci) = newCols(ci)
                    isChanged = True
                End If
            Next ci

            If isChanged Then
                resultRows.Add Array("変更", mergedCols)
                changedCount = changedCount + 1
            End If
        End If
    Next k

    ' --- 書き出し ---
    WriteResultSheet headers, resultRows, addedCount, deletedCount, changedCount, oldFile, newFile, oldDupKeys, newDupKeys

    Application.StatusBar = False

    ' --- 完了メッセージ ---
    Dim msg As String
    msg = "比較が完了しました。" & vbCrLf & _
          "追加: " & addedCount & " 件 / 削除: " & deletedCount & " 件 / 変更: " & changedCount & " 件" & vbCrLf & vbCrLf & _
          "結果は「" & OUTPUT_SHEET_NAME & "」シートに出力しました。"

    If oldSkipped > 0 Or newSkipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "【ご注意】列数が不足しているなどの理由で、" & (oldSkipped + newSkipped) & " 件の行をスキップしました。"
    End If

    If oldDupKeys.Count > 0 Or newDupKeys.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "【ご注意】同じキーが複数行ある箇所があったため、最初の行のみを比較に使用しました。詳細はシートをご確認ください。"
    End If

    MsgBox msg, vbInformation, "完了"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================
' Main_結果シート削除
'   「差分一覧」シートを削除する。
'==============================================================
Sub Main_結果シート削除()

    If Not SheetExists(OUTPUT_SHEET_NAME) Then
        MsgBox "「" & OUTPUT_SHEET_NAME & "」シートは存在しません。", vbInformation, "削除不要"
        Exit Sub
    End If

    If ThisWorkbook.Worksheets.Count - 1 < 1 Then
        MsgBox "削除すると全てのシートがなくなってしまうため中止しました。" & vbCrLf & _
               "操作用のシートなどを別に1つ残してから実行してください。", vbCritical, "中止"
        Exit Sub
    End If

    If MsgBox("「" & OUTPUT_SHEET_NAME & "」シートを削除します。" & vbCrLf & _
              "元に戻せません。よろしいですか？", vbYesNo + vbExclamation, "確認") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME).Delete
    Application.DisplayAlerts = True

    MsgBox "削除が完了しました。", vbInformation, "完了"

End Sub

'==============================================================
' 選択したファイルのパスを、ブック内部のドキュメントプロパティに
' 記録する（シートを汚さずに保持するため）
'==============================================================
Private Sub SetSettingValue(ByVal propName As String, ByVal value As String)

    Dim prop As Object
    On Error Resume Next
    Set prop = ThisWorkbook.CustomDocumentProperties(propName)
    On Error GoTo 0

    If prop Is Nothing Then
        ThisWorkbook.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, Type:=4, Value:=value
    Else
        prop.Value = value
    End If

End Sub

'==============================================================
' ブック内部のドキュメントプロパティから、記録済みのファイルパスを
' 取得する。未設定の場合は空文字を返す
'==============================================================
Private Function GetSettingValue(ByVal propName As String) As String

    Dim prop As Object
    On Error Resume Next
    Set prop = ThisWorkbook.CustomDocumentProperties(propName)
    On Error GoTo 0

    If prop Is Nothing Then
        GetSettingValue = ""
    Else
        GetSettingValue = CStr(prop.Value)
    End If

End Function

'==============================================================
' 指定行範囲を読み、キー→列データ配列のDictionaryを作る
'   同じキーが複数回出てきた場合は、最初の行のみを採用し、
'   そのキーが何行目に出現したか（行番号一覧）を dupKeyRows に記録する
'   行番号は、ファイル内の行番号（1行目から数える。HAS_HEADER=True の
'   場合は見出し行を1行目として数える）
'==============================================================
Private Sub BuildRowMap(ByRef lines() As String, ByVal lineCount As Long, ByVal startRow As Long, ByRef targetMap As Object, ByRef skippedCount As Long, ByRef dupKeyRows As Object)

    Dim occurrenceRows As Object
    Set occurrenceRows = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim cols() As String
    Dim key As String
    Dim rowList As Collection

    For r = startRow To lineCount - 1
        If Trim(lines(r)) <> "" Then
            cols = SplitCSVLine(lines(r))

            If KEY_COLUMN_INDEX < LBound(cols) + 1 Or KEY_COLUMN_INDEX > UBound(cols) + 1 Then
                skippedCount = skippedCount + 1
            Else
                key = cols(KEY_COLUMN_INDEX - 1)

                If occurrenceRows.Exists(key) Then
                    occurrenceRows(key).Add (r + 1)
                Else
                    Set rowList = New Collection
                    rowList.Add (r + 1)
                    occurrenceRows.Add key, rowList
                End If

                If Not targetMap.Exists(key) Then
                    targetMap.Add key, cols
                End If
            End If
        End If
    Next r

    ' 2行以上出現したキーだけを、行番号一覧つきで dupKeyRows に転記する
    Dim k As Variant
    For Each k In occurrenceRows.Keys
        If occurrenceRows(k).Count > 1 Then
            dupKeyRows.Add k, occurrenceRows(k)
        End If
    Next k

End Sub

'==============================================================
' 指定シートが存在するかどうかを判定する
'==============================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

'==============================================================
' ファイル選択ダイアログを表示し、選択されたパスを返す
'==============================================================
Private Function PickFile(ByVal dialogTitle As String) As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.Title = dialogTitle
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "CSVファイル", "*.csv"

    If ThisWorkbook.Path <> "" Then
        fd.InitialFileName = ThisWorkbook.Path & "\"
    End If

    If fd.Show = -1 Then
        PickFile = fd.SelectedItems(1)
    Else
        PickFile = ""
    End If

End Function

'==============================================================
' CSVファイルを文字コードを考慮して読み込み、行の配列を返す
'==============================================================
Private Sub ReadCSVLines(ByVal filePath As String, ByRef lines() As String, ByRef lineCount As Long)

    ' 完全に空（0バイト）のファイルは、ADODB.Streamで読み込もうとすると
    ' 環境によってはエラーになることがあるため、ファイルサイズの時点で
    ' 判定し、Streamに触れずに「データなし」として扱う
    If FileLen(filePath) = 0 Then
        lineCount = 0
        ReDim lines(0 To 0)
        Exit Sub
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = INPUT_ENCODING
    stream.Open
    stream.LoadFromFile filePath

    Dim content As String
    content = stream.ReadText(-1) ' adReadAll
    stream.Close

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    Dim arr() As String
    arr = Split(content, vbLf)

    Dim n As Long
    n = UBound(arr)
    Do While n >= 0 And Trim(arr(n)) = ""
        n = n - 1
    Loop

    If n < 0 Then
        lineCount = 0
        ReDim lines(0 To 0)
        Exit Sub
    End If

    lineCount = n + 1
    ReDim lines(0 To n)

    Dim i As Long
    For i = 0 To n
        lines(i) = arr(i)
    Next i

End Sub

'==============================================================
' CSVの1行を区切り文字で分割する（ダブルクォート対応）
'==============================================================
Private Function SplitCSVLine(ByVal lineText As String) As String()

    Dim result() As String
    Dim resultCount As Long
    resultCount = 0
    ReDim result(0 To 0)

    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim buf As String
    inQuotes = False
    buf = ""

    For i = 1 To Len(lineText)
        ch = Mid(lineText, i, 1)

        If inQuotes Then
            If ch = """" Then
                If i < Len(lineText) And Mid(lineText, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                buf = buf & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = DELIMITER Then
                result(resultCount) = buf
                resultCount = resultCount + 1
                ReDim Preserve result(0 To resultCount)
                buf = ""
            Else
                buf = buf & ch
            End If
        End If
    Next i

    result(resultCount) = buf
    SplitCSVLine = result

End Function

'==============================================================
' 「差分一覧」シートに結果を書き出す
'==============================================================
Private Sub WriteResultSheet(ByRef headers() As String, ByRef resultRows As Collection, ByVal addedCount As Long, ByVal deletedCount As Long, ByVal changedCount As Long, ByVal oldFile As String, ByVal newFile As String, ByRef oldDupKeys As Object, ByRef newDupKeys As Object)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(OUTPUT_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "変更前ファイル"
    ws.Cells(1, 2).Value = oldFile
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(2, 1).Value = "変更後ファイル"
    ws.Cells(2, 2).Value = newFile
    ws.Cells(2, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "実行日時"
    ws.Cells(3, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(3, 1).Font.Bold = True

    ws.Cells(4, 1).Value = "件数"
    ws.Cells(4, 1).Font.Bold = True
    ws.Cells(4, 2).Value = "追加: " & addedCount & " 件 / 削除: " & deletedCount & " 件 / 変更: " & changedCount & " 件"

    ws.Cells(5, 1).Value = "重複キー"
    ws.Cells(5, 1).Font.Bold = True
    ws.Cells(5, 2).Value = "変更前: " & FormatDupKeys(oldDupKeys) & " / 変更後: " & FormatDupKeys(newDupKeys)

    If oldDupKeys.Count > 0 Or newDupKeys.Count > 0 Then
        ws.Cells(5, 1).Font.Color = RGB(192, 0, 0)
        ws.Cells(5, 2).Font.Color = RGB(192, 0, 0)

        ws.Cells(6, 1).Value = "【注意】上記の重複キーは、ファイル内で最初に出現した行の内容のみを使って比較しています。" & _
                                "出現2回目以降の内容は、比較・出力に反映されていません。"
        ws.Cells(6, 1).Font.Bold = True
        ws.Cells(6, 1).Font.Color = RGB(192, 0, 0)
    End If

    Dim headerRow As Long
    headerRow = 8

    ws.Cells(headerRow, 1).Value = "判定"
    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, c + 2).Value = headers(c)
    Next c
    ws.Rows(headerRow).Font.Bold = True
    ws.Rows(headerRow).Interior.Color = RGB(221, 235, 247)

    Dim outRow As Long
    outRow = headerRow + 1

    Dim item As Variant
    Dim judgement As String
    Dim rowCols() As String
    Dim rowColor As Long
    Dim lastCol As Long

    Dim i As Long
    For i = 1 To resultRows.Count
        item = resultRows(i)
        judgement = item(0)
        rowCols = item(1)

        Select Case judgement
            Case "追加"
                rowColor = RGB(198, 239, 206) ' 緑系
            Case "削除"
                rowColor = RGB(255, 199, 206) ' 赤系
            Case Else ' 変更
                rowColor = RGB(255, 235, 156) ' 黄系
        End Select

        ws.Cells(outRow, 1).Value = judgement
        For c = LBound(rowCols) To UBound(rowCols)
            ws.Cells(outRow, c + 2).Value = rowCols(c)
        Next c

        lastCol = UBound(rowCols) + 2
        ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, lastCol)).Interior.Color = rowColor

        outRow = outRow + 1
    Next i

    If resultRows.Count = 0 Then
        ws.Cells(outRow, 1).Value = "差分はありませんでした。"
    End If

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select

End Sub

'==============================================================
' 重複キー→行番号一覧のDictionaryを、表示用の文字列に変換する
'   例: "田中太郎（2行目, 3行目）/ 鈴木一郎（5行目, 6行目, 7行目）"
'   重複なしの場合は "なし"
'==============================================================
Private Function FormatDupKeys(ByRef dupKeyRows As Object) As String

    If dupKeyRows.Count = 0 Then
        FormatDupKeys = "なし"
        Exit Function
    End If

    Dim result As String
    Dim k As Variant
    Dim rowList As Collection
    Dim rowsStr As String
    Dim i As Long

    For Each k In dupKeyRows.Keys
        Set rowList = dupKeyRows(k)

        rowsStr = ""
        For i = 1 To rowList.Count
            If rowsStr <> "" Then rowsStr = rowsStr & ", "
            rowsStr = rowsStr & rowList(i) & "行目"
        Next i

        If result <> "" Then result = result & " / "
        result = result & k & "（" & rowsStr & "）"
    Next k

    FormatDupKeys = result

End Function

'==============================================================
' 比較対象のデータが見つからなかった場合に、その旨を
' 「差分一覧」シートに書き出す
'==============================================================
Private Sub WriteNoDataSheet(ByVal noticeMessage As String, ByVal oldFile As String, ByVal newFile As String)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(OUTPUT_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "変更前ファイル"
    ws.Cells(1, 2).Value = oldFile
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(2, 1).Value = "変更後ファイル"
    ws.Cells(2, 2).Value = newFile
    ws.Cells(2, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "実行日時"
    ws.Cells(3, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(3, 1).Font.Bold = True

    ws.Cells(5, 1).Value = noticeMessage
    ws.Cells(5, 1).Font.Bold = True
    ws.Cells(5, 1).Font.Color = RGB(192, 0, 0)

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select

End Sub

'==============================================================
' シートが存在すれば取得、なければ新規作成して返す
'==============================================================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws

End Function
