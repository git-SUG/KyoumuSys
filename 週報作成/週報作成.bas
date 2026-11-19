Attribute VB_Name = "週報作成マクロ"
Option Explicit
'===============================================================================
' 週報作成マクロ
'
'  「年間行事予定（月別のワークシート）」ブックの ４月～３月 シートから、
'  指定した週の月曜日から31日分の
'     日直（C列）・校内行事（D列）・校外行事（E列）・提出文書・備考（F列）
'  を読み取り、週報シート（「原本」のコピー）の G13:L43 に書き込みます。
'  月をまたぐ週でも自動で次の月のシートから読み取ります。
'
'  使い方：Alt+F8 →「週報作成」→ 実行
'    ① 週の番号を入力（記入済みの最後の週の次が入っています）
'    ② 週の開始日（月曜日）を確認。前の週＋7日が入っているので、
'       長期休業明けなどで違うときだけ直す
'  （ショートカットキーは Alt+F8 →［オプション］で設定できます）
'===============================================================================

'----- 設定 -----------------------------------------------------------------
Private Const SHEET_GENPON As String = "原本"     ' 週報のひな形シート
Private Const DAYS_TO_COPY As Long = 31           ' 転記する日数（G13:L43 = 31行）
Private Const FIRST_ROW As Long = 13              ' 転記先の先頭行
Private Const SRC_FIRST_ROW As Long = 3           ' 月シートで「1日」がある行
Private Const TIDY_SCHEDULE As Boolean = True     ' 今後の予定欄(D11)で行事のない日を省く
Private Const NAME_SRC_PATH As String = "週報作成_行事予定パス"

'===============================================================================
' メイン：週番号を聞いて週報シートを作る
'===============================================================================
Public Sub 週報作成()
    ' 1) 週番号（既定値＝記入済みの最後の週の次）
    Dim defWeek As Long
    defWeek = LastFilledWeek() + 1

    Dim ans As String, weekNo As Long
    ans = InputBox("作成する週の番号を入力してください。", "週報作成", CStr(defWeek))
    If ans = "" Then Exit Sub
    ans = StrConv(ans, vbNarrow)
    If Not IsNumeric(ans) Then
        MsgBox "週の番号は数字で入力してください。", vbExclamation, "週報作成"
        Exit Sub
    End If
    weekNo = CLng(ans)
    If weekNo < 1 Then Exit Sub

    ' 2) 週の開始日（月曜日）：前の週から計算して表示し、訂正できるようにする
    Dim startDate As Date, note As String
    If Not AskStartDate(weekNo, startDate, note) Then Exit Sub

    ' 3) 行事予定ブックを用意して書き込む
    Dim wbSrc As Workbook, openedHere As Boolean
    Set wbSrc = GetSourceWorkbook(openedHere)
    If wbSrc Is Nothing Then Exit Sub

    Dim msg As String
    msg = MakeWeeklyReport(wbSrc, weekNo, startDate, True)

    If openedHere Then wbSrc.Close SaveChanges:=False
    If msg <> "" Then MsgBox msg & note, vbInformation, "週報作成"
End Sub

' 開始日を前の週から計算して表示し、訂正を受け付ける
'   訂正なし（OK のみ）→ 計算した日付のまま
'   訂正あり           → 入力した日付（月曜日でなければ、その週の月曜日）
Private Function AskStartDate(ByVal weekNo As Long, ByRef result As Date, ByRef note As String) As Boolean
    Dim calc As Date, prevNo As Long, prompt As String, defText As String
    note = ""
    If PrevWeekStart(weekNo, calc, prevNo) Then
        defText = Format(calc, "yyyy/m/d")
        prompt = "第" & weekNo & "週の開始日（月曜日）" & vbLf & vbLf & _
                 "前の週（第" & prevNo & "週：" & Format(calc - 7 * (weekNo - prevNo), "m/d") & _
                 "～）から計算すると " & Format(calc, "m/d") & "（" & WeekdayChar(calc) & "）です。" & vbLf & vbLf & _
                 "長期休業などで違う場合は、正しい日付に直してください。" & vbLf & _
                 "（例：2026/8/31　または　8/31）"
    Else
        defText = ""
        prompt = "第" & weekNo & "週の開始日（月曜日）を入力してください。" & vbLf & _
                 "（前の週のシートから日付を求められませんでした）" & vbLf & vbLf & _
                 "（例：2026/8/31　または　8/31）"
    End If

    Dim s As String
    s = InputBox(prompt, "週報作成：週の開始日", defText)
    If s = "" Then Exit Function                         ' キャンセル
    s = StrConv(Trim$(s), vbNarrow)

    If defText <> "" And s = defText Then
        result = calc                                    ' 訂正なし
    Else
        If Not IsDate(s) Then
            MsgBox "日付として読み取れませんでした：" & s, vbExclamation, "週報作成"
            Exit Function
        End If
        Dim d As Date
        d = DateValue(s)
        If InStr(s, "/") = InStrRev(s, "/") And defText <> "" Then
            ' 「8/31」のように年を省略したときは、計算した日付に近い年を使う
            d = DateSerial(Year(calc), Month(d), Day(d))
            If d - calc > 180 Then d = DateSerial(Year(calc) - 1, Month(d), Day(d))
            If calc - d > 180 Then d = DateSerial(Year(calc) + 1, Month(d), Day(d))
        End If
        result = MondayOf(d)
        If result <> d Then
            note = vbLf & vbLf & "※入力された " & Format(d, "m/d") & "（" & WeekdayChar(d) & _
                   "）は月曜日ではないため、その週の月曜日 " & Format(result, "m/d") & " にしました。"
        End If
    End If
    AskStartDate = True
End Function

'===============================================================================
' 本体：週報シートを用意してデータを書き込む（戻り値＝完了メッセージ、中止時は ""）
'===============================================================================
Public Function MakeWeeklyReport(ByVal wbSrc As Workbook, ByVal weekNo As Long, _
                                 ByVal startDate As Date, ByVal askOverwrite As Boolean) As String
    Dim ws As Worksheet
    Set ws = FindSheet(ThisWorkbook, CStr(weekNo))

    If ws Is Nothing Then
        Dim wsGen As Worksheet
        Set wsGen = FindSheet(ThisWorkbook, SHEET_GENPON)
        If wsGen Is Nothing Then
            MsgBox "「" & SHEET_GENPON & "」シートが見つかりません。", vbCritical, "週報作成"
            Exit Function
        End If
        wsGen.Copy After:=LastWeekSheet()
        Set ws = ActiveSheet
        ws.Name = CStr(weekNo)
        ws.Visible = xlSheetVisible
    ElseIf askOverwrite Then
        If Application.WorksheetFunction.CountA(ws.Range("G13:L43")) > 0 Then
            If MsgBox("「" & ws.Name & "」シートにはすでにデータがあります。" & vbLf & _
                      "行事予定から読み直して上書きしますか？" & vbLf & _
                      "（G13:L43 の貼付データだけを書き換えます。お知らせ欄などはそのままです）", _
                      vbYesNo + vbQuestion, "週報作成") <> vbYes Then Exit Function
        End If
    End If

    Application.ScreenUpdating = False

    ' 年度・月・週の初日
    ws.Range("F12").Value = SchoolYear(startDate)
    ws.Range("H12").Value = Month(startDate)
    ws.Range("F13").ClearContents
    ws.Range("F13").Value = startDate

    ' 31日分を転記
    Dim i As Long, d As Date, r As Long, wsM As Worksheet, curMonth As Long
    Dim missing As String
    curMonth = 0
    ws.Range("G" & FIRST_ROW & ":L" & FIRST_ROW + DAYS_TO_COPY - 1).ClearContents
    For i = 0 To DAYS_TO_COPY - 1
        d = startDate + i
        If Month(d) <> curMonth Then
            curMonth = Month(d)
            Set wsM = MonthSheet(wbSrc, curMonth)
            If wsM Is Nothing Then missing = missing & " " & curMonth & "月"
        End If
        r = FIRST_ROW + i
        ws.Cells(r, "G").Value = Day(d)
        ws.Cells(r, "H").Value = Mid("日月火水木金土", Weekday(d, vbSunday), 1)
        If Not wsM Is Nothing Then
            Dim sr As Long
            sr = SRC_FIRST_ROW + Day(d) - 1
            ws.Cells(r, "I").Value = CleanText(wsM.Cells(sr, "C"))   ' 日直
            ws.Cells(r, "J").Value = CleanText(wsM.Cells(sr, "D"))   ' 校内行事
            ws.Cells(r, "K").Value = CleanText(wsM.Cells(sr, "E"))   ' 校外行事
            ws.Cells(r, "L").Value = CleanText(wsM.Cells(sr, "F"))   ' 提出文書・備考
        End If
    Next i

    If TIDY_SCHEDULE Then TidyScheduleFormulas ws
    KeepOutsideSpaces ws

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A1").Select

    MakeWeeklyReport = "第" & weekNo & "週（" & Format(startDate, "m/d") & "～" & _
                       Format(startDate + 6, "m/d") & "）の週報を作成しました。" & vbLf & _
                       "行事予定ブック：" & wbSrc.Name
    If missing <> "" Then
        MakeWeeklyReport = MakeWeeklyReport & vbLf & vbLf & _
            "※次の月のシートが行事予定ブックに見つからず、空欄にしています：" & missing
    End If
End Function

'===============================================================================
' 補助
'===============================================================================

' 前の週の開始日＋7日を返す（見つかった前の週の番号も返す）
'   さかのぼって最初に正しい開始日を持つ週を使う（空のひな形シートは飛ばす）
Private Function PrevWeekStart(ByVal weekNo As Long, ByRef result As Date, ByRef prevNo As Long) As Boolean
    Dim k As Long, ws As Worksheet, d As Date
    For k = weekNo - 1 To 1 Step -1
        Set ws = FindSheet(ThisWorkbook, CStr(k))
        If Not ws Is Nothing Then
            If SheetStartDate(ws, d) Then
                result = d + 7 * (weekNo - k)
                prevNo = k
                PrevWeekStart = True
                Exit Function
            End If
        End If
    Next k
End Function

' 週シートの開始日（月曜日）。古い様式の「週初日」(H2/G2) → F13 の順に探す
'   月曜日で、今年度の日付であるものだけを正しい開始日とみなす
Private Function SheetStartDate(ByVal ws As Worksheet, ByRef result As Date) As Boolean
    Dim addr As Variant, v As Variant
    For Each addr In Array("H2", "G2", "F13")
        v = ws.Range(addr).Value
        If Not IsError(v) Then
            ' 日付、または日付を表す数値（2000/1/1～2099/12/31）だけを対象にする
            If VarType(v) = vbDate Or ((VarType(v) = vbDouble Or VarType(v) = vbLong _
                                       Or VarType(v) = vbInteger) And v >= 36526 And v <= 73415) Then
                If Weekday(CDate(v), vbMonday) = 1 And IsThisSchoolYear(CDate(v)) Then
                    result = CDate(v)
                    SheetStartDate = True
                    Exit Function
                End If
            End If
        End If
    Next addr
End Function

' 原本 F12 の年度と同じ年度の日付か（F12 が読めなければ判定しない）
Private Function IsThisSchoolYear(ByVal d As Date) As Boolean
    Dim ws As Worksheet, y As Variant
    IsThisSchoolYear = True
    Set ws = FindSheet(ThisWorkbook, SHEET_GENPON)
    If ws Is Nothing Then Exit Function
    y = ws.Range("F12").Value
    If IsNumeric(y) And Not IsEmpty(y) Then IsThisSchoolYear = (SchoolYear(d) = CLng(y))
End Function

' 記入済み（G13:L43 にデータがある）週シートのうち、番号が最大のもの
Private Function LastFilledWeek() As Long
    Dim ws As Worksheet, n As Long
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            n = CLng(ws.Name)
            If n > LastFilledWeek Then
                If Application.WorksheetFunction.CountA(ws.Range("G13:L43")) > 0 Then LastFilledWeek = n
            End If
        End If
    Next ws
End Function

Private Function MondayOf(ByVal d As Date) As Date
    MondayOf = d - (Weekday(d, vbMonday) - 1)
End Function

Private Function WeekdayChar(ByVal d As Date) As String
    WeekdayChar = Mid("日月火水木金土", Weekday(d, vbSunday), 1)
End Function

' 行事予定ブックを探す：開いていればそれを使い、なければ前回のパス→ファイル選択
Private Function GetSourceWorkbook(ByRef openedHere As Boolean) As Workbook
    Dim wb As Workbook
    openedHere = False
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If IsSourceWorkbook(wb) Then
                Set GetSourceWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    Dim path As String
    path = LoadSavedPath()
    If path <> "" Then
        If Dir(path) = "" Then path = ""
    End If
    If path = "" Then
        Dim f As Variant
        f = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                        "年間行事予定（月別のワークシート）のファイルを選んでください")
        If VarType(f) = vbBoolean Then Exit Function
        path = CStr(f)
    End If

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    ThisWorkbook.Activate
    If Not IsSourceWorkbook(wb) Then
        MsgBox "選んだファイルに「４月」～「３月」のシートが見つかりません。" & vbLf & wb.Name, _
               vbExclamation, "週報作成"
        wb.Close SaveChanges:=False
        SaveSourcePath ""
        Exit Function
    End If
    SaveSourcePath path
    openedHere = True
    Set GetSourceWorkbook = wb
End Function

Private Function IsSourceWorkbook(ByVal wb As Workbook) As Boolean
    IsSourceWorkbook = Not (MonthSheet(wb, 4) Is Nothing) And Not (MonthSheet(wb, 12) Is Nothing)
End Function

' 月シート（「１２月」のような全角、または「12月」）を返す
Private Function MonthSheet(ByVal wb As Workbook, ByVal m As Long) As Worksheet
    Set MonthSheet = FindSheet(wb, StrConv(CStr(m), vbWide) & "月")
    If MonthSheet Is Nothing Then Set MonthSheet = FindSheet(wb, CStr(m) & "月")
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

' 数字名の週シートのうち番号が最大のもの（なければ最後のシート）
Private Function LastWeekSheet() As Worksheet
    Dim ws As Worksheet, best As Worksheet, n As Long, bestN As Long
    bestN = -1
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            n = CLng(ws.Name)
            If n > bestN Then bestN = n: Set best = ws
        End If
    Next ws
    If best Is Nothing Then Set best = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set LastWeekSheet = best
End Function

' 空の数式が返す 0 やエラー値を空欄にする
Private Function CleanText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "0" Then Exit Function
        CleanText = v
    ElseIf IsNumeric(v) Then
        If v = 0 Then Exit Function
        CleanText = CStr(v)
    Else
        CleanText = CStr(v)
    End If
End Function

Private Function SchoolYear(ByVal d As Date) As Long
    If Month(d) < 4 Then SchoolYear = Year(d) - 1 Else SchoolYear = Year(d)
End Function

' 今後の予定（F20:F43 → G11 → D11）で、校内行事のない日を表示しないようにする
Private Sub TidyScheduleFormulas(ByVal ws As Worksheet)
    Dim r As Long, parts As String
    For r = 20 To 43
        ws.Range("F" & r).Formula = "=IF(J" & r & "="""","""",G" & r & "&""(""&H" & r & _
            "&"")" & ChrW(&H3000) & ChrW(&H3007) & """&SUBSTITUTE(J" & r & ",""" & ChrW(&H3000) & _
            """,CHAR(10)&""" & ChrW(&H3007) & """))"
        If r <= 30 Then
            parts = parts & ",SUBSTITUTE(F" & r & ",CHAR(10),"""")"
        Else
            parts = parts & ",F" & r
        End If
    Next r
    ws.Range("G11").Formula = "=TEXTJOIN(CHAR(10),TRUE" & parts & ")"
End Sub

' 校外行事欄（D4:D10）で全角スペースを消さないようにする
'  元の数式： =IF(K13="","",SUBSTITUTE(K13,"　","")&CHAR(10))&IF(J4="","",J4)
'  変更後　： =IF(K13="","",K13&CHAR(10))&IF(J4="","",J4)
Private Sub KeepOutsideSpaces(ByVal ws As Worksheet)
    Dim i As Long
    For i = 0 To 6
        ws.Range("D" & (4 + i)).Formula = "=IF(K" & (13 + i) & "="""","""",K" & (13 + i) & _
            "&CHAR(10))&IF(J" & (4 + i) & "="""","""",J" & (4 + i) & ")"
    Next i
End Sub

Private Function LoadSavedPath() As String
    On Error Resume Next
    Dim s As String
    s = ThisWorkbook.Names(NAME_SRC_PATH).RefersTo
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    ' RefersTo は ="C:\..." の形
    If Left$(s, 2) = "=""" Then s = Mid$(s, 3, Len(s) - 3)
    LoadSavedPath = Replace(s, """""", """")
End Function

Private Sub SaveSourcePath(ByVal path As String)
    On Error Resume Next
    ThisWorkbook.Names(NAME_SRC_PATH).Delete
    On Error GoTo 0
    If path = "" Then Exit Sub
    ThisWorkbook.Names.Add Name:=NAME_SRC_PATH, _
        RefersTo:="=""" & Replace(path, """", """""") & """", Visible:=False
End Sub
