Attribute VB_Name = "AbcReport"
Option Explicit

'======================================================================
' AbcReport モジュール
' ABC 区分の判定 (F-02) と、その結果と根拠をまとめた説明資料 (F-03) を作る。
'   - シート「分類レポート」に、結果のまとめ・対数グラフ・判定の方法・品目一覧を、印刷向けの配置で書く
'   - グラフ: 順位 x 処方回数 (対数目盛り) の散布図。点は区分ごとに色を分け、当てはめた直線と境界線、参考のひざを重ねる
'   - シートを PDF に書き出す (A4 縦、横幅 1 ページ)
' 載せるのは薬品単位の情報だけで、患者情報は扱わない。
'
' 依存: ErrorLogger.bas, AbcClassifier.bas, OrderCalc.bas (区分の定数)
'======================================================================

Private Const MODULE_NAME As String = "AbcReport"

' ---- シートと配置 ----
Private Const SHT_REPORT As String = "分類レポート"
Private Const COL_COUNT As Long = 6                ' 使う列の数 (A-F)
Private Const COL_W_1 As Double = 14
Private Const COL_W_2 As Double = 14
Private Const COL_W_3 As Double = 30
Private Const COL_W_4 As Double = 18
Private Const COL_W_5 As Double = 16
Private Const COL_W_6 As Double = 18
Private Const CHART_ROWS As Long = 24              ' グラフに使う行の数
Private Const TEXT_CHARS_PER_LINE As Long = 58     ' 説明の 1 行に入る全角の文字数の目安 (A-F を結合した幅)
Private Const TEXT_LINE_HEIGHT As Double = 15      ' 説明の 1 行の高さ (pt)
Private Const TITLE_FONT_SIZE As Double = 16
Private Const SECTION_FONT_SIZE As Double = 12
Private Const BODY_FONT_SIZE As Double = 10
Private Const HEAD_FILL As Long = 15652305         ' 表の見出しの色 (RGB 209, 226, 237 相当)
Private Const PAGE_MARGIN_IN As Double = 0.5

' ---- グラフの色 (RGB) ----
Private Const COLOR_A As Long = 12611584           ' A (青)
Private Const COLOR_B As Long = 3243501            ' B (橙)
Private Const COLOR_C As Long = 5287936            ' C (緑)
Private Const COLOR_BOUND As Long = 255            ' 境界線 (赤)
Private Const COLOR_KNEE As Long = 10921638        ' 参考のひざ (薄い灰)
Private Const LINE_WEIGHT As Double = 1.25
Private Const MARKER_SIZE As Long = 4
Private Const LOG_BASE As Double = 10           ' 縦軸の対数目盛りの底
Private Const AXIS_MIN_COUNT As Double = 1     ' 縦軸の最小値 (処方回数は 1 以上)
Private Const COLOR_FIT As Long = 0              ' 当てはめた直線 (黒)
Private Const BOUND_OFFSET As Double = 0.5         ' 境界線を順位と順位の間に引く

' ---- 報告の文言 ----
Private Const TXT_TITLE As String = "ABC分類レポート"
Private Const TXT_PERIOD As String = "対象期間: "
Private Const TXT_RUN_AT As String = "作成日時: "
Private Const TXT_SEC_SUMMARY As String = "1. 結果のまとめ"
Private Const TXT_SEC_GRAPH As String = "2. 判定の根拠 (順位と処方回数の対数グラフ)"
Private Const TXT_SEC_METHOD As String = "3. 判定の方法と条件"
Private Const TXT_SEC_ITEMS As String = "4. 品目一覧"
Private Const TXT_H_CLASS As String = "区分"
Private Const TXT_H_ITEMS As String = "品目数"
Private Const TXT_H_RANKS As String = "順位"
Private Const TXT_H_COUNTS As String = "処方回数"
Private Const TXT_H_SHARE As String = "処方回数の構成比"
Private Const TXT_H_DECAY As String = "1順位あたりの減り方"
Private Const TXT_H_RANK As String = "順位"
Private Const TXT_H_CODE As String = "薬品コード"
Private Const TXT_H_NAME As String = "薬品名"
Private Const TXT_H_SPEC As String = "規格"
Private Const TXT_H_COUNT As String = "処方回数"
Private Const TXT_ZERO_ROW As String = "C (処方回数0)"
Private Const TXT_RANGE_SEP As String = "-"
Private Const TXT_DECAY_UNIT As String = "% 減"
Private Const TXT_NONE As String = "-"
Private Const TXT_SERIES_POINTS As String = " の品目"
Private Const TXT_SERIES_FIT As String = " の当てはめ"
Private Const TXT_SERIES_AB As String = "A/B の境界 (順位 "
Private Const TXT_SERIES_BC As String = "B/C の境界 (順位 "
Private Const TXT_SERIES_KNEE As String = "参考: 累積構成比のひざ (順位 "
Private Const TXT_CLOSE As String = ")"
Private Const TXT_AXIS_X As String = "順位 (処方回数の多い順)"
Private Const TXT_AXIS_Y As String = "処方回数 (対数目盛り)"
Private Const TXT_READ_1 As String = "グラフの読み方: 横軸は処方回数の多い順の順位、縦軸は処方回数を対数目盛りで表す。対数目盛りでは、処方回数が一定の割合で減る並びがまっすぐな線になる。"
Private Const TXT_READ_A As String = "A の区間 (順位 1-"
Private Const TXT_READ_B As String = "B の区間 (順位 "
Private Const TXT_READ_C As String = "C の区間 (順位 "
Private Const TXT_READ_SLOPE As String = "): 1順位下がるごとに処方回数が約 "
Private Const TXT_READ_A_END As String = "% ずつ減る。傾きがゆるやかで、毎日のように出る品目の並び。"
Private Const TXT_READ_B_END As String = "% ずつ減る。傾きが急で、順位とともに処方回数が落ちていく並び。"
Private Const TXT_READ_C_END As String = "% ずつ減る。1回から数回の品目が平らに並ぶ裾。"
Private Const TXT_READ_WHY As String = "境界は、この3本の直線と点のずれ (二乗誤差の合計) がいちばん小さくなる位置として自動で決めた。傾きが変わる位置が A/B と B/C の境目になる。"
Private Const TXT_METHOD_1 As String = "(1) 処方回数0の品目を除き、処方回数の多い順に並べる (処方回数0の品目は C とする)。"
Private Const TXT_METHOD_2 As String = "(2) 順位を横軸、処方回数の対数を縦軸にとり、3つの区間それぞれに直線を当てはめる。"
Private Const TXT_METHOD_3 As String = "(3) 2つの折れ点の組み合わせをすべて試し、直線と点のずれ (二乗誤差の合計) が最小の組み合わせを選ぶ。"
Private Const TXT_METHOD_4 As String = "(4) 1つ目の折れ点までを A、2つ目の折れ点までを B、それより後ろを C とする。"
Private Const TXT_COND_MIN As String = "条件: 境界の候補にする最小の処方回数 "
Private Const TXT_COND_SEG As String = " 回 (裾の1回・2回の段々を境界にしないため) / 1つの区間の最小の品目数 "
Private Const TXT_COND_SSE As String = "二乗誤差の合計 (最小): "
Private Const TXT_COND_KNEE As String = "参考: 累積構成比の曲線のひざは順位 "
Private Const TXT_COND_KNEE_2 As String = " (ここまでで全処方回数の "
Private Const TXT_COND_KNEE_3 As String = " を占める)"
Private Const TXT_NOTE_AUTO As String = "注: この区分は自動判定の結果。品目マスタの固定区分とフラグ品の扱いは、品目マスタへの反映のときに優先される。"
Private Const TXT_FAILED As String = "品目が少ない、または境界の候補が無いため、境界を判定できなかった。すべての品目を C としている。"
Private Const TXT_ITEMS_UNIT As String = " 品目"
Private Const PDF_PREFIX As String = "ABC分類レポート_"
Private Const PDF_EXT As String = ".pdf"
Private Const RUN_TIME_FORMAT As String = "yyyy-mm-dd hh:nn"
Private Const SHARE_FORMAT As String = "0.0%"
Private Const COUNT_FORMAT As String = "#,##0"
Private Const DECAY_FORMAT As String = "0.0"
Private Const SSE_FORMAT As String = "0.000"
Private Const PCT As Double = 100
Private Const CLASS_ROWS As Long = 4                ' まとめの表の行 (A / B / C / 処方回数0)
Private Const SEG_COUNT As Long = 3                 ' 区分 (区間) の数
Private Const SEG_A As Long = 1
Private Const SEG_B As Long = 2
Private Const SEG_C As Long = 3
Private Const ROW_SUMMARY_SEC As Long = 4           ' 「1. 結果のまとめ」の行
Private Const READ_LINE_LAST As Long = 5            ' グラフの読み方の行の番号の最後 (0: 読み方 1-3: 区間 4: 理由 5: 注)
Private Const READ_LINE_WHY As Long = 4
Private Const METHOD_LINE_LAST As Long = 7          ' 判定の方法と条件の行の数
Private Const TITLE_GAP As String = "    "


'----------------------------------------------------------------------
' CreateAbcReport: 区分を判定し、分類レポートのシートと PDF を作る
'   codes() names() specs() counts()  薬品ごとの値 (1 始まり、並び順は問わない)
'   n               品目数
'   periodLabel     対象期間の表示 (例: "2026-09")。PDF の名前にも使う
'   minCandCount    境界の候補にする最小の処方回数
'   minSegPoints    1 つの区間の最小の品目数
'   outFolder       PDF を書くフォルダ (空なら PDF は書かない)
'   itemClass()     区分 (出力。入力と同じ並び順)
'   pdfPath         書いた PDF のパス (出力。書かなければ空)
'   戻り値          境界を判定できたら True (判定できないときもシートは作り、すべて C とする)
'----------------------------------------------------------------------
Public Function CreateAbcReport(ByRef codes() As String, ByRef names() As String, ByRef specs() As String, _
                                ByRef counts() As Double, ByVal n As Long, ByVal periodLabel As String, _
                                ByVal minCandCount As Double, ByVal minSegPoints As Long, ByVal outFolder As String, _
                                ByRef itemClass() As String, ByRef pdfPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim order() As Long
    Dim m As Long
    Dim zeroN As Long
    Dim i As Long
    Dim j As Long
    Dim gap As Long
    Dim tmp As Long
    Dim sc() As Double
    Dim ok As Boolean
    Dim breakAB As Long
    Dim breakBC As Long
    Dim kneeRank As Long
    Dim sse As Double
    Dim segFrom(1 To SEG_COUNT) As Long
    Dim segTo(1 To SEG_COUNT) As Long
    Dim slope(1 To SEG_COUNT) As Double
    Dim icpt(1 To SEG_COUNT) As Double
    Dim seg As Long
    Dim sx As Double
    Dim sy As Double
    Dim sxx As Double
    Dim sxy As Double
    Dim cnt As Double
    Dim y As Double
    Dim totalCount As Double
    Dim classCount(1 To SEG_COUNT) As Double
    Dim classItems(1 To SEG_COUNT) As Long
    Dim kneeShare As Double
    Dim ws As Worksheet
    Dim r As Long
    Dim rowSummary As Long
    Dim rowChart As Long
    Dim rowRead As Long
    Dim rowItems As Long
    Dim grid() As Variant
    Dim txt As String
    Dim lines As Long
    Dim co As ChartObject
    Dim ch As Chart
    Dim s As Series
    Dim maxCount As Double
    Dim pFrom As Long
    Dim pTo As Long
    Dim classNames(1 To SEG_COUNT) As String
    Dim colors(1 To SEG_COUNT) As Long
    Dim readEnds(1 To SEG_COUNT) As String
    Dim readHeads(1 To SEG_COUNT) As String
    Dim prevAlerts As Boolean

    ReDim itemClass(1 To n)
    classNames(SEG_A) = OrderCalc.CLASS_A
    classNames(SEG_B) = OrderCalc.CLASS_B
    classNames(SEG_C) = OrderCalc.CLASS_C
    colors(SEG_A) = COLOR_A
    colors(SEG_B) = COLOR_B
    colors(SEG_C) = COLOR_C

    ' ---- 処方回数の多い順に並べる (0 の品目は後ろへ。シェルソート、同数は元の順) ----
    ReDim order(1 To n)
    For i = 1 To n
        order(i) = i
        If counts(i) > 0 Then m = m + 1 Else zeroN = zeroN + 1
        If counts(i) > 0 Then totalCount = totalCount + counts(i)
    Next i
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            tmp = order(i)
            j = i
            Do While j > gap
                If counts(order(j - gap)) < counts(tmp) Or _
                   (counts(order(j - gap)) = counts(tmp) And order(j - gap) > tmp) Then
                    order(j) = order(j - gap)
                    j = j - gap
                Else
                    Exit Do
                End If
            Loop
            order(j) = tmp
        Next i
        gap = gap \ 2
    Loop

    ' ---- 境界の判定 ----
    If m > 0 Then
        ReDim sc(1 To m)
        For i = 1 To m
            sc(i) = counts(order(i))
        Next i
        ok = AbcClassifier.FindBoundaries(sc, m, minCandCount, minSegPoints, breakAB, breakBC, kneeRank, sse)
    End If
    If Not ok Then
        breakAB = 0
        breakBC = 0
    End If

    ' 区分 (順位 breakAB まで A、breakBC まで B、それ以降と処方回数 0 は C)
    For i = 1 To n
        If i <= breakAB And i <= m Then
            itemClass(order(i)) = OrderCalc.CLASS_A
            seg = SEG_A
        ElseIf i <= breakBC And i <= m Then
            itemClass(order(i)) = OrderCalc.CLASS_B
            seg = SEG_B
        Else
            itemClass(order(i)) = OrderCalc.CLASS_C
            seg = SEG_C
        End If
        classItems(seg) = classItems(seg) + 1
        If i <= m Then classCount(seg) = classCount(seg) + counts(order(i))
    Next i

    ' ---- 区間ごとの直線 (順位 -> log(処方回数) の回帰直線) ----
    segFrom(SEG_A) = 1: segTo(SEG_A) = breakAB
    segFrom(SEG_B) = breakAB + 1: segTo(SEG_B) = breakBC
    segFrom(SEG_C) = breakBC + 1: segTo(SEG_C) = m
    If ok Then
        For seg = 1 To SEG_COUNT
            sx = 0: sy = 0: sxx = 0: sxy = 0
            cnt = segTo(seg) - segFrom(seg) + 1
            For i = segFrom(seg) To segTo(seg)
                y = Log(sc(i))
                sx = sx + i
                sy = sy + y
                sxx = sxx + CDbl(i) * i
                sxy = sxy + i * y
            Next i
            If cnt > 1 And (sxx - sx * sx / cnt) > 0 Then
                slope(seg) = (sxy - sx * sy / cnt) / (sxx - sx * sx / cnt)
            Else
                slope(seg) = 0
            End If
            icpt(seg) = (sy - slope(seg) * sx) / cnt
        Next seg
        ' 参考: ひざまでの構成比
        For i = 1 To kneeRank
            kneeShare = kneeShare + sc(i)
        Next i
        If totalCount > 0 Then kneeShare = kneeShare / totalCount
    End If

    ' ---- シートを作り直す ----
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHT_REPORT Then
            ws.Delete
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = prevAlerts
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHT_REPORT
    ws.Cells.Font.Size = BODY_FONT_SIZE
    ws.Columns(1).ColumnWidth = COL_W_1
    ws.Columns(2).ColumnWidth = COL_W_2
    ws.Columns(3).ColumnWidth = COL_W_3
    ws.Columns(4).ColumnWidth = COL_W_4
    ws.Columns(5).ColumnWidth = COL_W_5
    ws.Columns(6).ColumnWidth = COL_W_6

    ' 表題
    ws.Cells(1, 1).Value = TXT_TITLE
    ws.Cells(1, 1).Font.Size = TITLE_FONT_SIZE
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = TXT_PERIOD & periodLabel & TITLE_GAP & TXT_RUN_AT & Format(Now, RUN_TIME_FORMAT)

    ' ---- 1. 結果のまとめ ----
    r = ROW_SUMMARY_SEC
    ws.Cells(r, 1).Value = TXT_SEC_SUMMARY
    ws.Cells(r, 1).Font.Size = SECTION_FONT_SIZE
    ws.Cells(r, 1).Font.Bold = True
    rowSummary = r + 1
    ReDim grid(1 To CLASS_ROWS + 1, 1 To COL_COUNT)
    grid(1, 1) = TXT_H_CLASS: grid(1, 2) = TXT_H_ITEMS: grid(1, 3) = TXT_H_RANKS
    grid(1, 4) = TXT_H_COUNTS: grid(1, 5) = TXT_H_SHARE: grid(1, 6) = TXT_H_DECAY
    For seg = 1 To SEG_COUNT
        grid(seg + 1, 1) = classNames(seg)
        If seg = SEG_C Then
            grid(seg + 1, 2) = classItems(seg) - zeroN
        Else
            grid(seg + 1, 2) = classItems(seg)
        End If
        If segTo(seg) >= segFrom(seg) And segFrom(seg) >= 1 Then
            grid(seg + 1, 3) = segFrom(seg) & TXT_RANGE_SEP & segTo(seg)
            grid(seg + 1, 4) = Format(sc(segTo(seg)), COUNT_FORMAT) & TXT_RANGE_SEP & Format(sc(segFrom(seg)), COUNT_FORMAT)
        Else
            grid(seg + 1, 3) = TXT_NONE
            grid(seg + 1, 4) = TXT_NONE
        End If
        If totalCount > 0 Then grid(seg + 1, 5) = Format(classCount(seg) / totalCount, SHARE_FORMAT) Else grid(seg + 1, 5) = TXT_NONE
        If ok Then
            grid(seg + 1, 6) = Format((1 - Exp(slope(seg))) * PCT, DECAY_FORMAT) & TXT_DECAY_UNIT
        Else
            grid(seg + 1, 6) = TXT_NONE
        End If
    Next seg
    grid(CLASS_ROWS + 1, 1) = TXT_ZERO_ROW
    grid(CLASS_ROWS + 1, 2) = zeroN
    grid(CLASS_ROWS + 1, 3) = TXT_NONE
    grid(CLASS_ROWS + 1, 4) = 0
    grid(CLASS_ROWS + 1, 5) = TXT_NONE
    grid(CLASS_ROWS + 1, 6) = TXT_NONE
    With ws.Range(ws.Cells(rowSummary, 1), ws.Cells(rowSummary + CLASS_ROWS, COL_COUNT))
        .NumberFormat = "@"
        .Value = grid
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(rowSummary, 1), ws.Cells(rowSummary, COL_COUNT)).Interior.Color = HEAD_FILL
    ws.Range(ws.Cells(rowSummary, 1), ws.Cells(rowSummary, COL_COUNT)).Font.Bold = True
    r = rowSummary + CLASS_ROWS + 2
    If Not ok Then
        ws.Cells(r, 1).Value = TXT_FAILED
        r = r + 2
    End If

    ' ---- 2. 判定の根拠 (グラフ) ----
    ws.Cells(r, 1).Value = TXT_SEC_GRAPH
    ws.Cells(r, 1).Font.Size = SECTION_FONT_SIZE
    ws.Cells(r, 1).Font.Bold = True
    rowChart = r + 1
    r = rowChart + CHART_ROWS + 1

    ' グラフの読み方 (A-F を結合し、文字数から行の高さを決める)
    rowRead = r
    readHeads(SEG_A) = TXT_READ_A
    readHeads(SEG_B) = TXT_READ_B
    readHeads(SEG_C) = TXT_READ_C
    readEnds(SEG_A) = TXT_READ_A_END
    readEnds(SEG_B) = TXT_READ_B_END
    readEnds(SEG_C) = TXT_READ_C_END
    For i = 0 To READ_LINE_LAST
        Select Case i
            Case 0
                txt = TXT_READ_1
            Case SEG_A, SEG_B, SEG_C
                If ok Then
                    If i = SEG_A Then
                        txt = readHeads(i) & segTo(i)
                    Else
                        txt = readHeads(i) & segFrom(i) & TXT_RANGE_SEP & segTo(i)
                    End If
                    txt = txt & TXT_READ_SLOPE & Format((1 - Exp(slope(i))) * PCT, DECAY_FORMAT) & readEnds(i)
                Else
                    txt = ""
                End If
            Case READ_LINE_WHY
                txt = IIf(ok, TXT_READ_WHY, "")
            Case Else
                txt = TXT_NOTE_AUTO
        End Select
        If Len(txt) > 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Merge
            ws.Cells(r, 1).Value = txt
            ws.Cells(r, 1).WrapText = True
            ws.Cells(r, 1).VerticalAlignment = xlTop
            lines = -Int(-Len(txt) / TEXT_CHARS_PER_LINE)
            ws.Rows(r).RowHeight = lines * TEXT_LINE_HEIGHT
            r = r + 1
        End If
    Next i
    r = r + 1

    ' ---- 3. 判定の方法と条件 ----
    ws.Cells(r, 1).Value = TXT_SEC_METHOD
    ws.Cells(r, 1).Font.Size = SECTION_FONT_SIZE
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    For i = 1 To METHOD_LINE_LAST
        Select Case i
            Case 1: txt = TXT_METHOD_1
            Case 2: txt = TXT_METHOD_2
            Case 3: txt = TXT_METHOD_3
            Case 4: txt = TXT_METHOD_4
            Case 5: txt = TXT_COND_MIN & minCandCount & TXT_COND_SEG & minSegPoints & TXT_ITEMS_UNIT
            Case 6: txt = IIf(ok, TXT_COND_SSE & Format(sse, SSE_FORMAT), "")
            Case Else
                If ok And kneeRank > 0 Then
                    txt = TXT_COND_KNEE & kneeRank & TXT_COND_KNEE_2 & Format(kneeShare, SHARE_FORMAT) & TXT_COND_KNEE_3
                Else
                    txt = ""
                End If
        End Select
        If Len(txt) > 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Merge
            ws.Cells(r, 1).Value = txt
            ws.Cells(r, 1).WrapText = True
            ws.Cells(r, 1).VerticalAlignment = xlTop
            lines = -Int(-Len(txt) / TEXT_CHARS_PER_LINE)
            ws.Rows(r).RowHeight = lines * TEXT_LINE_HEIGHT
            r = r + 1
        End If
    Next i
    r = r + 1

    ' ---- 4. 品目一覧 (処方回数の多い順。0 の品目は後ろ) ----
    ws.Cells(r, 1).Value = TXT_SEC_ITEMS
    ws.Cells(r, 1).Font.Size = SECTION_FONT_SIZE
    ws.Cells(r, 1).Font.Bold = True
    rowItems = r + 1
    ReDim grid(1 To n + 1, 1 To COL_COUNT)
    grid(1, 1) = TXT_H_RANK: grid(1, 2) = TXT_H_CODE: grid(1, 3) = TXT_H_NAME
    grid(1, 4) = TXT_H_SPEC: grid(1, 5) = TXT_H_COUNT: grid(1, 6) = TXT_H_CLASS
    For i = 1 To n
        If i <= m Then grid(i + 1, 1) = i Else grid(i + 1, 1) = TXT_NONE
        grid(i + 1, 2) = codes(order(i))
        grid(i + 1, 3) = names(order(i))
        grid(i + 1, 4) = specs(order(i))
        grid(i + 1, 5) = counts(order(i))
        grid(i + 1, 6) = itemClass(order(i))
    Next i
    ws.Range(ws.Cells(rowItems + 1, 2), ws.Cells(rowItems + n, 4)).NumberFormat = "@"
    With ws.Range(ws.Cells(rowItems, 1), ws.Cells(rowItems + n, COL_COUNT))
        .Value = grid
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(rowItems, 1), ws.Cells(rowItems, COL_COUNT)).Interior.Color = HEAD_FILL
    ws.Range(ws.Cells(rowItems, 1), ws.Cells(rowItems, COL_COUNT)).Font.Bold = True
    ws.Range(ws.Cells(rowItems + 1, 5), ws.Cells(rowItems + n, 5)).NumberFormat = COUNT_FORMAT

    ' ---- グラフ (品目一覧の順位と処方回数の列を使う) ----
    If m > 0 Then
        Set co = ws.ChartObjects.Add(ws.Cells(rowChart, 1).Left, ws.Cells(rowChart, 1).Top, _
                                     ws.Range(ws.Cells(rowChart, 1), ws.Cells(rowChart, COL_COUNT)).Width, _
                                     ws.Range(ws.Cells(rowChart, 1), ws.Cells(rowChart + CHART_ROWS - 1, 1)).Height)
        Set ch = co.Chart
        ch.ChartType = xlXYScatter
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop
        maxCount = sc(1)

        ' 点 (順位, 処方回数)。区分ごとに色を分ける (判定できないときは全部 C)
        For seg = 1 To SEG_COUNT
            If ok Then
                pFrom = segFrom(seg)
                pTo = segTo(seg)
            ElseIf seg = SEG_C Then
                pFrom = 1
                pTo = m
            Else
                pFrom = 0
                pTo = -1
            End If
            If pTo >= pFrom And pFrom >= 1 Then
                Set s = ch.SeriesCollection.NewSeries
                s.Name = classNames(seg) & TXT_SERIES_POINTS
                s.XValues = ws.Range(ws.Cells(rowItems + pFrom, 1), ws.Cells(rowItems + pTo, 1))
                s.Values = ws.Range(ws.Cells(rowItems + pFrom, 5), ws.Cells(rowItems + pTo, 5))
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = MARKER_SIZE
                s.MarkerForegroundColor = colors(seg)
                s.MarkerBackgroundColor = colors(seg)
                s.Format.Line.Visible = msoFalse
            End If
        Next seg

        If ok Then
            ' 区間ごとの当てはめた直線
            For seg = 1 To SEG_COUNT
                If segTo(seg) >= segFrom(seg) Then
                    Set s = ch.SeriesCollection.NewSeries
                    s.Name = classNames(seg) & TXT_SERIES_FIT
                    s.XValues = Array(segFrom(seg), segTo(seg))
                    s.Values = Array(Exp(icpt(seg) + slope(seg) * segFrom(seg)), Exp(icpt(seg) + slope(seg) * segTo(seg)))
                    s.ChartType = xlXYScatterLinesNoMarkers
                    s.Format.Line.Visible = msoTrue
                    s.Format.Line.ForeColor.RGB = COLOR_FIT
                    s.Format.Line.Weight = LINE_WEIGHT
                End If
            Next seg
            ' 境界線 (縦線)
            For seg = SEG_A To SEG_B
                Set s = ch.SeriesCollection.NewSeries
                If seg = 1 Then
                    s.Name = TXT_SERIES_AB & breakAB & TXT_CLOSE
                    s.XValues = Array(breakAB + BOUND_OFFSET, breakAB + BOUND_OFFSET)
                Else
                    s.Name = TXT_SERIES_BC & breakBC & TXT_CLOSE
                    s.XValues = Array(breakBC + BOUND_OFFSET, breakBC + BOUND_OFFSET)
                End If
                s.Values = Array(sc(m), maxCount)
                s.ChartType = xlXYScatterLinesNoMarkers
                s.Format.Line.Visible = msoTrue
                s.Format.Line.ForeColor.RGB = COLOR_BOUND
                s.Format.Line.DashStyle = msoLineDash
            Next seg
            ' 参考: 累積構成比のひざ
            If kneeRank > 0 Then
                Set s = ch.SeriesCollection.NewSeries
                s.Name = TXT_SERIES_KNEE & kneeRank & TXT_CLOSE
                s.XValues = Array(kneeRank, kneeRank)
                s.Values = Array(sc(m), maxCount)
                s.ChartType = xlXYScatterLinesNoMarkers
                s.Format.Line.Visible = msoTrue
                s.Format.Line.ForeColor.RGB = COLOR_KNEE
                s.Format.Line.DashStyle = msoLineRoundDot
            End If
        End If

        With ch.Axes(xlValue)
            .ScaleType = xlLogarithmic
            .LogBase = LOG_BASE
            .MinimumScale = AXIS_MIN_COUNT
            .HasTitle = True
            .AxisTitle.Text = TXT_AXIS_Y
        End With
        With ch.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = m + 1
            .HasTitle = True
            .AxisTitle.Text = TXT_AXIS_X
        End With
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionBottom
        ch.HasTitle = False
    End If

    ' ---- 印刷の設定と PDF ----
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(PAGE_MARGIN_IN)
        .RightMargin = Application.InchesToPoints(PAGE_MARGIN_IN)
        .TopMargin = Application.InchesToPoints(PAGE_MARGIN_IN)
        .BottomMargin = Application.InchesToPoints(PAGE_MARGIN_IN)
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(rowItems + n, COL_COUNT)).Address
    End With
    pdfPath = ""
    If Len(outFolder) > 0 Then
        pdfPath = outFolder & IIf(Right$(outFolder, 1) = "\", "", "\") & PDF_PREFIX & periodLabel & PDF_EXT
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                               IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
    CreateAbcReport = ok

    LogSuccess MODULE_NAME, "CreateAbcReport", "Classified " & n & " items (A=" & classItems(1) & ", B=" & classItems(2) & _
               ", C=" & classItems(3) & ", zero=" & zeroN & "); boundaries=" & breakAB & "/" & breakBC & "; pdf=" & pdfPath
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CreateAbcReport", Err.Number, Err.Description, Erl, _
             "n=" & n & "; m=" & m & "; i=" & i & "; seg=" & seg & "; r=" & r & "; periodLabel=" & periodLabel
    CreateAbcReport = False
End Function
