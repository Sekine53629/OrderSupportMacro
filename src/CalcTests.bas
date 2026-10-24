Attribute VB_Name = "CalcTests"
Option Explicit

'======================================================================
' CalcTests モジュール
' AbcClassifier と OrderCalc の計算を、架空のデータで確かめる。
' RunAllTests は結果を "PASS ..." / "FAIL ..." の行で返す (tests/run-tests.ps1 から呼ぶ)。
' 値はすべてコードの中で作る架空のもので、実データは使わない。
'
' 依存: ErrorLogger.bas, AbcClassifier.bas, OrderCalc.bas, AbcReport.bas
'======================================================================

Private Const MODULE_NAME As String = "CalcTests"

' ---- ABC 判定のテスト用ランキング (A 40 / B 120 / C 300 品目) ----
Private Const A_ITEMS As Long = 40
Private Const B_ITEMS As Long = 120
Private Const C_ITEMS_3 As Long = 60               ' C のうち処方回数 3 の品目
Private Const C_ITEMS_2 As Long = 100              ' C のうち処方回数 2 の品目
Private Const C_ITEMS_1 As Long = 140              ' C のうち処方回数 1 の品目
Private Const A_TOP_COUNT As Double = 900          ' A の先頭の処方回数
Private Const A_SLOPE As Double = 0.005            ' A の log(処方回数) の 1 順位あたりの減り方 (ゆるやか)
Private Const B_SLOPE As Double = 0.03             ' B の減り方 (急)
Private Const C_COUNT_3 As Double = 3
Private Const C_COUNT_2 As Double = 2
Private Const C_COUNT_1 As Double = 1
Private Const MIN_CAND_COUNT As Double = 3         ' 折れ点の候補にする最小の処方回数
Private Const MIN_SEG_POINTS As Long = 5
Private Const BOUNDARY_TOL As Long = 3             ' 境界の順位の許容のずれ
Private Const FEW_ITEMS As Long = 8                ' 少なすぎて判定できない品目数

' ---- 発注点のテスト ----
Private Const LEAD_TIME As Double = 2
Private Const STOCK_DAYS_A As Double = 5
Private Const STOCK_DAYS_B As Double = 2
Private Const UPPER_DAYS_A As Double = 10
Private Const UPPER_DAYS_B As Double = 5
Private Const REORDER_ITEMS As Long = 5
Private Const SAFETY_DAYS As Double = 2
Private Const REQUIRED_ITEMS As Long = 7

' ---- 期待する理由の文言 ----
Private Const EXPECT_C_SPLIT As String = "C品目／分割発注"
Private Const EXPECT_C_UNREG As String = "C品目／マスタ未登録"

' ---- 分類レポートのテスト ----
Private Const ZERO_ITEMS As Long = 20              ' 処方回数 0 の品目
Private Const ROUND_HALF As Double = 0.5          ' 処方回数を整数に丸める
Private Const TEST_CODE_PREFIX As String = "TEST"
Private Const TEST_CODE_FORMAT As String = "0000"
Private Const TEST_NAME_PREFIX As String = "テスト薬"
Private Const TEST_SPEC As String = "10mg"
Private Const TEST_PERIOD As String = "TEST"
Private Const SHT_REPORT As String = "分類レポート"

Private Const RES_PASS As String = "PASS "
Private Const RES_FAIL As String = "FAIL "
Private Const RES_SEP As String = ": "


'----------------------------------------------------------------------
' RunAllTests: すべてのテストを実行し、結果の行をまとめて返す
'----------------------------------------------------------------------
Public Function RunAllTests() As String
    On Error GoTo ErrHandler
    Dim out As String

    out = out & TestAbcBoundaries()
    out = out & TestReorderPoints()
    out = out & TestRequiredQty()
    RunAllTests = out

    LogSuccess MODULE_NAME, "RunAllTests", "Ran all calculation tests"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "RunAllTests", Err.Number, Err.Description, Erl, ""
    RunAllTests = out & RES_FAIL & "RunAllTests" & RES_SEP & Err.Description & vbCrLf
End Function

'----------------------------------------------------------------------
' TestAbcBoundaries: 3 つの区間を持つランキングで、A/B・B/C の境界が求まるか
'----------------------------------------------------------------------
Private Function TestAbcBoundaries() As String
    On Error GoTo ErrHandler
    Dim counts() As Double
    Dim n As Long
    Dim k As Long
    Dim lastA As Double
    Dim breakAB As Long
    Dim breakBC As Long
    Dim kneeRank As Long
    Dim sse As Double
    Dim ok As Boolean
    Dim out As String
    Dim few() As Double

    n = A_ITEMS + B_ITEMS + C_ITEMS_3 + C_ITEMS_2 + C_ITEMS_1
    ReDim counts(1 To n)
    For k = 1 To A_ITEMS
        counts(k) = A_TOP_COUNT * Exp(-A_SLOPE * (k - 1))
    Next k
    lastA = counts(A_ITEMS)
    For k = 1 To B_ITEMS
        counts(A_ITEMS + k) = lastA * Exp(-B_SLOPE * k)
    Next k
    For k = 1 To C_ITEMS_3 + C_ITEMS_2 + C_ITEMS_1
        If k <= C_ITEMS_3 Then
            counts(A_ITEMS + B_ITEMS + k) = C_COUNT_3
        ElseIf k <= C_ITEMS_3 + C_ITEMS_2 Then
            counts(A_ITEMS + B_ITEMS + k) = C_COUNT_2
        Else
            counts(A_ITEMS + B_ITEMS + k) = C_COUNT_1
        End If
    Next k

    ok = AbcClassifier.FindBoundaries(counts, n, MIN_CAND_COUNT, MIN_SEG_POINTS, breakAB, breakBC, kneeRank, sse)
    If ok And Abs(breakAB - A_ITEMS) <= BOUNDARY_TOL And Abs(breakBC - (A_ITEMS + B_ITEMS)) <= BOUNDARY_TOL Then
        out = out & RES_PASS & "abc boundaries (AB=" & breakAB & ", BC=" & breakBC & ")" & vbCrLf
    Else
        out = out & RES_FAIL & "abc boundaries" & RES_SEP & "ok=" & ok & " AB=" & breakAB & " (expect " & A_ITEMS & _
              ") BC=" & breakBC & " (expect " & (A_ITEMS + B_ITEMS) & ")" & vbCrLf
    End If
    If kneeRank > 0 And kneeRank < n Then
        out = out & RES_PASS & "abc knee reference (rank " & kneeRank & ")" & vbCrLf
    Else
        out = out & RES_FAIL & "abc knee reference" & RES_SEP & "kneeRank=" & kneeRank & vbCrLf
    End If

    ' 品目が少なすぎるときは判定しない
    ReDim few(1 To FEW_ITEMS)
    For k = 1 To FEW_ITEMS
        few(k) = FEW_ITEMS - k + 1
    Next k
    If Not AbcClassifier.FindBoundaries(few, FEW_ITEMS, MIN_CAND_COUNT, MIN_SEG_POINTS, breakAB, breakBC, kneeRank, sse) Then
        out = out & RES_PASS & "abc too few items returns False" & vbCrLf
    Else
        out = out & RES_FAIL & "abc too few items" & RES_SEP & "returned True" & vbCrLf
    End If
    TestAbcBoundaries = out

    LogSuccess MODULE_NAME, "TestAbcBoundaries", "Tested ABC boundaries on " & n & " fictitious items"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "TestAbcBoundaries", Err.Number, Err.Description, Erl, "n=" & n & "; k=" & k
    TestAbcBoundaries = out & RES_FAIL & "TestAbcBoundaries" & RES_SEP & Err.Description & vbCrLf
End Function

'----------------------------------------------------------------------
' TestReorderPoints: 発注点と在庫上限
'   1: A 平均 10            -> 発注点 10x(2+5)=70、上限 10x10=100
'   2: B 平均 3.3           -> 発注点 3.3x(2+2)=13.2 -> 14、上限 3.3x5=16.5 -> 17
'   3: A 平均 10、保管上限 50 -> 上限 50、発注点 70 は上限にそろえて 50
'   4: C                    -> 0 / 0
'   5: A でもフラグ品        -> 0 / 0
'----------------------------------------------------------------------
Private Function TestReorderPoints() As String
    On Error GoTo ErrHandler
    Dim avgDaily(1 To REORDER_ITEMS) As Double
    Dim itemClass(1 To REORDER_ITEMS) As String
    Dim isFlag(1 To REORDER_ITEMS) As Boolean
    Dim storageLimit(1 To REORDER_ITEMS) As Double
    Dim reorderPoint() As Double
    Dim stockUpper() As Double
    Dim expectRp As Variant
    Dim expectUp As Variant
    Dim i As Long
    Dim out As String
    Dim bad As String

    avgDaily(1) = 10: itemClass(1) = OrderCalc.CLASS_A
    avgDaily(2) = 3.3: itemClass(2) = OrderCalc.CLASS_B
    avgDaily(3) = 10: itemClass(3) = OrderCalc.CLASS_A: storageLimit(3) = 50
    avgDaily(4) = 4: itemClass(4) = OrderCalc.CLASS_C
    avgDaily(5) = 10: itemClass(5) = OrderCalc.CLASS_A: isFlag(5) = True
    expectRp = Array(70, 14, 50, 0, 0)
    expectUp = Array(100, 17, 50, 0, 0)

    OrderCalc.CalcReorderPoints avgDaily, itemClass, isFlag, storageLimit, REORDER_ITEMS, _
                                LEAD_TIME, STOCK_DAYS_A, STOCK_DAYS_B, UPPER_DAYS_A, UPPER_DAYS_B, reorderPoint, stockUpper
    For i = 1 To REORDER_ITEMS
        If reorderPoint(i) <> expectRp(i - 1) Or stockUpper(i) <> expectUp(i - 1) Then
            bad = bad & " item" & i & "=" & reorderPoint(i) & "/" & stockUpper(i) & " (expect " & expectRp(i - 1) & "/" & expectUp(i - 1) & ")"
        End If
    Next i
    If Len(bad) = 0 Then
        out = RES_PASS & "reorder points and stock upper limits" & vbCrLf
    Else
        out = RES_FAIL & "reorder points" & RES_SEP & bad & vbCrLf
    End If
    TestReorderPoints = out

    LogSuccess MODULE_NAME, "TestReorderPoints", "Tested reorder points on " & REORDER_ITEMS & " fictitious items"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "TestReorderPoints", Err.Number, Err.Description, Erl, "i=" & i
    TestReorderPoints = out & RES_FAIL & "TestReorderPoints" & RES_SEP & Err.Description & vbCrLf
End Function

'----------------------------------------------------------------------
' TestRequiredQty: 手入力発注の必要数と理由 (安全在庫日数 2)
'   1: C 予定 25 在庫 3 平均 1 包装 10            -> 25+2-3=24 -> 30、理由 C
'   2: C 同上で在庫上限 20                        -> 余地 17 -> 10、理由 C+分割発注
'   3: C 予定 0 在庫 5                            -> 発注なし、リストに載せない
'   4: フラグ品 (A) 予定 0 在庫 5                 -> 0、理由 要判断 (残す)
'   5: 未登録 予定 7 包装 0 (1 として扱う)         -> 7、理由 C+マスタ未登録
'   6: A 予定 50 在庫 10 発注残 20                -> 0、理由 不足警告
'   7: A 予定 10 在庫 20                          -> 0、載せない
'----------------------------------------------------------------------
Private Function TestRequiredQty() As String
    On Error GoTo ErrHandler
    Dim planned(1 To REQUIRED_ITEMS) As Double
    Dim stock(1 To REQUIRED_ITEMS) As Double
    Dim backorder(1 To REQUIRED_ITEMS) As Double
    Dim avgDaily(1 To REQUIRED_ITEMS) As Double
    Dim packSize(1 To REQUIRED_ITEMS) As Double
    Dim stockUpper(1 To REQUIRED_ITEMS) As Double
    Dim itemClass(1 To REQUIRED_ITEMS) As String
    Dim isFlag(1 To REQUIRED_ITEMS) As Boolean
    Dim isUnregistered(1 To REQUIRED_ITEMS) As Boolean
    Dim required() As Double
    Dim reason() As Long
    Dim expectQty As Variant
    Dim expectReason As Variant
    Dim texts As Variant
    Dim i As Long
    Dim out As String
    Dim bad As String

    planned(1) = 25: stock(1) = 3: avgDaily(1) = 1: packSize(1) = 10: itemClass(1) = OrderCalc.CLASS_C
    planned(2) = 25: stock(2) = 3: avgDaily(2) = 1: packSize(2) = 10: itemClass(2) = OrderCalc.CLASS_C: stockUpper(2) = 20
    planned(3) = 0: stock(3) = 5: packSize(3) = 10: itemClass(3) = OrderCalc.CLASS_C
    planned(4) = 0: stock(4) = 5: packSize(4) = 10: itemClass(4) = OrderCalc.CLASS_A: isFlag(4) = True
    planned(5) = 7: itemClass(5) = OrderCalc.CLASS_C: isUnregistered(5) = True
    planned(6) = 50: stock(6) = 10: backorder(6) = 20: itemClass(6) = OrderCalc.CLASS_A
    planned(7) = 10: stock(7) = 20: itemClass(7) = OrderCalc.CLASS_A
    expectQty = Array(30, 10, 0, 0, 7, 0, 0)
    expectReason = Array(OrderCalc.REASON_C_ITEM, OrderCalc.REASON_C_ITEM Or OrderCalc.REASON_SPLIT, OrderCalc.REASON_NONE, _
                         OrderCalc.REASON_FLAG, OrderCalc.REASON_C_ITEM Or OrderCalc.REASON_UNREGISTERED, _
                         OrderCalc.REASON_SHORTAGE, OrderCalc.REASON_NONE)

    OrderCalc.CalcRequiredQty planned, stock, backorder, avgDaily, packSize, stockUpper, itemClass, isFlag, isUnregistered, _
                              REQUIRED_ITEMS, SAFETY_DAYS, required, reason
    For i = 1 To REQUIRED_ITEMS
        If required(i) <> expectQty(i - 1) Or reason(i) <> expectReason(i - 1) Then
            bad = bad & " item" & i & "=" & required(i) & "/" & reason(i) & " (expect " & expectQty(i - 1) & "/" & expectReason(i - 1) & ")"
        End If
    Next i
    If Len(bad) = 0 Then
        out = RES_PASS & "required quantities and reasons" & vbCrLf
    Else
        out = RES_FAIL & "required quantities" & RES_SEP & bad & vbCrLf
    End If

    ' 理由の文言
    texts = OrderCalc.ReasonTexts()
    If texts(OrderCalc.REASON_C_ITEM Or OrderCalc.REASON_SPLIT) = EXPECT_C_SPLIT And _
       texts(OrderCalc.REASON_C_ITEM Or OrderCalc.REASON_UNREGISTERED) = EXPECT_C_UNREG Then
        out = out & RES_PASS & "reason texts" & vbCrLf
    Else
        out = out & RES_FAIL & "reason texts" & vbCrLf
    End If
    TestRequiredQty = out

    LogSuccess MODULE_NAME, "TestRequiredQty", "Tested required quantities on " & REQUIRED_ITEMS & " fictitious items"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "TestRequiredQty", Err.Number, Err.Description, Erl, "i=" & i
    TestRequiredQty = out & RES_FAIL & "TestRequiredQty" & RES_SEP & Err.Description & vbCrLf
End Function

'----------------------------------------------------------------------
' RunReportTest: 架空のランキングで分類レポート (シート・グラフ・PDF) を作れるか
'   outFolder  PDF を書くフォルダ (tests/out)
'   A 40 / B 120 / C 300 品目 (TestAbcBoundaries と同じ並び) に、処方回数 0 の品目 20 を足す。
'   並び順は崩して渡し、区分が元の並び順で返ることも確かめる
'----------------------------------------------------------------------
Public Function RunReportTest(ByVal outFolder As String) As String
    On Error GoTo ErrHandler
    Dim n As Long
    Dim k As Long
    Dim idx As Long
    Dim lastA As Double
    Dim cnt As Double
    Dim codes() As String
    Dim names() As String
    Dim specs() As String
    Dim counts() As Double
    Dim itemClass() As String
    Dim pdfPath As String
    Dim ok As Boolean
    Dim ws As Worksheet
    Dim chartN As Long
    Dim nA As Long
    Dim nB As Long
    Dim nC As Long
    Dim out As String

    n = A_ITEMS + B_ITEMS + C_ITEMS_3 + C_ITEMS_2 + C_ITEMS_1 + ZERO_ITEMS
    ReDim codes(1 To n)
    ReDim names(1 To n)
    ReDim specs(1 To n)
    ReDim counts(1 To n)
    lastA = A_TOP_COUNT * Exp(-A_SLOPE * (A_ITEMS - 1))
    For k = 1 To n
        If k <= A_ITEMS Then
            cnt = A_TOP_COUNT * Exp(-A_SLOPE * (k - 1))
        ElseIf k <= A_ITEMS + B_ITEMS Then
            cnt = lastA * Exp(-B_SLOPE * (k - A_ITEMS))
        ElseIf k <= A_ITEMS + B_ITEMS + C_ITEMS_3 Then
            cnt = C_COUNT_3
        ElseIf k <= A_ITEMS + B_ITEMS + C_ITEMS_3 + C_ITEMS_2 Then
            cnt = C_COUNT_2
        ElseIf k <= A_ITEMS + B_ITEMS + C_ITEMS_3 + C_ITEMS_2 + C_ITEMS_1 Then
            cnt = C_COUNT_1
        Else
            cnt = 0
        End If
        ' 並び順を崩す (後ろから詰める)
        idx = n - k + 1
        codes(idx) = TEST_CODE_PREFIX & Format(k, TEST_CODE_FORMAT)
        names(idx) = TEST_NAME_PREFIX & k
        specs(idx) = TEST_SPEC
        counts(idx) = Int(cnt + ROUND_HALF)
    Next k

    ok = AbcReport.CreateAbcReport(codes, names, specs, counts, n, TEST_PERIOD, MIN_CAND_COUNT, MIN_SEG_POINTS, _
                                   outFolder, itemClass, pdfPath)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHT_REPORT Then chartN = ws.ChartObjects.Count
    Next ws
    For k = 1 To n
        Select Case itemClass(k)
            Case OrderCalc.CLASS_A: nA = nA + 1
            Case OrderCalc.CLASS_B: nB = nB + 1
            Case Else: nC = nC + 1
        End Select
    Next k

    If ok And chartN = 1 And Len(Dir$(pdfPath)) > 0 And Abs(nA - A_ITEMS) <= BOUNDARY_TOL _
       And Abs(nB - B_ITEMS) <= BOUNDARY_TOL And nA + nB + nC = n Then
        out = RES_PASS & "abc report (A=" & nA & ", B=" & nB & ", C=" & nC & ", chart, pdf)" & vbCrLf
    Else
        out = RES_FAIL & "abc report" & RES_SEP & "ok=" & ok & " charts=" & chartN & " pdf=" & pdfPath & _
              " A=" & nA & " B=" & nB & " C=" & nC & vbCrLf
    End If
    ' 区分は元の並び順で返る: 最後の要素は処方回数が最も多い品目 (A)
    If itemClass(n) = OrderCalc.CLASS_A And itemClass(1) = OrderCalc.CLASS_C Then
        out = out & RES_PASS & "abc report keeps input order" & vbCrLf
    Else
        out = out & RES_FAIL & "abc report keeps input order" & RES_SEP & itemClass(n) & "/" & itemClass(1) & vbCrLf
    End If
    RunReportTest = out

    LogSuccess MODULE_NAME, "RunReportTest", "Built ABC report for " & n & " fictitious items; pdf=" & pdfPath
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "RunReportTest", Err.Number, Err.Description, Erl, "k=" & k & "; outFolder=" & outFolder
    RunReportTest = out & RES_FAIL & "RunReportTest" & RES_SEP & Err.Description & vbCrLf
End Function
