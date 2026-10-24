Attribute VB_Name = "AbcClassifier"
Option Explicit

'======================================================================
' AbcClassifier モジュール
' 処方回数のランキングから A / B / C の境界を求める (要件定義書 4-1、F-02)。
'   - 順位 (横軸) と log(処方回数) (縦軸) の点に、3 本の直線 (区間ごとの回帰直線) を当てはめる
'   - 折れ点 2 か所の組み合わせを総当たりし、二乗誤差の合計が最小の組み合わせを選ぶ
'     区間の二乗誤差は累積和から一定時間で求めるので、品目数 n に対して n^2 / 2 通りを試す
'   - 1 つ目の折れ点 (A の最後の順位) を A/B の境界、2 つ目 (B の最後の順位) を B/C の境界とする
'   - 処方回数が minCandCount 未満の品目は折れ点の候補にしない (裾の 1 回・2 回の段々を避ける)
'   - 参考値として、累積構成比の曲線のひざ (始点と終点を結ぶ直線から最も離れた順位) も返す
' 入力は処方回数 0 を除き、多い順に並べた配列とする (並べ替えは呼び出し元で行う)。
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "AbcClassifier"

Private Const SEGMENT_COUNT As Long = 3            ' 当てはめる直線の数 (A / B / C)
Private Const MIN_SEG_POINTS_FLOOR As Long = 2     ' 1 つの区間に必要な最小の点の数 (直線を引くのに 2 点)


'----------------------------------------------------------------------
' FindBoundaries: 処方回数のランキングから A/B と B/C の境界の順位を求める
'   counts()       処方回数 (1 始まり、多い順、すべて 0 より大きい)
'   n              品目数
'   minCandCount   折れ点の候補にする最小の処方回数
'   minSegPoints   1 つの区間に必要な最小の品目数
'   breakAB        A の最後の順位 (この順位までが A)
'   breakBC        B の最後の順位 (この順位までが B、それより後ろが C)
'   kneeRank       累積構成比のひざの順位 (参考値)
'   bestSse        最小の二乗誤差の合計
'   戻り値         境界を求められたら True (品目が少ない、候補が無いときは False)
'----------------------------------------------------------------------
Public Function FindBoundaries(ByRef counts() As Double, ByVal n As Long, ByVal minCandCount As Double, _
                               ByVal minSegPoints As Long, ByRef breakAB As Long, ByRef breakBC As Long, _
                               ByRef kneeRank As Long, ByRef bestSse As Double) As Boolean
    On Error GoTo ErrHandler
    Dim sx() As Double
    Dim sy() As Double
    Dim sxx() As Double
    Dim sxy() As Double
    Dim syy() As Double
    Dim k As Long
    Dim y As Double
    Dim i As Long
    Dim j As Long
    Dim a As Long
    Dim b As Long
    Dim seg As Long
    Dim m As Double
    Dim tx As Double
    Dim ty As Double
    Dim txx As Double
    Dim txy As Double
    Dim tyy As Double
    Dim cxx As Double
    Dim cxy As Double
    Dim cyy As Double
    Dim segSse As Double
    Dim total As Double
    Dim found As Boolean
    Dim cum As Double
    Dim sumAll As Double
    Dim dist As Double
    Dim bestDist As Double
    Dim fx As Double
    Dim fy As Double
    Dim y1 As Double

    breakAB = 0
    breakBC = 0
    kneeRank = 0
    bestSse = 0
    If minSegPoints < MIN_SEG_POINTS_FLOOR Then minSegPoints = MIN_SEG_POINTS_FLOOR
    If n < minSegPoints * SEGMENT_COUNT Then
        FindBoundaries = False
        LogSuccess MODULE_NAME, "FindBoundaries", "Too few items to fit 3 segments: n=" & n
        Exit Function
    End If

    ' ---- 累積和 (x = 順位、y = log(処方回数)) ----
    ReDim sx(0 To n)
    ReDim sy(0 To n)
    ReDim sxx(0 To n)
    ReDim sxy(0 To n)
    ReDim syy(0 To n)
    For k = 1 To n
        y = Log(counts(k))
        sx(k) = sx(k - 1) + k
        sy(k) = sy(k - 1) + y
        sxx(k) = sxx(k - 1) + CDbl(k) * k
        sxy(k) = sxy(k - 1) + k * y
        syy(k) = syy(k - 1) + y * y
        sumAll = sumAll + counts(k)
    Next k

    ' ---- 折れ点 i (A の最後) と j (B の最後) を総当たりする ----
    For i = minSegPoints To n - 2 * minSegPoints
        If counts(i) >= minCandCount Then
            For j = i + minSegPoints To n - minSegPoints
                If counts(j) >= minCandCount Then
                    ' 3 つの区間 [1, i] [i+1, j] [j+1, n] の、回帰直線からの二乗誤差の合計
                    total = 0
                    For seg = 1 To SEGMENT_COUNT
                        Select Case seg
                            Case 1
                                a = 1
                                b = i
                            Case 2
                                a = i + 1
                                b = j
                            Case Else
                                a = j + 1
                                b = n
                        End Select
                        m = b - a + 1
                        tx = sx(b) - sx(a - 1)
                        ty = sy(b) - sy(a - 1)
                        txx = sxx(b) - sxx(a - 1)
                        txy = sxy(b) - sxy(a - 1)
                        tyy = syy(b) - syy(a - 1)
                        cxx = txx - tx * tx / m
                        cxy = txy - tx * ty / m
                        cyy = tyy - ty * ty / m
                        If cxx > 0 Then
                            segSse = cyy - cxy * cxy / cxx
                        Else
                            segSse = cyy
                        End If
                        If segSse < 0 Then segSse = 0      ' 丸め誤差でわずかに負になるのを防ぐ
                        total = total + segSse
                    Next seg
                    If Not found Or total < bestSse Then
                        found = True
                        bestSse = total
                        breakAB = i
                        breakBC = j
                    End If
                End If
            Next j
        End If
    Next i

    ' ---- 参考値: 累積構成比の曲線のひざ ----
    ' 横軸 = 順位 / n、縦軸 = 累積構成比。始点 (1/n, 先頭の構成比) と終点 (1, 1) を結ぶ直線から最も離れた点
    If sumAll > 0 Then
        y1 = counts(1) / sumAll
        cum = 0
        For k = 1 To n
            cum = cum + counts(k)
            fx = k / n
            fy = cum / sumAll
            ' 直線の上側にある距離 (定数倍は比較に影響しないので省く)
            dist = (fy - y1) * (1 - 1 / n) - (fx - 1 / n) * (1 - y1)
            If dist > bestDist Then
                bestDist = dist
                kneeRank = k
            End If
        Next k
    End If

    FindBoundaries = found
    LogSuccess MODULE_NAME, "FindBoundaries", "n=" & n & "; breakAB=" & breakAB & "; breakBC=" & breakBC & _
               "; kneeRank=" & kneeRank & "; sse=" & Format(bestSse, "0.0000") & "; minCandCount=" & minCandCount
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FindBoundaries", Err.Number, Err.Description, Erl, _
             "n=" & n & "; i=" & i & "; j=" & j & "; k=" & k
    FindBoundaries = False
End Function
