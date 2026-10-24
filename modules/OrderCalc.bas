Attribute VB_Name = "OrderCalc"
Option Explicit

'======================================================================
' OrderCalc モジュール
' 発注点・在庫上限 (要件定義書 4-2、F-06・F-07) と、手入力発注の必要数 (4-3、F-11〜F-13) を計算する。
' 品目ごとに関数を呼ぶと LogSuccess がそのたびにログを書き直すので、品目の配列をまとめて計算する。
'
'   発注点   = 1 日あたり平均使用量 x (リードタイム + 区分ごとの在庫日数)          (A/B のみ)
'   在庫上限 = min(1 日あたり平均使用量 x 区分ごとの上限日数, 保管上限)            (A/B のみ。保管上限 0 は上限なし)
'   必要数   = 予定使用量 + 安全在庫 - 現在庫 - 発注残                            (C・フラグ・マスタ未登録)
'              -> 包装単位で切り上げ -> 在庫上限で頭打ち -> 0 以下なら発注なし
'   安全在庫 = 1 日あたり平均使用量 x 安全在庫日数
'
' 区分は CLASS_A / CLASS_B / CLASS_C の文字列 ("A" "B" "C")。フラグ品は isFlag、マスタ未登録品は isUnregistered。
' 理由はビットの組み合わせで返す (REASON_*)。表示の文言の表は ReasonTexts で作る。
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "OrderCalc"

' ---- 区分 ----
Public Const CLASS_A As String = "A"
Public Const CLASS_B As String = "B"
Public Const CLASS_C As String = "C"

' ---- 理由 (ビット) ----
Public Const REASON_NONE As Long = 0
Public Const REASON_C_ITEM As Long = 1           ' C 品目
Public Const REASON_FLAG As Long = 2             ' 要判断 (フラグ品)
Public Const REASON_SPLIT As Long = 4            ' 分割発注 (在庫上限で頭打ち)
Public Const REASON_SHORTAGE As Long = 8         ' 不足警告 (A/B 品目。自動発注に任せたまま警告だけ出す)
Public Const REASON_UNREGISTERED As Long = 16    ' マスタ未登録 (C として扱う)
Private Const REASON_BITS As String = "1|2|4|8|16"
Private Const REASON_KINDS As Long = 5             ' 理由の種類の数 (REASON_BITS の要素数)

' ---- 理由の文言 ----
Private Const TXT_C_ITEM As String = "C品目"
Private Const TXT_FLAG As String = "要判断"
Private Const TXT_SPLIT As String = "分割発注"
Private Const TXT_SHORTAGE As String = "不足警告"
Private Const TXT_UNREGISTERED As String = "マスタ未登録"
Private Const TXT_SEP As String = "／"
Private Const LIST_SEP As String = "|"

Private Const MIN_PACK As Double = 1               ' 包装単位が未設定 (0 以下) のときに使う単位


'----------------------------------------------------------------------
' CalcReorderPoints: A/B 品目の発注点と在庫上限を計算する (C とフラグ品は 0)
'   avgDaily()      1 日あたり平均使用量
'   itemClass()     区分 ("A" "B" "C")
'   isFlag()        フラグ品なら True (自動発注の対象外 = 発注点 0)
'   storageLimit()  保管上限 (0 以下は上限なし)
'   reorderPoint()  発注点 (出力。小数は切り上げ)
'   stockUpper()    在庫上限 (出力。小数は切り上げ)
'----------------------------------------------------------------------
Public Sub CalcReorderPoints(ByRef avgDaily() As Double, ByRef itemClass() As String, ByRef isFlag() As Boolean, _
                               ByRef storageLimit() As Double, ByVal n As Long, _
                               ByVal leadTimeDays As Double, ByVal stockDaysA As Double, ByVal stockDaysB As Double, _
                               ByVal upperDaysA As Double, ByVal upperDaysB As Double, _
                               ByRef reorderPoint() As Double, ByRef stockUpper() As Double)
      On Error GoTo ErrHandler
      Dim i As Long
      Dim stockDays As Double
      Dim upperDays As Double
      Dim upperVal As Double
      Dim targetN As Long
      Dim errNo As Long
      Dim errDesc As String

10    ReDim reorderPoint(1 To n)
20    ReDim stockUpper(1 To n)
30    For i = 1 To n
40        If (itemClass(i) = CLASS_A Or itemClass(i) = CLASS_B) And Not isFlag(i) Then
50            If itemClass(i) = CLASS_A Then
60                stockDays = stockDaysA
70                upperDays = upperDaysA
              Else
80                stockDays = stockDaysB
90                upperDays = upperDaysB
              End If
              ' 発注点 = 平均使用量 x (リードタイム + 在庫日数)。端数は切り上げる
100           reorderPoint(i) = -Int(-(avgDaily(i) * (leadTimeDays + stockDays)))
              ' 在庫上限 = min(平均使用量 x 上限日数, 保管上限)
110           upperVal = -Int(-(avgDaily(i) * upperDays))
120           If storageLimit(i) > 0 And storageLimit(i) < upperVal Then upperVal = storageLimit(i)
130           stockUpper(i) = upperVal
              ' 発注点が在庫上限を超えるときは、在庫上限にそろえる (上限を超えて発注しない)
140           If reorderPoint(i) > stockUpper(i) Then reorderPoint(i) = stockUpper(i)
150           targetN = targetN + 1
          End If
      Next i

      LogSuccess MODULE_NAME, "CalcReorderPoints", "Calculated reorder points for " & targetN & " A/B items out of " & n
160   Exit Sub
ErrHandler:
      ' LogError は Err を消すので、先に控えておく
      errNo = Err.Number
      errDesc = Err.Description
      LogError MODULE_NAME, "CalcReorderPoints", errNo, errDesc, Erl, "n=" & n & "; i=" & i
      Err.Raise errNo, MODULE_NAME & ".CalcReorderPoints", errDesc
End Sub

'----------------------------------------------------------------------
' CalcRequiredQty: 手入力発注の必要数と理由を計算する
'   planned()        予定使用量 (次回納品までの来局予定の合計)
'   stock()          現在庫
'   backorder()      発注残
'   avgDaily()       1 日あたり平均使用量 (安全在庫に使う。マスタ未登録品は 0)
'   packSize()       包装単位 (0 以下は 1 として扱う)
'   stockUpper()     在庫上限 (0 以下は上限なし)。発注後の在庫 (現在庫 + 発注残 + 必要数) がこれを超えないようにする
'   itemClass()      区分。A/B は必要数を出さず、予定使用量が現在庫 + 発注残を上回るときだけ不足警告を付ける
'   isFlag()         フラグ品 (区分に関係なく、手入力発注の対象として要判断を付ける)
'   isUnregistered() マスタ未登録品 (C として計算し、マスタ未登録を付ける)
'   required()       必要数 (出力。包装単位の倍数)
'   reason()         理由 (出力。REASON_* の組み合わせ。REASON_NONE ならリストに載せない)
'----------------------------------------------------------------------
Public Sub CalcRequiredQty(ByRef planned() As Double, ByRef stock() As Double, ByRef backorder() As Double, _
                             ByRef avgDaily() As Double, ByRef packSize() As Double, ByRef stockUpper() As Double, _
                             ByRef itemClass() As String, ByRef isFlag() As Boolean, ByRef isUnregistered() As Boolean, _
                             ByVal n As Long, ByVal safetyDays As Double, _
                             ByRef required() As Double, ByRef reason() As Long)
      On Error GoTo ErrHandler
      Dim i As Long
      Dim pack As Double
      Dim raw As Double
      Dim qty As Double
      Dim room As Double
      Dim manual As Boolean
      Dim listedN As Long
      Dim errNo As Long
      Dim errDesc As String

10    ReDim required(1 To n)
20    ReDim reason(1 To n)
30    For i = 1 To n
40        manual = isFlag(i) Or isUnregistered(i) Or Not (itemClass(i) = CLASS_A Or itemClass(i) = CLASS_B)
50        If Not manual Then
              ' A/B 品目は自動発注に任せる。足りなくなりそうなときだけ警告する
60            If planned(i) > stock(i) + backorder(i) Then reason(i) = REASON_SHORTAGE
          Else
70            If isFlag(i) Then
80                reason(i) = REASON_FLAG
              Else
90                reason(i) = REASON_C_ITEM
              End If
100           If isUnregistered(i) Then reason(i) = reason(i) Or REASON_UNREGISTERED

110           pack = packSize(i)
120           If pack <= 0 Then pack = MIN_PACK
130           raw = planned(i) + avgDaily(i) * safetyDays - stock(i) - backorder(i)
140           If raw > 0 Then
                  ' 包装単位で切り上げる
150               qty = -Int(-raw / pack) * pack
                  ' 在庫上限で頭打ちにする (発注後の在庫が上限を超えない、包装単位の倍数まで)
160               If stockUpper(i) > 0 Then
170                   room = stockUpper(i) - stock(i) - backorder(i)
180                   If qty > room Then
190                       If room > 0 Then
200                           qty = Int(room / pack) * pack
                          Else
210                           qty = 0
                          End If
220                       reason(i) = reason(i) Or REASON_SPLIT
                      End If
                  End If
230               required(i) = qty
              Else
                  ' 0 以下なら発注なし。フラグ品・未登録品は要判断として残し、C 品目はリストに載せない
240               required(i) = 0
250               If Not isFlag(i) And Not isUnregistered(i) Then reason(i) = REASON_NONE
              End If
          End If
260       If reason(i) <> REASON_NONE Then listedN = listedN + 1
      Next i

      LogSuccess MODULE_NAME, "CalcRequiredQty", "Calculated required quantities for " & n & " items; listed=" & listedN & _
                 "; safetyDays=" & safetyDays
270   Exit Sub
ErrHandler:
      ' LogError は Err を消すので、先に控えておく
      errNo = Err.Number
      errDesc = Err.Description
      LogError MODULE_NAME, "CalcRequiredQty", errNo, errDesc, Erl, "n=" & n & "; i=" & i
      Err.Raise errNo, MODULE_NAME & ".CalcRequiredQty", errDesc
End Sub

'----------------------------------------------------------------------
' ReasonText: 理由のビットを表示用の文言にする (例: "C品目／分割発注")
'   発注リストの 1 行ごとに呼ぶと LogSuccess が重いので、文言の表 (texts) を一度だけ作り、
'   呼び出し元で reasonText(i) = texts(reason(i)) のように引く
'----------------------------------------------------------------------
Public Function ReasonTexts() As Variant
      On Error GoTo ErrHandler
      Dim texts() As String
      Dim code As Long
      Dim bits() As String
      Dim names(1 To REASON_KINDS) As String
      Dim k As Long
      Dim maxCode As Long

10    names(1) = TXT_C_ITEM
20    names(2) = TXT_FLAG
30    names(3) = TXT_SPLIT
40    names(4) = TXT_SHORTAGE
50    names(5) = TXT_UNREGISTERED
60    bits = Split(REASON_BITS, LIST_SEP)
70    maxCode = REASON_C_ITEM Or REASON_FLAG Or REASON_SPLIT Or REASON_SHORTAGE Or REASON_UNREGISTERED
80    ReDim texts(0 To maxCode)
90    For code = 0 To maxCode
100       For k = 0 To UBound(bits)
110           If (code And CLng(bits(k))) <> 0 Then
120               texts(code) = texts(code) & IIf(Len(texts(code)) > 0, TXT_SEP, "") & names(k + 1)
              End If
          Next k
      Next code
130   ReasonTexts = texts

      LogSuccess MODULE_NAME, "ReasonTexts", "Built " & (maxCode + 1) & " reason texts"
140   Exit Function
ErrHandler:
      LogError MODULE_NAME, "ReasonTexts", Err.Number, Err.Description, Erl, "code=" & code
      ReasonTexts = Empty
End Function
