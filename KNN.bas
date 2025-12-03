Option Explicit

' ======== 設定 ========
Private Const K_NEIGHBORS As Long = 5                   ' 近傍K
Private Const EPSILON As Double = 0.000000001           ' 0除算回避用

'==========================================
' エントリポイント
'【概要】CSV を読み込み、特徴量列・ラベル列を抽出し、KNN 分類を実行する全体制御処理
'        ユーザー入力値を含むクエリベクトルを受け取り、最終ラベルと信頼度を表示する
'==========================================
Public Sub main_generic_knn_classify()
    Dim csvPath As Variant
    Dim allData As Variant
    Dim headers As Variant, feature_cols As Variant
    Dim label_col As Long
    Dim X As Variant, y As Variant
    Dim x_query As Variant
    Dim means As Variant, stds As Variant
    Dim top_idx As Variant
    Dim decided_species As String, decided_conf As Double
    
    Dim targetColName As String
    Dim featureNameList As String
    Dim inputStr As String
    Dim strParts() As String
    Dim i As Long, p As Long, k As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrHandler

    ' --- CSVファイル選択 ---
    csvPath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "学習データ(CSV)を選択してください")
    If csvPath = False Then GoTo Cleanup

    ' --- CSV読み込み ---
    allData = LoadCSV(CStr(csvPath))
    If IsEmpty(allData) Then
        MsgBox "CSVファイルの読み込みに失敗しました。", vbCritical
        GoTo Cleanup
    End If

    ' ヘッダ取得（1行目）
    headers = GetRowFromArray(allData, 1)
    
    ' --- ターゲット列名をユーザーに要求 ---
    targetColName = InputBox("分類対象（正解ラベル）の列名を入力してください。", "ターゲット列指定", "species")
    If targetColName = "" Then GoTo Cleanup
    
    ' ラベル列インデックス検索
    label_col = FindLabelColIndex(headers, targetColName)
    If label_col <= 0 Then
        MsgBox "データに '" & targetColName & "' 列が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' 特徴量列インデックス取得
    feature_cols = GetFeatureColIndices(headers, label_col)
    If IsEmpty(feature_cols) Then
        MsgBox "特徴量列が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' データセット(X, y)作成
    parse_dataset_vectors allData, 2, feature_cols, label_col, X, y
    If IsEmpty(X) Then
        MsgBox "学習データが空です。", vbCritical
        GoTo Cleanup
    End If

    ' --- 入力プロンプトの文字列準備 ---
    p = UBound(feature_cols) - LBound(feature_cols) + 1
    featureNameList = ""
    For k = 1 To p
        If featureNameList <> "" Then featureNameList = featureNameList & ", "
        featureNameList = featureNameList & headers(CLng(feature_cols(k)))
    Next k

    ' --- 特徴量のユーザー入力 ---
    inputStr = InputBox("特徴量をカンマ区切りで入力してください" & vbCrLf & _
                        "項目数: " & p & vbCrLf & _
                        "項目順: " & featureNameList, _
                        "KNN分類実行", "6.4,3.2,4.5,1.5")
    If inputStr = "" Then GoTo Cleanup

    strParts = Split(inputStr, ",")
    
    If UBound(strParts) + 1 <> p Then
        MsgBox "入力された値の数が特徴量の数(" & p & ")と一致しません。" & vbCrLf & _
               "期待される項目: " & featureNameList, vbExclamation
        GoTo Cleanup
    End If

    ' クエリベクトル作成
    ReDim x_query(1 To p)
    For i = 0 To UBound(strParts)
        x_query(i + 1) = to_double(Trim(strParts(i)))
    Next i

    ' 標準化用統計量算出
    calc_mean_std X, means, stds

    ' KNN 実行
    top_idx = topk_neighbors(X, x_query, means, stds, K_NEIGHBORS)
    knn_final_decision X, y, x_query, top_idx, means, stds, decided_species, decided_conf

    ' 結果表示
    MsgBox "分類結果: " & decided_species & vbCrLf & _
           "信頼度: " & Format(decided_conf, "0.0%"), vbInformation, "KNN推論結果"

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "実行時エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical, "エラー"
    Resume Cleanup
End Sub

'==========================================
' CSV読み込み
'【概要】指定した CSV ファイルを丸ごと読み込み、
'        行×列の 1-based 二次元配列として返す
'        ・改行コード統一
'        ・空行の除外
'        ・単純なカンマ Split による列分割
'==========================================
Private Function LoadCSV(ByVal filePath As String) As Variant
    Dim fso As Object, ts As Object
    Dim fileContent As String
    Dim lines() As String
    Dim r As Long, c As Long
    Dim cols() As String
    Dim maxCol As Long
    Dim rowCount As Long
    Dim data() As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    
    Set ts = fso.OpenTextFile(filePath, 1)
    If ts.AtEndOfStream Then
        ts.Close
        Exit Function
    End If
    
    fileContent = ts.ReadAll
    ts.Close
    
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)
    lines = Split(fileContent, vbLf)
    
    rowCount = UBound(lines) + 1
    Do While rowCount > 0
        If Trim(lines(rowCount - 1)) <> "" Then Exit Do
        rowCount = rowCount - 1
    Loop
    
    If rowCount <= 0 Then Exit Function
    
    cols = Split(lines(0), ",")
    maxCol = UBound(cols) + 1
    
    ReDim data(1 To rowCount, 1 To maxCol)
    
    For r = 0 To rowCount - 1
        cols = Split(lines(r), ",")
        For c = 0 To UBound(cols)
            If c < maxCol Then
                data(r + 1, c + 1) = Trim(cols(c))
            End If
        Next c
    Next r
    
    LoadCSV = data
End Function

'==========================================
' 配列の指定行を1次元配列として返す
'【概要】2次元配列 data の rowIdx 行目だけを切り出し、
'        1次元配列に変換して返す
'==========================================
Private Function GetRowFromArray(ByVal data As Variant, ByVal rowIdx As Long) As Variant
    Dim c As Long, cols As Long
    cols = UBound(data, 2)
    Dim res() As Variant
    ReDim res(1 To cols)
    For c = 1 To cols
        res(c) = data(rowIdx, c)
    Next c
    GetRowFromArray = res
End Function

'==========================================
' データセット(X, y)ベクトル作成
'【概要】全データ allData から
'        ・特徴量行列 X(m, p)
'        ・ラベルベクトル y(m, 1)
'        を抽出して生成する
'==========================================
Private Sub parse_dataset_vectors(ByVal allData As Variant, ByVal startRow As Long, _
                                  ByVal feature_cols As Variant, ByVal label_col As Long, _
                                  ByRef X As Variant, ByRef y As Variant)
    Dim totalRows As Long
    totalRows = UBound(allData, 1)
    
    If totalRows < startRow Then Exit Sub

    Dim m As Long, p As Long
    m = totalRows - startRow + 1
    p = UBound(feature_cols) - LBound(feature_cols) + 1
    
    ReDim X(1 To m, 1 To p)
    ReDim y(1 To m, 1 To 1)

    Dim i As Long, j As Long, rowIdx As Long
    rowIdx = 0
    For i = startRow To totalRows
        rowIdx = rowIdx + 1
        For j = 1 To p
            X(rowIdx, j) = allData(i, CLng(feature_cols(j)))
        Next j
        y(rowIdx, 1) = allData(i, label_col)
    Next i
End Sub

'==========================================
' KNN最終決定
'【概要】近傍 top_idx のラベル y に対し、
'        ・逆距離重みを計算
'        ・クラスごとの重み合計を算出
'        ・最大重みのラベルを予測クラスとする
'        ・重み比から信頼度を算出
'==========================================
Private Sub knn_final_decision( _
        ByVal X As Variant, ByVal y As Variant, ByVal xq As Variant, _
        ByVal top_idx As Variant, ByVal means As Variant, ByVal stds As Variant, _
        ByRef species_out As String, ByRef conf_out As Double)
    
    Dim k As Long: k = UBound(top_idx) - LBound(top_idx) + 1
    Dim weights() As Double: ReDim weights(1 To k)
    Dim labels() As String: ReDim labels(1 To k)
    Dim i As Long, idx As Long, d As Double

    For i = 1 To k
        idx = CLng(top_idx(i))
        d = zdist(X, xq, means, stds, idx)
        weights(i) = 1# / (d + EPSILON)
        labels(i) = CStr(y(idx, 1))
    Next i

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To k
        If Not dict.Exists(labels(i)) Then
            dict.Add labels(i), weights(i)
        Else
            dict(labels(i)) = dict(labels(i)) + weights(i)
        End If
    Next i

    Dim best_label As String: best_label = ""
    Dim best_sum As Double: best_sum = -1#
    Dim key As Variant, total_sum As Double: total_sum = 0#

    For Each key In dict.Keys
        If dict(key) > best_sum Then
            best_sum = dict(key)
            best_label = CStr(key)
        End If
        total_sum = total_sum + CDbl(dict(key))
    Next key

    If total_sum <= 0# Then total_sum = 1#
    species_out = best_label
    conf_out = best_sum / total_sum
End Sub

'==========================================
' ラベル列インデックス検索
'【概要】ヘッダ配列 headers の中から、
'        指定ラベル名 labelName と一致する列番号を返す
'        ※大文字小文字は無視して比較
'==========================================
Private Function FindLabelColIndex(ByVal headers As Variant, ByVal labelName As String) As Long
    Dim c As Long, n As Long
    If Not IsArray(headers) Then Exit Function
    n = UBound(headers)
    For c = 1 To n
        If StrComp(CStr(headers(c)), labelName, vbTextCompare) = 0 Then
            FindLabelColIndex = c
            Exit Function
        End If
    Next c
    FindLabelColIndex = 0
End Function

'==========================================
' 特徴量列インデックス生成
'【概要】ヘッダ行から、
'        ラベル列以外すべてを「特徴量列」としてインデックス配列化する
'==========================================
Private Function GetFeatureColIndices(ByVal headers As Variant, ByVal label_col As Long) As Variant
    Dim n As Long, i As Long, k As Long
    If Not IsArray(headers) Then Exit Function
    n = UBound(headers)
    If n <= 1 Then Exit Function

    Dim tmp() As Long: ReDim tmp(1 To n - 1)
    k = 0
    For i = 1 To n
        If i <> label_col Then
            k = k + 1
            tmp(k) = i
        End If
    Next i
    
    If k > 0 Then
        Dim res() As Long: ReDim res(1 To k)
        For i = 1 To k: res(i) = tmp(i): Next i
        GetFeatureColIndices = res
    End If
End Function

'==========================================
' 平均・標準偏差の算出
'【概要】特徴量行列 X の各列 p について、
'        ・平均値 μ
'        ・標準偏差 σ（母標準偏差）
'        を計算して返す
'==========================================
Private Sub calc_mean_std(ByVal X As Variant, ByRef means As Variant, ByRef stds As Variant)
    Dim m As Long: m = UBound(X, 1)
    Dim p As Long: p = UBound(X, 2)
    ReDim means(1 To p): ReDim stds(1 To p)
    
    Dim j As Long, i As Long
    Dim s As Double, ss As Double, mu As Double, sigma As Double

    For j = 1 To p
        s = 0#: ss = 0#
        For i = 1 To m
            Dim v As Double: v = to_double(X(i, j))
            s = s + v
            ss = ss + v * v
        Next i
        mu = s / m
        sigma = Sqr(Application.Max(0#, (ss / m) - (mu * mu)))
        If sigma < 0.000000001 Then sigma = 0.000000001
        means(j) = mu
        stds(j) = sigma
    Next j
End Sub

'==========================================
' 標準化距離（ユークリッド距離）の算出
'【概要】行 X(rowIdx, :) とクエリ xq の
'        標準化済みベクトル間のユークリッド距離を返す
'==========================================
Private Function zdist(ByVal X As Variant, ByVal xq As Variant, ByVal means As Variant, ByVal stds As Variant, ByVal rowIdx As Long) As Double
    Dim p As Long: p = UBound(X, 2)
    Dim j As Long, d As Double, v As Double, q As Double
    d = 0#
    For j = 1 To p
        v = (to_double(X(rowIdx, j)) - means(j)) / stds(j)
        q = (to_double(xq(j)) - means(j)) / stds(j)
        d = d + (v - q) * (v - q)
    Next j
    zdist = Sqr(d)
End Function

'==========================================
' 近傍K件の抽出
'【概要】全データとクエリの距離を計算し、
'        (index, distance) のペアを距離昇順でソートして
'        上位 K 件のデータインデックスを返す
'==========================================
Private Function topk_neighbors(ByVal X As Variant, ByVal xq As Variant, ByVal means As Variant, ByVal stds As Variant, ByVal k As Long) As Variant
    Dim m As Long: m = UBound(X, 1)
    If k > m Then k = m
    Dim pairs() As Variant: ReDim pairs(1 To m, 1 To 2)
    Dim i As Long
    
    For i = 1 To m
        pairs(i, 1) = i
        pairs(i, 2) = zdist(X, xq, means, stds, i)
    Next i

    quicksort_pairs pairs, 1, m
    
    Dim res() As Long: ReDim res(1 To k)
    For i = 1 To k
        res(i) = CLng(pairs(i, 1))
    Next i
    topk_neighbors = res
End Function

'==========================================
' ペア配列のクイックソート
'【概要】(index, distance) 形式の 2列配列を
'        距離列（2列目）で昇順ソートする
'==========================================
Private Sub quicksort_pairs(ByRef a() As Variant, ByVal l As Long, ByVal r As Long)
    Dim i As Long, j As Long
    Dim p As Double, t0 As Variant, t1 As Variant
    i = l: j = r
    p = CDbl(a((l + r) \ 2, 2))
    
    Do While i <= j
        Do While CDbl(a(i, 2)) < p: i = i + 1: Loop
        Do While CDbl(a(j, 2)) > p: j = j - 1: Loop
        If i <= j Then
            t0 = a(i, 1): t1 = a(i, 2)
            a(i, 1) = a(j, 1): a(i, 2) = a(j, 2)
            a(j, 1) = t0: a(j, 2) = t1
            i = i + 1: j = j - 1
        End If
    Loop
    If l < j Then quicksort_pairs a, l, j
    If i < r Then quicksort_pairs a, i, r
End Sub

'==========================================
' 数値変換
'【概要】値が数値なら CDbl により Double 化して返し、
'        数値でなければ 0 を返す簡易変換関数
'==========================================
Private Function to_double(ByVal v As Variant) As Double
    If IsNumeric(v) Then to_double = CDbl(v) Else to_double = 0#
End Function
