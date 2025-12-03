Attribute VB_Name = "KNN"
Option Explicit
' ==============================================================================
' KNN分類メインモジュール
' ==============================================================================

' --- 定数設定 ---
Private Const K_NEIGHBORS As Long = 5       ' 近傍数 K
Private Const EPSILON As Double = 1E-09     ' ゼロ除算回避用の微小値

' ==============================================================================
' エントリポイント: KNN分類実行
' 概要: CSV読み込み -> 特徴量選択 -> 学習 -> ユーザー入力クエリの分類 -> PCA可視化
' ==============================================================================
Public Sub main_generic_knn_classify()
    Dim csvPath As Variant
    Dim allData As Variant
    Dim headers As Variant, feature_cols As Variant
    Dim label_col As Long
    Dim X As Variant, y As Variant
    Dim x_query() As Double ' Double配列に変更して高速化
    Dim means As Variant, stds As Variant
    Dim top_idx As Variant
    Dim decided_species As String, decided_conf As Double
    
    Dim targetColName As String
    Dim featureNameList As String
    Dim inputStr As String
    Dim strParts() As String
    Dim i As Long, p As Long, k As Long

    ' 高速化: 画面更新と自動計算を停止
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    On Error GoTo ErrHandler

    ' 1. 学習データ(CSV)の選択
    csvPath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "学習データ(CSV)を選択してください")
    If csvPath = False Then GoTo Cleanup

    ' 2. CSV読み込み (UTF-8対応)
    allData = LoadCSV(CStr(csvPath))
    If IsEmpty(allData) Then
        MsgBox "CSVファイルの読み込みに失敗しました。", vbCritical
        GoTo Cleanup
    End If

    ' 3. ヘッダ解析とターゲット列の特定
    headers = GetRowFromArray(allData, 1)
    
    targetColName = InputBox("分類対象（正解ラベル）の列名を入力してください。", "ターゲット列指定", "species")
    If targetColName = "" Then GoTo Cleanup
    
    label_col = FindLabelColIndex(headers, targetColName)
    If label_col <= 0 Then
        MsgBox "列 '" & targetColName & "' が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' ターゲット列以外を特徴量として抽出
    feature_cols = GetFeatureColIndices(headers, label_col)
    If IsEmpty(feature_cols) Then
        MsgBox "特徴量列が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' 4. データセット(X, y)の構築
    parse_dataset_vectors allData, 2, feature_cols, label_col, X, y
    If IsEmpty(X) Then
        MsgBox "学習データが空です。", vbCritical
        GoTo Cleanup
    End If

    ' 5. ユーザーからのクエリ入力
    p = UBound(feature_cols) - LBound(feature_cols) + 1
    featureNameList = JoinFeatures(headers, feature_cols) ' ヘルパー関数で結合

    inputStr = InputBox("特徴量をカンマ区切りで入力してください" & vbCrLf & _
                        "項目数: " & p & vbCrLf & _
                        "項目順: " & featureNameList, _
                        "KNN分類実行", "6.4,3.2,4.5,1.5")
    If inputStr = "" Then GoTo Cleanup

    strParts = Split(inputStr, ",")
    If UBound(strParts) + 1 <> p Then
        MsgBox "入力値の数が特徴量数(" & p & ")と一致しません。", vbExclamation
        GoTo Cleanup
    End If

    ' クエリベクトル作成 (Double型へ変換)
    ReDim x_query(1 To p)
    For i = 0 To UBound(strParts)
        x_query(i + 1) = to_double(Trim(strParts(i)))
    Next i

    ' 6. 学習と推論
    ' 標準化用の統計量算出
    calc_mean_std X, means, stds

    ' KNN実行 (距離計算 -> ソート -> 投票)
    top_idx = topk_neighbors(X, x_query, means, stds, K_NEIGHBORS)
    knn_final_decision X, y, x_query, top_idx, means, stds, decided_species, decided_conf

    ' 結果表示
    MsgBox "分類結果: " & decided_species & vbCrLf & _
           "信頼度: " & Format(decided_conf, "0.0%"), vbInformation, "KNN推論結果"

    ' 7. PCAグラフ作成 (可視化モジュール呼び出し)
    CreatePCAGraph X, y, x_query, decided_species

Cleanup:
    ' アプリケーション設定の復元
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    Exit Sub

ErrHandler:
    MsgBox "実行時エラー: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==============================================================================
' CSV読み込み (ADODB.Stream使用)
' 概要: UTF-8のCSVファイルを読み込み、2次元配列として返す
' ==============================================================================
Public Function LoadCSV(ByVal filePath As String) As Variant
    Dim adoSt As Object
    Dim fileContent As String
    Dim lines() As String
    Dim r As Long, c As Long
    Dim cols() As String
    Dim maxCol As Long
    Dim rowCount As Long
    Dim data() As Variant
    
    Set adoSt = CreateObject("ADODB.Stream")
    With adoSt
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        fileContent = .ReadText
        .Close
    End With
    Set adoSt = Nothing
    
    ' 改行コード統一と行分割
    fileContent = Replace(Replace(fileContent, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(fileContent, vbLf)
    
    ' 有効行数のカウント (末尾の空行除外)
    rowCount = UBound(lines) + 1
    Do While rowCount > 0
        If Trim(lines(rowCount - 1)) <> "" Then Exit Do
        rowCount = rowCount - 1
    Loop
    
    If rowCount <= 0 Then Exit Function
    
    ' 列数決定と配列確保
    cols = Split(lines(0), ",")
    maxCol = UBound(cols) + 1
    ReDim data(1 To rowCount, 1 To maxCol)
    
    ' データ格納
    For r = 0 To rowCount - 1
        cols = Split(lines(r), ",")
        For c = 0 To UBound(cols)
            If c < maxCol Then data(r + 1, c + 1) = Trim(cols(c))
        Next c
    Next r
    
    LoadCSV = data
End Function

' ==============================================================================
' 配列操作ヘルパー: 指定行の抽出
' ==============================================================================
Public Function GetRowFromArray(ByVal data As Variant, ByVal rowIdx As Long) As Variant
    Dim c As Long, cols As Long
    cols = UBound(data, 2)
    Dim res() As Variant
    ReDim res(1 To cols)
    For c = 1 To cols
        res(c) = data(rowIdx, c)
    Next c
    GetRowFromArray = res
End Function

' ==============================================================================
' データセット構築: 特徴量行列Xとラベルベクトルyの生成
' ==============================================================================
Public Sub parse_dataset_vectors(ByVal allData As Variant, ByVal startRow As Long, _
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
    Dim srcRow As Long
    
    ' ループ内で頻繁にアクセスする配列は型指定変数を介すと若干高速
    rowIdx = 0
    For srcRow = startRow To totalRows
        rowIdx = rowIdx + 1
        For j = 1 To p
            X(rowIdx, j) = allData(srcRow, CLng(feature_cols(j)))
        Next j
        y(rowIdx, 1) = allData(srcRow, label_col)
    Next srcRow
End Sub

' ==============================================================================
' KNN推論: 最終決定 (重み付き多数決)
' ==============================================================================
Private Sub knn_final_decision( _
        ByVal X As Variant, ByVal y As Variant, ByVal xq As Variant, _
        ByVal top_idx As Variant, ByVal means As Variant, ByVal stds As Variant, _
        ByRef species_out As String, ByRef conf_out As Double)
    
    Dim k As Long: k = UBound(top_idx) - LBound(top_idx) + 1
    Dim weights() As Double: ReDim weights(1 To k)
    Dim labels() As String: ReDim labels(1 To k)
    Dim i As Long, idx As Long, d As Double

    ' 逆距離重みの計算
    For i = 1 To k
        idx = CLng(top_idx(i))
        d = zdist(X, xq, means, stds, idx)
        weights(i) = 1# / (d + EPSILON)
        labels(i) = CStr(y(idx, 1))
    Next i

    ' クラスごとの重み集計
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To k
        If dict.Exists(labels(i)) Then
            dict(labels(i)) = dict(labels(i)) + weights(i)
        Else
            dict.Add labels(i), weights(i)
        End If
    Next i

    ' 最尤クラスの決定
    Dim best_label As String
    Dim best_sum As Double: best_sum = -1#
    Dim total_sum As Double: total_sum = 0#
    Dim key As Variant

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

' ==============================================================================
' ヘルパー: ラベル列のインデックス検索 (大文字小文字無視)
' ==============================================================================
Public Function FindLabelColIndex(ByVal headers As Variant, ByVal labelName As String) As Long
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

' ==============================================================================
' ヘルパー: 特徴量列インデックスの取得 (ラベル列以外)
' ==============================================================================
Public Function GetFeatureColIndices(ByVal headers As Variant, ByVal label_col As Long) As Variant
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

' ==============================================================================
' 統計計算: 平均と標準偏差
' ==============================================================================
Public Sub calc_mean_std(ByVal X As Variant, ByRef means As Variant, ByRef stds As Variant)
    Dim m As Long: m = UBound(X, 1)
    Dim p As Long: p = UBound(X, 2)
    ReDim means(1 To p): ReDim stds(1 To p)
    
    Dim j As Long, i As Long
    Dim s As Double, ss As Double, v As Double
    Dim mu As Double, sigma As Double

    For j = 1 To p
        s = 0#: ss = 0#
        For i = 1 To m
            v = to_double(X(i, j))
            s = s + v
            ss = ss + v * v
        Next i
        mu = s / m
        ' 母分散計算: (Σx^2 / N) - μ^2
        sigma = Sqr(Application.Max(0#, (ss / m) - (mu * mu)))
        If sigma < EPSILON Then sigma = EPSILON
        means(j) = mu
        stds(j) = sigma
    Next j
End Sub

' ==============================================================================
' 距離計算: 標準化ユークリッド距離
' ==============================================================================
Private Function zdist(ByVal X As Variant, ByVal xq As Variant, _
                       ByVal means As Variant, ByVal stds As Variant, _
                       ByVal rowIdx As Long) As Double
    Dim p As Long: p = UBound(X, 2)
    Dim j As Long
    Dim d As Double, v As Double, q As Double
    
    d = 0#
    For j = 1 To p
        ' 事前に計算した平均・標準偏差で標準化しながら距離加算
        v = (to_double(X(rowIdx, j)) - means(j)) / stds(j)
        q = (to_double(xq(j)) - means(j)) / stds(j)
        d = d + (v - q) * (v - q)
    Next j
    zdist = Sqr(d)
End Function

' ==============================================================================
' 近傍探索: 距離計算と上位K件の抽出
' ==============================================================================
Private Function topk_neighbors(ByVal X As Variant, ByVal xq As Variant, _
                                ByVal means As Variant, ByVal stds As Variant, _
                                ByVal k As Long) As Variant
    Dim m As Long: m = UBound(X, 1)
    If k > m Then k = m
    
    ' (index, distance) のペア配列
    Dim pairs() As Variant: ReDim pairs(1 To m, 1 To 2)
    Dim i As Long
    
    For i = 1 To m
        pairs(i, 1) = i
        pairs(i, 2) = zdist(X, xq, means, stds, i)
    Next i

    ' 距離でソート
    quicksort_pairs pairs, 1, m
    
    ' 上位K件のインデックスのみ抽出
    Dim res() As Long: ReDim res(1 To k)
    For i = 1 To k
        res(i) = CLng(pairs(i, 1))
    Next i
    topk_neighbors = res
End Function

' ==============================================================================
' ソート: クイックソート (距離昇順)
' ==============================================================================
Private Sub quicksort_pairs(ByRef a() As Variant, ByVal l As Long, ByVal r As Long)
    Dim i As Long, j As Long
    Dim p As Double, t0 As Variant, t1 As Variant
    
    i = l: j = r
    p = CDbl(a((l + r) \ 2, 2)) ' ピボット
    
    Do While i <= j
        Do While CDbl(a(i, 2)) < p: i = i + 1: Loop
        Do While CDbl(a(j, 2)) > p: j = j - 1: Loop
        If i <= j Then
            ' Swap
            t0 = a(i, 1): t1 = a(i, 2)
            a(i, 1) = a(j, 1): a(i, 2) = a(j, 2)
            a(j, 1) = t0: a(j, 2) = t1
            i = i + 1: j = j - 1
        End If
    Loop
    
    If l < j Then quicksort_pairs a, l, j
    If i < r Then quicksort_pairs a, i, r
End Sub

' ==============================================================================
' ユーティリティ: 数値変換 (非数値は0)
' ==============================================================================
Public Function to_double(ByVal v As Variant) As Double
    If IsNumeric(v) Then to_double = CDbl(v) Else to_double = 0#
End Function

' ==============================================================================
' ユーティリティ: 特徴量名の結合 (表示用)
' ==============================================================================
Private Function JoinFeatures(headers As Variant, feature_cols As Variant) As String
    Dim k As Long, s As String
    s = ""
    For k = LBound(feature_cols) To UBound(feature_cols)
        If s <> "" Then s = s & ", "
        s = s & headers(CLng(feature_cols(k)))
    Next k
    JoinFeatures = s
End Function
