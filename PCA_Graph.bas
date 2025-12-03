Attribute VB_Name = "PCA_Graph"
Option Explicit

' ==============================================================================
' PCAグラフ作成モジュール
' 概要: KNN分類結果を受け取り、PCAによる次元圧縮と可視化を行う
' ==============================================================================

' ==============================================================================
' PCA実行とグラフ描画のメイン処理
' ==============================================================================
Public Sub CreatePCAGraph(ByVal X As Variant, ByVal y As Variant, _
                          ByVal x_query As Variant, ByVal predicted_label As String)
    Dim m As Long, p As Long
    Dim means As Variant, stds As Variant
    Dim Z() As Double ' 標準化データ (Double配列で高速化)
    Dim Cov() As Double
    Dim eigVal As Variant, eigVec As Variant
    Dim PC1() As Double, PC2() As Double, PC3() As Double
    Dim queryPC1 As Double, queryPC2 As Double, queryPC3 As Double
    Dim i As Long, j As Long, k As Long
    Dim sumVal As Double
    Dim use3D As Boolean
    
    m = UBound(X, 1)
    p = UBound(X, 2)
    
    ' 3次元計算が可能か判定
    If p >= 3 Then use3D = True Else use3D = False
    
    ' 1. データの標準化 (Standardization)
    '    KNNの距離計算と整合させるため、平均0・分散1に変換
    calc_mean_std X, means, stds
    ReDim Z(1 To m, 1 To p)
    
    For i = 1 To m
        For j = 1 To p
            Z(i, j) = (to_double(X(i, j)) - means(j)) / stds(j)
        Next j
    Next i
    
    ' 2. 分散共分散行列の計算: C = (Z^T * Z) / m
    ReDim Cov(1 To p, 1 To p)
    For i = 1 To p
        For j = 1 To p
            sumVal = 0#
            For k = 1 To m
                sumVal = sumVal + Z(k, i) * Z(k, j)
            Next k
            Cov(i, j) = sumVal / m
        Next j
    Next i
    
    ' 3. 固有値分解 (Jacobi法)
    Jacobi Cov, eigVal, eigVec
    
    ' 固有値の降順ソート
    SortEigen eigVal, eigVec
    
    ' 4. 主成分への射影 (PC1, PC2, PC3)
    ReDim PC1(1 To m)
    ReDim PC2(1 To m)
    If use3D Then ReDim PC3(1 To m)
    
    For i = 1 To m
        Dim val1 As Double, val2 As Double, val3 As Double
        val1 = 0#: val2 = 0#: val3 = 0#
        For j = 1 To p
            val1 = val1 + Z(i, j) * eigVec(j, 1)
            val2 = val2 + Z(i, j) * eigVec(j, 2)
            If use3D Then val3 = val3 + Z(i, j) * eigVec(j, 3)
        Next j
        PC1(i) = val1
        PC2(i) = val2
        If use3D Then PC3(i) = val3
    Next i
    
    ' --- 軸の向きの調整 (Flip Sign) ---
    ' 視認性向上のため、PC1/PC2/PC3の平均が負なら符号を反転させる
    If PC1(1) < 0 Then
        For i = 1 To m: PC1(i) = -PC1(i): Next i
        For j = 1 To p: eigVec(j, 1) = -eigVec(j, 1): Next j
    End If
    
    If PC2(1) < 0 Then
        For i = 1 To m: PC2(i) = -PC2(i): Next i
        For j = 1 To p: eigVec(j, 2) = -eigVec(j, 2): Next j
    End If
    
    If use3D Then
        If PC3(1) < 0 Then
            For i = 1 To m: PC3(i) = -PC3(i): Next i
            For j = 1 To p: eigVec(j, 3) = -eigVec(j, 3): Next j
        End If
    End If
    
    ' 5. クエリデータの射影
    Dim zq() As Double
    ReDim zq(1 To p)
    
    ' クエリも同様に標準化
    For j = 1 To p
        zq(j) = (to_double(x_query(j)) - means(j)) / stds(j)
    Next j
    
    queryPC1 = 0#: queryPC2 = 0#: queryPC3 = 0#
    For j = 1 To p
        queryPC1 = queryPC1 + zq(j) * eigVec(j, 1)
        queryPC2 = queryPC2 + zq(j) * eigVec(j, 2)
        If use3D Then queryPC3 = queryPC3 + zq(j) * eigVec(j, 3)
    Next j
    
    ' 6. シート出力とグラフ作成
    OutputAndChart PC1, PC2, PC3, y, queryPC1, queryPC2, queryPC3, predicted_label, use3D
End Sub

' ==============================================================================
' シートへのデータ出力と散布図作成
' ==============================================================================
Private Sub OutputAndChart(PC1() As Double, PC2() As Double, PC3() As Double, y As Variant, _
                           q1 As Double, q2 As Double, q3 As Double, qLabel As String, is3D As Boolean)
    Dim ws As Worksheet
    Dim m As Long
    Dim i As Long
    Dim outputData() As Variant
    Dim startCol As Long
    Dim chartObj As ChartObject
    Dim colCount As Long
    
    Set ws = ActiveSheet
    m = UBound(PC1)
    
    ' --- 1. シート初期化 ---
    ws.Cells.Clear
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' --- 2. データ出力 (配列で一括書き込みにより高速化) ---
    startCol = 1
    If is3D Then colCount = 4 Else colCount = 3
    
    ReDim outputData(1 To m + 2, 1 To colCount) ' +2はヘッダとクエリ行分
    
    ' ヘッダ
    outputData(1, 1) = "PC1"
    outputData(1, 2) = "PC2"
    If is3D Then
        outputData(1, 3) = "PC3"
        outputData(1, 4) = "Label"
    Else
        outputData(1, 3) = "Label"
    End If
    
    ' 学習データ
    For i = 1 To m
        outputData(i + 1, 1) = PC1(i)
        outputData(i + 1, 2) = PC2(i)
        If is3D Then
            outputData(i + 1, 3) = PC3(i)
            outputData(i + 1, 4) = y(i, 1)
        Else
            outputData(i + 1, 3) = y(i, 1)
        End If
    Next i
    
    ' クエリデータ (末尾に追加)
    Dim qRow As Long: qRow = m + 2
    outputData(qRow, 1) = q1
    outputData(qRow, 2) = q2
    If is3D Then
        outputData(qRow, 3) = q3
        outputData(qRow, 4) = "Query(" & qLabel & ")"
    Else
        outputData(qRow, 3) = "Query(" & qLabel & ")"
    End If
    
    ' セルへ一括出力
    ws.Range(ws.Cells(1, startCol), ws.Cells(qRow, startCol + colCount - 1)).Value = outputData
    
    ' --- 3. データのソート (系列作成のためラベル順に並べる) ---
    ' クエリ行(最終行)を除いてソート
    Dim sortRange As Range
    Dim labelColIdx As Long
    If is3D Then labelColIdx = startCol + 3 Else labelColIdx = startCol + 2
    
    Set sortRange = ws.Range(ws.Cells(1, startCol), ws.Cells(m + 1, startCol + colCount - 1))
    sortRange.Sort Key1:=ws.Cells(1, labelColIdx), Order1:=xlAscending, Header:=xlYes
    
    ' --- 4. グラフ作成 ---
    ' Excel標準では3D散布図がないため、PC1 vs PC2 の2D散布図を作成する
    ' (PC3はデータとして出力されているので、ユーザーが必要に応じて利用可能)
    
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(1, startCol + colCount + 1).Left, Top:=50, Width:=450, Height:=350)
    
    With chartObj.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        If is3D Then
            .ChartTitle.Text = "PCA Result (PC1 vs PC2) [3D Data Available]"
        Else
            .ChartTitle.Text = "PCA Result (PC1 vs PC2)"
        End If
        
        ' 既存系列の削除
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' ラベルごとに系列を追加 (連続領域を利用)
        Dim r As Long
        Dim currentLabel As String, nextLabel As String
        Dim startRow As Long
        
        startRow = 2
        currentLabel = ws.Cells(startRow, labelColIdx).Value
        
        For r = 2 To m + 1
            If r < m + 1 Then
                nextLabel = ws.Cells(r + 1, labelColIdx).Value
            Else
                nextLabel = ""
            End If
            
            If nextLabel <> currentLabel Then
                Dim s As Series
                Set s = .SeriesCollection.NewSeries
                s.Name = currentLabel
                s.XValues = ws.Range(ws.Cells(startRow, startCol), ws.Cells(r, startCol))
                s.Values = ws.Range(ws.Cells(startRow, startCol + 1), ws.Cells(r, startCol + 1))
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = 7
                
                currentLabel = nextLabel
                startRow = r + 1
            End If
        Next r
        
        ' クエリデータの系列追加
        Dim sQuery As Series
        Set sQuery = .SeriesCollection.NewSeries
        sQuery.Name = "Query: " & qLabel
        sQuery.XValues = ws.Cells(qRow, startCol)
        sQuery.Values = ws.Cells(qRow, startCol + 1)
        
        ' クエリ点のスタイル設定 (赤丸)
        sQuery.MarkerStyle = xlMarkerStyleCircle
        sQuery.MarkerSize = 10
        sQuery.MarkerForegroundColor = RGB(255, 0, 0)
        sQuery.MarkerBackgroundColor = RGB(255, 0, 0)
        
        ' 軸ラベル
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "PC1"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "PC2"
    End With
End Sub

' ==============================================================================
' Jacobi法による固有値・固有ベクトル計算
' 入力: 対称行列 A (N x N)
' 出力: 固有値 eigVal, 固有ベクトル eigVec
' ==============================================================================
Private Sub Jacobi(ByVal A As Variant, ByRef eigVal As Variant, ByRef eigVec As Variant)
    Dim n As Long
    n = UBound(A, 1)
    
    ReDim eigVec(1 To n, 1 To n)
    Dim i As Long, j As Long
    
    ' 単位行列で初期化
    For i = 1 To n
        For j = 1 To n
            If i = j Then eigVec(i, j) = 1# Else eigVec(i, j) = 0#
        Next j
    Next i
    
    Dim maxIter As Long: maxIter = 100
    Dim iter As Long
    Dim eps As Double: eps = 1E-15
    
    For iter = 1 To maxIter
        ' 非対角成分の最大値を探索
        Dim maxVal As Double: maxVal = 0#
        Dim p As Long, q As Long
        p = 1: q = 2
        
        For i = 1 To n - 1
            For j = i + 1 To n
                If Abs(A(i, j)) > maxVal Then
                    maxVal = Abs(A(i, j))
                    p = i: q = j
                End If
            Next j
        Next i
        
        If maxVal < eps Then Exit For
        
        ' 回転角の計算
        Dim theta As Double
        Dim App As Double, Aqq As Double, Apq As Double
        App = A(p, p)
        Aqq = A(q, q)
        Apq = A(p, q)
        
        If Abs(App - Aqq) < eps Then
            theta = 3.14159265358979 / 4# * Sgn(Apq)
        Else
            theta = 0.5 * Atn(2# * Apq / (App - Aqq))
        End If
        
        Dim c As Double, s As Double
        c = Cos(theta)
        s = Sin(theta)
        
        ' 行列 A の更新 (相似変換)
        Dim App_new As Double, Aqq_new As Double
        App_new = c * c * App + s * s * Aqq + 2# * s * c * Apq
        Aqq_new = s * s * App + c * c * Aqq - 2# * s * c * Apq
        A(p, q) = 0#: A(q, p) = 0#
        A(p, p) = App_new: A(q, q) = Aqq_new
        
        Dim Api As Double, Aqi As Double
        For i = 1 To n
            If i <> p And i <> q Then
                Api = A(p, i)
                Aqi = A(q, i)
                A(p, i) = c * Api + s * Aqi
                A(i, p) = A(p, i)
                A(q, i) = -s * Api + c * Aqi
                A(i, q) = A(q, i)
            End If
        Next i
        
        ' 固有ベクトルの更新
        Dim Vpi As Double, Vqi As Double
        For i = 1 To n
            Vpi = eigVec(i, p)
            Vqi = eigVec(i, q)
            eigVec(i, p) = c * Vpi + s * Vqi
            eigVec(i, q) = -s * Vpi + c * Vqi
        Next i
    Next iter
    
    ' 対角成分を固有値として抽出
    ReDim eigVal(1 To n)
    For i = 1 To n
        eigVal(i) = A(i, i)
    Next i
End Sub

' ==============================================================================
' 固有値・固有ベクトルのソート (降順)
' ==============================================================================
Private Sub SortEigen(ByRef eigVal As Variant, ByRef eigVec As Variant)
    Dim n As Long
    n = UBound(eigVal)
    Dim i As Long, j As Long, k As Long
    Dim tempVal As Double
    Dim tempVec As Double
    
    For i = 1 To n - 1
        For j = i + 1 To n
            If eigVal(j) > eigVal(i) Then
                ' 固有値の交換
                tempVal = eigVal(i)
                eigVal(i) = eigVal(j)
                eigVal(j) = tempVal
                
                ' 固有ベクトルの交換 (列単位)
                For k = 1 To UBound(eigVec, 1)
                    tempVec = eigVec(k, i)
                    eigVec(k, i) = eigVec(k, j)
                    eigVec(k, j) = tempVec
                Next k
            End If
        Next j
    Next i
End Sub
