Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices

''' <summary>
''' ヒストグラム差分アルゴリズムを実装するモジュールです。
''' </summary>
Public Module HistogramDiffModule

    ''' <summary>ヒストグラムに含める文字列の出現頻度の閾値。値が小さいほど、より固有な行を優先して一致を検出します。</summary>
    Public Const FREQUENCY As Integer = 64

    ''' <summary>編集操作の種類を表す列挙型</summary>
    Public Enum EditTypeEnum
        ''' <summary>一致（変更なし）</summary>
        Match
        ''' <summary>差分（置換）</summary>
        Diff
        ''' <summary>挿入</summary>
        Insert
        ''' <summary>削除</summary>
        Delete
    End Enum

    ''' <summary>
    ''' 編集ブロックを表す構造体。各ブロックは、編集操作の種類と、ソース側と宛先側の行番号範囲を保持します。
    ''' </summary>
    Public Structure EvaluationBlock
        ''' <summary>編集操作の種類</summary>
        Public ReadOnly Property EditType As EditTypeEnum
        ''' <summary>ソース側の行番号範囲と宛先側の行番号範囲を表すプロパティ。</summary>
        Public ReadOnly Property SourceStart As Integer
        ''' <summary>ソース側の行番号範囲と宛先側の行番号範囲を表すプロパティ。</summary>
        Public ReadOnly Property SourceEnd As Integer
        ''' <summary>ソース側の行番号範囲と宛先側の行番号範囲を表すプロパティ。</summary>
        Public ReadOnly Property DestinationStart As Integer
        ''' <summary>ソース側の行番号範囲と宛先側の行番号範囲を表すプロパティ。</summary>
        Public ReadOnly Property DestinationEnd As Integer

        ''' <summary>編集ブロックを初期化します。</summary>
        ''' <param name="editType">編集操作の種類</param>
        ''' <param name="srcStart">ソース側の開始行番号</param>
        ''' <param name="srcEnd">ソース側の終了行番号</param>
        ''' <param name="destStart">宛先側の開始行番号</param>
        ''' <param name="destEnd">宛先側の終了行番号</param>
        Public Sub New(editType As EditTypeEnum, srcStart As Integer, srcEnd As Integer,
                       destStart As Integer, destEnd As Integer)
            Me.EditType = editType
            Me.SourceStart = srcStart
            Me.SourceEnd = srcEnd
            Me.DestinationStart = destStart
            Me.DestinationEnd = destEnd
        End Sub
    End Structure

    ''' <summary>
    ''' ソースと宛先の文字列配列を比較し、編集ブロックのリストを返します。
    ''' </summary>
    ''' <param name="source">比較元の文字列配列</param>
    ''' <param name="destination">比較先の文字列配列</param>
    ''' <returns>編集ブロックのリスト</returns>
    Public Function Diff(source As IEnumerable(Of String), destination As IEnumerable(Of String)) As List(Of EvaluationBlock)
        ' 引数の検査
        If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))
        If destination Is Nothing Then Throw New ArgumentNullException(NameOf(destination))

        ' 文字列配列を行番号付きのLineInfo配列に変換
        Dim src = ConvertAll(source)
        Dim dest = ConvertAll(destination)

        ' ヒストグラムを作成
        Dim blocks As New List(Of EvaluationBlock)()
        PlotBestMatch(blocks, src, src.ChangeHistogram(), 0, src.Length, dest, dest.ChangeHistogram(), 0, dest.Length)
        For Each block In blocks
            Select Case block.EditType
                Case EditTypeEnum.Match
                    ' 一致ブロックはそのまま
                    Console.WriteLine($"Match: src({block.SourceStart}-{block.SourceEnd}) <-> dest({block.DestinationStart}-{block.DestinationEnd})")
                    For i = block.SourceStart To block.SourceEnd - 1
                        Console.WriteLine(src(i).Str)
                    Next
                Case EditTypeEnum.Insert
                    ' 挿入ブロックは、ソース側の行番号を-1にして、宛先側の行番号をそのままにする
                    Console.WriteLine($"Insert: dest({block.DestinationStart}-{block.DestinationEnd})")
                    For i = block.DestinationStart To block.DestinationEnd - 1
                        Console.WriteLine($"{dest(i).Line:00000} | {dest(i).Str}")
                    Next
                Case EditTypeEnum.Delete
                    ' 削除ブロックは、ソース側の行番号をそのままにして、宛先側の行番号を-1にする
                    Console.WriteLine($"Delete: src({block.SourceStart}-{block.SourceEnd})")
                    For i = block.SourceStart To block.SourceEnd - 1
                        Console.WriteLine($"{src(i).Line:00000} | {src(i).Str}")
                    Next
                Case EditTypeEnum.Diff
                    ' 差分ブロックは、ソース側と宛先側の行番号をそのままにする（必要に応じて編集内容を追加で保持することも可能）
            End Select
        Next
        Return blocks
    End Function

#Region "比較対象の事前処理"

    ''' <summary>文字列配列をLineInfo配列に変換します</summary>
    ''' <param name="arr">変換対象の文字列配列</param>
    ''' <returns>行番号付きのLineInfo配列</returns>
    Private Function ConvertAll(arr As IEnumerable(Of String)) As LineInfo()
        Dim result As New List(Of LineInfo)
        Dim index As Integer = 0
        For Each str As String In arr
            result.Add(New LineInfo(index, str))
            index += 1
        Next
        Return result.ToArray()
    End Function

    ''' <summary>行番号とその内容を表す構造体</summary>
    Public Structure LineInfo

        ''' <summary>元の配列内での行番号（0ベースのインデックス）</summary>
        Public ReadOnly Property Line As Integer

        ''' <summary>行の文字列内容</summary>
        Public ReadOnly Property Str As String

        ''' <summary>
        ''' 行番号と文字列内容を指定してLineInfo構造体を初期化します
        ''' </summary>
        ''' <param name="line">行番号</param>
        ''' <param name="str">行の文字列内容</param>
        Public Sub New(line As Integer, str As String)
            Me.Line = line
            Me.Str = str
        End Sub
    End Structure

#End Region

#Region "ヒストグラムの作成"

    ''' <summary>
    ''' 文字列とその出現行番号を保持する構造体。
    ''' </summary>
    Private Structure Histogram

        ''' <summary>対象文字列</summary>
        Public ReadOnly Property Str As String

        ''' <summary>出現行番号の配列（昇順）</summary>
        Public ReadOnly Property Lines As Integer()

        ''' <summary>Histogram を初期化します。</summary>
        ''' <param name="str">対象文字列</param>
        ''' <param name="lines">出現行番号の配列</param>
        Public Sub New(str As String, lines() As Integer)
            Me.Str = str
            Me.Lines = lines
        End Sub

    End Structure

    ''' <summary>
    ''' ソースと宛先の両方に出現する文字列の行番号情報をグループ化します。
    ''' </summary>
    Private NotInheritable Class GroupedHistogram
        ''' <summary>出現頻度。値が小さいほど固有性が高い。</summary>
        Public ReadOnly Property Occurrence As Long

        ''' <summary>グループ化された文字列</summary>
        Public ReadOnly Property Line As String

        ''' <summary>ソース側の出現行番号配列</summary>
        Public ReadOnly Property SourceLines As Integer()

        ''' <summary>宛先側の出現行番号配列</summary>
        Public ReadOnly Property DestinationLines As Integer()

        ''' <summary>
        ''' GroupedHistogramを初期化します。
        ''' </summary>
        ''' <param name="line">文字列</param>
        ''' <param name="sourceLines">ソース側の行番号配列</param>
        ''' <param name="destinationLines">宛先側の行番号配列</param>
        Public Sub New(line As String, sourceLines As Integer(), destinationLines As Integer())
            Me.Line = line
            Me.SourceLines = sourceLines
            Me.DestinationLines = destinationLines
            Me.Occurrence = sourceLines.Length * sourceLines.Length + destinationLines.Length * destinationLines.Length
        End Sub
    End Class

    ''' <summary>
    ''' LineInfo 配列からヒストグラムを生成します。
    ''' 各文字列の出現位置を記録し、最初の出現位置でソートして返します。
    ''' </summary>
    ''' <param name="line">ヒストグラム生成対象のLineInfo配列</param>
    ''' <returns>文字列ごとの出現位置情報を含むヒストグラムのリスト（最初の行番号でソート済み）</returns>
    <Extension()>
    Private Function ChangeHistogram(line As LineInfo()) As List(Of Histogram)
        Dim temp As New Dictionary(Of String, List(Of Integer))()

        ' 文字列ごとに行番号のリストを構築
        For i = 0 To line.Length - 1
            Dim str = line(i).Str
            Dim list As List(Of Integer) = Nothing

            If Not temp.TryGetValue(str, list) Then
                list = New List(Of Integer)()
                temp(str) = list
            End If

            list.Add(i)
        Next

        ' 文字列ごとの行番号リストをヒストグラム構造体に変換し、最初の行番号でソートして返す
        Dim result As New List(Of Histogram)()
        For Each kvp In temp
            If kvp.Value.Count > FREQUENCY Then
                result.Add(New Histogram(kvp.Key, kvp.Value.ToArray()))
            End If
        Next
        result.Sort(Function(a, b) a.Lines(0).CompareTo(b.Lines(0)))
        Return result
    End Function

    ''' <summary>
    ''' ヒストグラムを指定範囲でフィルタリングし、範囲内の行番号のみを抽出します。
    ''' </summary>
    ''' <param name="source">フィルタリング対象のヒストグラムリスト</param>
    ''' <param name="startIndex">開始インデックス（含む）</param>
    ''' <param name="endIndex">終了インデックス（含まない）</param>
    ''' <returns>
    ''' フィルタリング後のヒストグラムリストと、文字列と行番号配列の辞書を含むタプル
    ''' </returns>
    <Extension()>
    Private Function Analysis(source As List(Of Histogram),
                              startIndex As Integer,
                              endIndex As Integer) As (source As List(Of Histogram), analysis As Dictionary(Of String, Integer()))
        Dim hist As New List(Of Histogram)()
        Dim strs As New Dictionary(Of String, Integer())()
        For Each kvp In source
            ' ヒストグラムは行番号の昇順でソートされているため、範囲外のものはスキップできる
            If endIndex <= kvp.Lines(0) Then
                Exit For
            End If

            ' 範囲内に行番号が存在するヒストグラムのみを処理
            If startIndex <= kvp.Lines(kvp.Lines.Length - 1) AndAlso kvp.Lines(0) < endIndex Then
                If startIndex > kvp.Lines(0) OrElse kvp.Lines(kvp.Lines.Length - 1) >= endIndex Then
                    ' 範囲外の行番号を除外して新しい配列を作成
                    Dim st As Integer = 0
                    Do While st < kvp.Lines.Length AndAlso startIndex > kvp.Lines(st)
                        st += 1
                    Loop
                    Dim ed As Integer = kvp.Lines.Length - 1
                    Do While ed >= 0 AndAlso kvp.Lines(ed) >= endIndex
                        ed -= 1
                    Loop

                    If ed >= st Then
                        Dim arr = New Integer(ed - st) {}
                        For i As Integer = st To ed
                            arr(i - st) = kvp.Lines(i)
                        Next
                        hist.Add(kvp)
                        strs(kvp.Str) = arr
                    End If
                Else
                    ' 範囲内の行番号が全て含まれている場合は、そのまま追加
                    hist.Add(kvp)
                    strs(kvp.Str) = kvp.Lines
                End If
            End If
        Next
        Return (hist, strs)
    End Function

    ''' <summary>
    ''' ヒストグラム差分アルゴリズムを使用して、ソースと宛先の範囲間で最適な一致ブロックを再帰的に検出し、
    ''' 編集操作のリストを構築します。
    ''' </summary>
    ''' <param name="blocks">検出された編集ブロック（一致、挿入、削除）を格納するリスト</param>
    ''' <param name="srcLines">ソース側の行情報配列</param>
    ''' <param name="srcHistogram">ソース側のヒストグラム（文字列とその出現位置のマッピング）</param>
    ''' <param name="srcStart">ソース側の検索開始位置（0ベース、この位置を含む）</param>
    ''' <param name="srcEnd">ソース側の検索終了位置（0ベース、この位置を含まない）</param>
    ''' <param name="destLines">宛先側の行情報配列</param>
    ''' <param name="destHistogram">宛先側のヒストグラム（文字列とその出現位置のマッピング）</param>
    ''' <param name="destStart">宛先側の検索開始位置（0ベース、この位置を含む）</param>
    ''' <param name="destEnd">宛先側の検索終了位置（0ベース、この位置を含まない）</param>
    ''' <remarks>
    ''' このメソッドは以下の処理フローで動作します：
    ''' 
    ''' 1. **ベースケース処理**
    '''    - ソース範囲が空の場合：宛先範囲全体を挿入（Insert）ブロックとして追加
    '''    - 宛先範囲が空の場合：ソース範囲全体を削除（Delete）ブロックとして追加
    ''' 
    ''' 2. **ヒストグラム分析**
    '''    - 指定範囲内のヒストグラムを分析し、該当する行情報のみを抽出
    ''' 
    ''' 3. **最適一致の検出**
    '''    - FindBestMatchメソッドを使用して、出現頻度が最も低い（固有性が高い）一致ブロックを検出
    ''' 
    ''' 4. **再帰的分割処理**
    '''    - 一致ブロックが見つかった場合：
    '''      a. 一致ブロック前の範囲を再帰的に処理（左側の差分）
    '''      b. 一致ブロックを結果リストに追加（Match）
    '''      c. 一致ブロック後の範囲を再帰的に処理（右側の差分）
    '''    - 一致ブロックが見つからない場合：
    '''      a. ソース範囲全体を削除（Delete）ブロックとして追加
    '''      b. 宛先範囲全体を挿入（Insert）ブロックとして追加
    ''' 
    ''' このアルゴリズムは、Gitで使用されるヒストグラム差分アルゴリズムに基づいており、
    ''' 固有性の高い行を優先して一致を検出することで、より直感的な差分結果を生成します。
    ''' </remarks>
    Private Sub PlotBestMatch(blocks As List(Of EvaluationBlock),
                              srcLines As LineInfo(),
                              srcHistogram As List(Of Histogram),
                              srcStart As Integer,
                              srcEnd As Integer,
                              destLines As LineInfo(),
                              destHistogram As List(Of Histogram),
                              destStart As Integer,
                              destEnd As Integer)
        ' ヒストグラムが空の場合は、ベースケースと同様に処理
        If srcHistogram.Count <= 0 OrElse destHistogram.Count <= 0 Then
            If srcStart < srcEnd Then
                blocks.Add(New EvaluationBlock(EditTypeEnum.Delete, srcStart, srcEnd, -1, -1))
            End If
            If destStart < destEnd Then
                blocks.Add(New EvaluationBlock(EditTypeEnum.Insert, -1, -1, destStart, destEnd))
            End If
            Return
        End If

        ' ベースケース：どちらかの範囲が空の場合は、残りの範囲を挿入または削除として処理
        If srcStart >= srcEnd AndAlso destStart >= destEnd Then
            Return ' 何もしない
        End If

        If srcStart >= srcEnd Then
            blocks.Add(New EvaluationBlock(EditTypeEnum.Insert, -1, -1, destStart, destEnd))
            Return
        End If

        If destStart >= destEnd Then
            blocks.Add(New EvaluationBlock(EditTypeEnum.Delete, srcStart, srcEnd, -1, -1))
            Return
        End If

        ' ヒストグラム分析：指定範囲内のヒストグラムを抽出
        Dim srcAnalysis = srcHistogram.Analysis(srcStart, srcEnd)
        Dim destAnalysis = destHistogram.Analysis(destStart, destEnd)

        ' 最適一致の検出：出現頻度が最も低い（固有性が高い）一致ブロックを検出
        Dim best = FindBestMatch(srcAnalysis.analysis, srcLines, srcStart, srcEnd, destAnalysis.analysis, destLines, destStart, destEnd)

        If best.find Then
            ' 一致ブロックが見つかった場合は、前後の範囲を再帰的に処理し、一致ブロックを結果リストに追加
            PlotBestMatch(blocks, srcLines, srcAnalysis.source, srcStart, best.srcStart, destLines, destAnalysis.source, destStart, best.destStart)
            blocks.Add(New EvaluationBlock(EditTypeEnum.Match, best.srcStart, best.srcStart + best.length, best.destStart, best.destStart + best.length))
            PlotBestMatch(blocks, srcLines, srcAnalysis.source, best.srcStart + best.length, srcEnd, destLines, destAnalysis.source, best.destStart + best.length, destEnd)
        Else
            ' 一致ブロックが見つからない場合は、ソース範囲全体を削除、宛先範囲全体を挿入として処理
            blocks.Add(New EvaluationBlock(EditTypeEnum.Delete, srcStart, srcEnd, -1, -1))
            blocks.Add(New EvaluationBlock(EditTypeEnum.Insert, -1, -1, destStart, destEnd))
        End If
    End Sub

    ''' <summary>
    ''' 2つのテキスト範囲間で最適な一致ブロックを検出します。
    ''' ヒストグラム差分アルゴリズムに基づき、出現頻度が低い行を優先して一致を探索します。
    ''' </summary>
    ''' <param name="srcAnalysis">ソース側の文字列とその行番号配列の辞書</param>
    ''' <param name="srcLines">ソース側の行情報配列</param>
    ''' <param name="srcStart">ソース側の検索開始位置（0ベース）</param>
    ''' <param name="srcEnd">ソース側の検索終了位置（この位置は含まない）</param>
    ''' <param name="destAnalysis">宛先側の文字列とその行番号配列の辞書</param>
    ''' <param name="destLines">宛先側の行情報配列</param>
    ''' <param name="destStart">宛先側の検索開始位置（0ベース）</param>
    ''' <param name="destEnd">宛先側の検索終了位置（この位置は含まない）</param>
    ''' <returns>
    ''' 一致ブロックの情報を含むタプル：
    ''' - find: 一致ブロックが見つかったかどうか
    ''' - occurrence: 見つかったブロックの出現頻度（小さいほど固有）
    ''' - length: 一致ブロックの長さ（行数）
    ''' - srcStart: ソース側の一致ブロック開始位置
    ''' - destStart: 宛先側の一致ブロック開始位置
    ''' </returns>
    ''' <remarks>
    ''' このメソッドは以下の優先順位でベストマッチを選択します：
    ''' 1. 出現頻度が最も低い（固有性が高い）
    ''' 2. 同じ出現頻度の場合は、一致ブロックの長さが最も長い
    ''' 3. さらに同じ場合は、宛先側の開始位置が最も小さい
    ''' 
    ''' アルゴリズムは、各一致候補の前後に連続して一致する行を伸ばすことで、最大の一致ブロックを検出します。
    ''' </remarks>
    Private Function FindBestMatch(srcAnalysis As Dictionary(Of String, Integer()),
                                   srcLines As LineInfo(),
                                   srcStart As Integer,
                                   srcEnd As Integer,
                                   destAnalysis As Dictionary(Of String, Integer()),
                                   destLines As LineInfo(),
                                   destStart As Integer,
                                   destEnd As Integer) As (find As Boolean, occurrence As Long, length As Integer, srcStart As Integer, destStart As Integer)
        ' まず、両方の分析結果に存在する文字列を抽出し、出現頻度を計算してソートします
        Dim analysis As New List(Of GroupedHistogram)()
        For Each kvp In srcAnalysis
            If destAnalysis.ContainsKey(kvp.Key) Then
                analysis.Add(New GroupedHistogram(kvp.Key, kvp.Value, destAnalysis(kvp.Key)))
            End If
        Next
        analysis.Sort(Function(a, b) a.Occurrence.CompareTo(b.Occurrence))

        Dim bestFind As Boolean = False
        Dim bestOccurrence As Long = Long.MaxValue
        Dim bestLength As Integer = 0
        Dim bestSrcStart As Integer = -1
        Dim bestDestStart As Integer = -1

        For Each hist In analysis
            For Each s In hist.SourceLines
                For Each d In hist.DestinationLines
                    ' 前後へ伸ばして連続一致ブロックを測る
                    ' 一致ブロックの長さは、前方と後方の一致行数の合計で表される

                    ' 遡って一致する行を数える
                    Dim back As Integer = 0
                    While (s - back - 1) >= srcStart AndAlso (d - back - 1) >= destStart AndAlso
                          srcLines(s - back - 1).Str = destLines(d - back - 1).Str
                        back += 1
                    End While

                    ' 先へ一致する行を数える
                    Dim fwd As Integer = 0
                    While (s + fwd) < srcEnd AndAlso (d + fwd) < destEnd AndAlso
                          srcLines(s + fwd).Str = destLines(d + fwd).Str
                        fwd += 1
                    End While

                    ' 一致ブロックの開始位置と長さを計算
                    Dim srcBlockStart = s - back
                    Dim destBlockStart = d - back
                    Dim blockLength = back + fwd

                    ' ベスト更新判定：occ（小）→ length（大）→ 開始位置（小）
                    Dim occ = hist.Occurrence.CompareTo(bestOccurrence)
                    If occ < 0 OrElse
                       (occ = 0 AndAlso blockLength > bestLength) OrElse
                       (occ = 0 AndAlso blockLength = bestLength AndAlso destBlockStart < bestDestStart) Then
                        bestFind = True
                        bestOccurrence = hist.Occurrence
                        bestLength = blockLength
                        bestSrcStart = srcBlockStart
                        bestDestStart = destBlockStart
                    ElseIf occ > 0 Then
                        ' 出現頻度がこれ以上大きい候補はベストになり得ないので、次のヒストグラムへ
                        Return (bestFind, bestOccurrence, bestLength, bestSrcStart, bestDestStart)
                    End If
                Next
            Next
        Next

        Return (bestFind, bestOccurrence, bestLength, bestSrcStart, bestDestStart)
    End Function

#End Region

End Module
