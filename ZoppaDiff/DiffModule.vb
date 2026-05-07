Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaDiff.Collections

''' <summary>
''' Myers差分アルゴリズムとA*アルゴリズムを使用して文字列配列間の差分を検出するモジュール
''' </summary>
Public Module DiffModule

    ''' <summary>編集操作の種類を表す列挙型</summary>
    <Flags>
    Public Enum EditTypeEnum
        ''' <summary>一致（変更なし）</summary>
        Match = 0
        ''' <summary>差分（置換）</summary>
        Diff = 1
        ''' <summary>挿入</summary>
        Insert = 2
        ''' <summary>削除</summary>
        Delete = 4
    End Enum

    ''' <summary>空の行を表すLineInfoオブジェクト</summary>
    Private emptyLineInfo As New LineInfo(-1, String.Empty)

    ''' <summary>
    ''' Myers差分アルゴリズムを使用して2つの文字列配列間の差分を計算します
    ''' </summary>
    ''' <param name="src">比較元の文字列配列</param>
    ''' <param name="dest">比較先の文字列配列</param>
    ''' <returns>差分操作のリスト</returns>
    ''' <remarks>
    ''' このアルゴリズムは最初にMyers差分アルゴリズムを使用して大まかな差分を計算し、
    ''' その後A*アルゴリズムを使用してより詳細な差分を計算します
    ''' </remarks>
    Public Function Diff(src() As String, dest() As String) As List(Of Answer)
        If src Is Nothing Then Throw New ArgumentNullException(NameOf(src))
        If dest Is Nothing Then Throw New ArgumentNullException(NameOf(dest))

        Dim source = ConvertAll(src)
        Dim destination = ConvertAll(dest)

        Dim visits As New Dictionary(Of Integer, VisitPosition)(source.Length + destination.Length)
        Dim cur As New VisitPosition(Nothing, 0, 0)

        ' 最初の一致部分をスキップ
        Do While cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X).Str = destination(cur.Y).Str
            cur = New VisitPosition(cur, cur.X + 1, cur.Y + 1)
        Loop
        visits.Add(0, cur)

        ' Myers差分アルゴリズムのメインループ
        Dim answer As VisitPosition = Nothing
        For d As Integer = 1 To source.Length + destination.Length
            ' Myers差分アルゴリズムはK-lineを-dから+dまで2ずつ増やして探索
            For k = -d To d Step 2
                If k = -d OrElse (k <> d AndAlso visits(k - 1).X < visits(k + 1).X) Then
                    Dim prev = visits(k + 1)
                    cur = New VisitPosition(prev, prev.X, prev.X - k)
                Else
                    Dim prev = visits(k - 1)
                    cur = New VisitPosition(prev, prev.X + 1, (prev.X + 1) - k)
                End If

                ' 一致する部分を進める
                Do While cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X).Str = destination(cur.Y).Str
                    cur = New VisitPosition(cur, cur.X + 1, cur.Y + 1)
                Loop
                visits(k) = cur

                ' 終端に到達したかチェック
                If cur.X >= source.Length AndAlso cur.Y >= destination.Length Then
                    answer = cur
                    Exit For
                End If
            Next
            If answer IsNot Nothing Then
                Exit For
            End If
        Next

        ' 結果をトレースバックして最初のステップを構築
        Dim firstStepAnswer As New List(Of Answer)()
        Do While answer.From IsNot Nothing
            firstStepAnswer.Add(New Answer(source, answer.From.X, answer.X, destination, answer.From.Y, answer.Y, Nothing))
            answer = answer.From
        Loop

        ' A*アルゴリズムを使用して詳細な差分を計算
        Dim finalAnswer As New List(Of Answer)(firstStepAnswer.Count)
        Dim i As Integer = firstStepAnswer.Count - 1
        Do While i >= 0
            Select Case firstStepAnswer(i).EditType
                Case EditTypeEnum.Match
                    ' 連続する一致をグループ化
                    Do While i >= 0 AndAlso firstStepAnswer(i).EditType = EditTypeEnum.Match
                        finalAnswer.Add(firstStepAnswer(i))
                        i -= 1
                    Loop

                Case EditTypeEnum.Insert, EditTypeEnum.Delete
                    ' 挿入と削除のグループに対してA*差分を適用
                    Dim tmpSrc As New List(Of LineInfo)()
                    Dim tmpDest As New List(Of LineInfo)()
                    Do While i >= 0
                        Select Case firstStepAnswer(i).EditType
                            Case EditTypeEnum.Match
                                Exit Do
                            Case EditTypeEnum.Insert
                                tmpDest.Add(firstStepAnswer(i).Destination)
                            Case EditTypeEnum.Delete
                                tmpSrc.Add(firstStepAnswer(i).Source)
                        End Select
                        i -= 1
                    Loop
                    finalAnswer.AddRange(AStarDiff(tmpSrc.ToArray(), tmpDest.ToArray()))

                Case Else
                    Throw New InvalidOperationException($"不正な編集タイプ: {firstStepAnswer(i).EditType}")
            End Select
        Loop

        Return finalAnswer
    End Function

    ''' <summary>
    ''' 文字列配列をLineInfo配列に変換します
    ''' </summary>
    ''' <param name="arr">変換対象の文字列配列</param>
    ''' <returns>行番号付きのLineInfo配列</returns>
    Private Function ConvertAll(arr() As String) As LineInfo()
        Dim result(arr.Length - 1) As LineInfo
        For i As Integer = 0 To arr.Length - 1
            result(i) = New LineInfo(i, arr(i))
        Next
        Return result
    End Function

    ''' <summary>
    ''' A*アルゴリズムを使用して文字列配列間の最適な差分を計算します
    ''' </summary>
    ''' <param name="source">比較元の文字列配列</param>
    ''' <param name="destination">比較先の文字列配列</param>
    ''' <returns>最適化された差分操作のリスト</returns>
    ''' <remarks>
    ''' このメソッドは文字列の長さを考慮したコストベースの最適化を行います
    ''' </remarks>
    Private Function AStarDiff(source() As LineInfo, destination() As LineInfo) As List(Of Answer)
        ' ベースケース: 一方が空の場合
        If source.Length = 0 Then
            Return CreateEmptyAnswer(destination, Function(s, i) New Answer(EditTypeEnum.Insert, emptyLineInfo, s))
        End If
        If destination.Length = 0 Then
            Return CreateEmptyAnswer(source, Function(s, i) New Answer(EditTypeEnum.Delete, s, emptyLineInfo))
        End If

        ' 文字列長の配列を作成（ヒューリスティック計算用）
        Dim srclen = CalculateLengthCost(source)
        Dim destlen = CalculateLengthCost(destination)

        ' A*アルゴリズムの初期化
        Dim startPos As New CostPosition(Nothing, 0, 0, 0, 0, Nothing)

        Dim open As New SortedSet(Of CostPosition)()
        Dim closed As New SortedSet(Of CostPosition)()
        open.Add(New CostPosition(Nothing, 0, 0, 0, 0, Nothing))

        Dim order As New BPlusTree(Of CostPosition)(AddressOf CostPositionComparer)
        order.Insert(startPos)

        ' A*探索のメインループ
        Dim answer As CostPosition = Nothing
        Do While order.Count > 0
            Dim cur = order(0)
            order.Remove(cur)

            ' オープンリストから現在の位置を取得してクローズドリストに移動
            open.Remove(cur)
            closed.Add(cur)

            ' ゴール状態に到達したかチェック
            If cur.X >= source.Length AndAlso cur.Y >= destination.Length Then
                answer = cur
                Exit Do
            End If

            ' 隣接する状態を生成
            If cur.X < source.Length AndAlso cur.Y < destination.Length Then
                ' 差異、挿入、削除の各操作を評価して次の位置に移動
                Dim editChars As New List(Of EditChar)(source(cur.X).Str.Length + destination(cur.Y).Str.Length)
                Dim editCost As Integer = SplitAStarCharDiff(source(cur.X).Str, destination(cur.Y).Str, editChars)
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 1, editCost, srclen, destlen, editChars.ToArray()))
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, source(cur.X).Str.Length, srclen, destlen, Nothing))
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, destination(cur.Y).Str.Length, srclen, destlen, Nothing))
            Else
                ' 挿入、削除の各操作を評価して次の位置に移動
                If cur.X < source.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, source(cur.X).Str.Length, srclen, destlen, Nothing))
                End If
                If cur.Y < destination.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, destination(cur.Y).Str.Length, srclen, destlen, Nothing))
                End If
            End If
        Loop

        ' 結果をトレースバックして構築
        Dim ans As New List(Of Answer)()
        Dim traceStack As New Stack(Of CostPosition)()
        Do While answer IsNot Nothing AndAlso answer.From IsNot Nothing
            traceStack.Push(answer)
            answer = answer.From
        Loop
        Do While traceStack.Count > 0
            Dim pos = traceStack.Pop()
            ans.Add(New Answer(source, pos.From.X, pos.X, destination, pos.From.Y, pos.Y, pos.EditChars))
        Loop
        Return ans
    End Function

    ''' <summary>
    ''' 文字列配列の各要素に対して指定された作成関数を適用し、Answerリストを生成します
    ''' </summary>
    ''' <param name="target">処理対象の文字列配列</param>
    ''' <param name="createMethod">文字列からAnswerを作成する関数</param>
    ''' <returns>Answerオブジェクトのリスト</returns>
    Private Function CreateEmptyAnswer(target() As LineInfo, createMethod As Func(Of LineInfo, Integer, Answer)) As List(Of Answer)
        Dim ans As New List(Of Answer)()
        For i As Integer = 0 To target.Length - 1
            ans.Add(createMethod(target(i), i))
        Next
        Return ans
    End Function

    ''' <summary>
    ''' 文字列配列の各位置から末尾までの累積文字長を計算します
    ''' </summary>
    ''' <param name="arr">文字列配列</param>
    ''' <returns>各位置から末尾までの文字長の合計を格納した配列</returns>
    ''' <remarks>A*アルゴリズムのヒューリスティック計算で使用されます</remarks>
    Private Function CalculateLengthCost(arr() As LineInfo) As Integer()
        Dim result(arr.Length) As Integer
        For i As Integer = arr.Length - 1 To 0 Step -1
            result(i) = result(i + 1) + arr(i).Str.Length
        Next
        Return result
    End Function

    ''' <summary>2つの文字列間の編集距離をA*アルゴリズムで計算します</summary>
    ''' <param name="source">比較元の文字列</param>
    ''' <param name="destination">比較先の文字列</param>
    ''' <param name="editChars">編集操作の詳細を格納するリスト（置換操作の内容を文字として格納）</param>
    ''' <returns>編集距離（コスト）</returns>
    Private Function SplitAStarCharDiff(source As String, destination As String, editChars As List(Of EditChar)) As Integer
        ' 文字列の先頭の空白をスキップして、空白部分と非空白部分を分割
        Dim ssplit = IndexOfNotSpace(source)
        Dim dsplit = IndexOfNotSpace(destination)

        ' 先頭の空白部分のコストを計算
        Dim cost = AStarCharDiff(source.Substring(0, ssplit), destination.Substring(0, dsplit), editChars)

        ' 非空白部分のコストを計算
        cost += AStarCharDiff(source.Substring(ssplit), destination.Substring(dsplit), editChars)

        Return cost
    End Function

    ''' <summary>文字列の先頭から最初の非空白文字の位置を返します</summary>
    ''' <param name="str">入力文字列</param>
    ''' <returns>最初の非空白文字のインデックス。全て空白の場合は文字列の長さを返します。</returns>
    Private Function IndexOfNotSpace(str As String) As Integer
        For i As Integer = 0 To str.Length - 1
            If Not Char.IsWhiteSpace(str(i)) Then
                Return i
            End If
        Next
        Return str.Length
    End Function

    ''' <summary>
    ''' A*アルゴリズムを使用して2つの文字列間の編集距離を計算します
    ''' </summary>
    ''' <param name="source">比較元の文字列</param>
    ''' <param name="destination">比較先の文字列</param>
    ''' <param name="editChars">編集操作の詳細を格納するリスト（置換操作の内容を文字として格納）</param>
    ''' <returns>編集距離（コスト）</returns>
    ''' <remarks>
    ''' このメソッドは文字レベルでの編集距離を計算し、A*差分アルゴリズムで使用されます
    ''' </remarks>
    Public Function AStarCharDiff(source As String, destination As String, editChars As List(Of EditChar)) As Integer
        ' 開始位置の初期化
        Dim startPos As New CostPosition(Nothing, 0, 0, 0, 0, Nothing)

        ' A*アルゴリズムの初期化
        Dim open As New SortedSet(Of CostPosition)()
        Dim closed As New SortedSet(Of CostPosition)()
        Dim order As New BPlusTree(Of CostPosition)(AddressOf CostPositionComparer)

        ' 開始位置をオープンリストと評価リストに追加
        open.Add(startPos)
        order.Insert(startPos)

        ' A*探索のメインループ
        Dim answer As CostPosition = Nothing
        Do While order.Count > 0
            ' 最小コストの位置を取得
            Dim cur = order(0)
            order.Remove(cur)

            ' オープンリストから現在の位置を取得してクローズドリストに移動
            open.Remove(cur)
            closed.Add(cur)

            ' ゴール状態に到達したかチェック
            If cur.X >= source.Length AndAlso cur.Y >= destination.Length Then
                answer = cur
                Exit Do
            End If

            If cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X) = destination(cur.Y) Then
                ' 一致する場合はコスト0で次の位置に移動
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 1, 0, source.Length, destination.Length))
            Else
                ' 挿入、削除の各操作を評価して次の位置に移動
                If cur.Y < destination.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, 1, source.Length, destination.Length))
                End If
                If cur.X < source.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, 1, source.Length, destination.Length))
                End If
            End If
        Loop

        ' 結果をトレースバックして編集操作の詳細を構築
        Dim ans As Integer = 0
        Dim trace As New Stack(Of EditChar)()
        Do While answer IsNot Nothing AndAlso answer.From IsNot Nothing
            If answer.From.X < answer.X AndAlso answer.From.Y = answer.Y Then
                trace.Push(New EditChar(EditTypeEnum.Delete, source(answer.From.X))) ' 削除
                ans += 1
            ElseIf answer.From.X = answer.X AndAlso answer.From.Y < answer.Y Then
                trace.Push(New EditChar(EditTypeEnum.Insert, destination(answer.From.Y))) ' 挿入
                ans += 1
            Else
                trace.Push(New EditChar(EditTypeEnum.Match, source(answer.From.X))) ' 一致
            End If
            answer = answer.From
        Loop

        ' トレースバックした編集操作をリストに追加
        Do While trace.Count > 0
            editChars.Add(trace.Pop())
        Loop
        Return ans
    End Function

    ''' <summary>
    ''' A*アルゴリズムで使用するCostPositionオブジェクトの比較関数
    ''' </summary>
    ''' <param name="a">比較対象の最初のCostPosition</param>
    ''' <param name="b">比較対象の二番目のCostPosition</param>
    ''' <returns>
    ''' 比較結果を表す整数値：
    ''' <list type="bullet">
    ''' <item>負の値: aがbより小さい（優先度が高い）</item>
    ''' <item>0: aとbが等しい</item>
    ''' <item>正の値: aがbより大きい（優先度が低い）</item>
    ''' </list>
    ''' </returns>
    ''' <remarks>
    ''' この比較関数は以下の優先順位で比較を行います：
    ''' <list type="number">
    ''' <item><see cref="CostPosition.TotalCost"/>: 総コスト（到達コスト + ヒューリスティックコスト）が最優先</item>
    ''' <item><see cref="CostPosition.Y"/>: Y座標（デスティネーション配列のインデックス）が第二優先</item>
    ''' <item><see cref="CostPosition.X"/>: X座標（ソース配列のインデックス）が第三優先</item>
    ''' </list>
    ''' 
    ''' この比較順序により、A*アルゴリズムのオープンリストで最も有望な（総コストが低い）位置が
    ''' 優先的に選択され、同じコストの場合は座標による一貫した順序付けが保証されます。
    ''' 
    ''' <para>
    ''' <strong>使用箇所:</strong><br/>
    ''' - <see cref="BPlusTree(Of CostPosition)"/>での挿入・削除時の順序付け<br/>
    ''' - A*探索における次に展開すべきノードの決定
    ''' </para>
    ''' </remarks>
    Private Function CostPositionComparer(a As CostPosition, b As CostPosition) As Integer
        ' まず総コストで比較
        Dim r = a.TotalCost.CompareTo(b.TotalCost)
        If r <> 0 Then
            Return r
        End If

        ' 次にY座標で比較
        r = a.Y.CompareTo(b.Y)
        If r <> 0 Then
            Return r
        End If

        ' 最後にX座標で比較
        Return a.X.CompareTo(b.X)
    End Function

    ''' <summary>
    ''' A*アルゴリズムでオープンリスト内の位置を更新または追加します
    ''' </summary>
    ''' <param name="open">探索対象のオープンリスト（SortedSet）</param>
    ''' <param name="closed">探索済みのクローズドリスト（SortedSet）</param>
    ''' <param name="order">効率的な最小コスト取得のためのB+ツリー</param>
    ''' <param name="newCur">新しく評価する位置とコスト情報</param>
    ''' <remarks>
    ''' このメソッドはA*アルゴリズムの核となる処理で、以下の動作を行います：
    ''' <list type="number">
    ''' <item>クローズドリストに既に存在する位置は処理をスキップ</item>
    ''' <item>オープンリストに同じ位置が存在する場合、より低いコストの経路のみ採用</item>
    ''' <item>新しい位置の場合はオープンリストとB+ツリーに追加</item>
    ''' <item>更新時はオープンリストとB+ツリーの両方を同期して更新</item>
    ''' </list>
    ''' B+ツリーは最小コストの位置を効率的に取得するために使用されます。
    ''' </remarks>
    Private Sub UpdatePosition(open As SortedSet(Of CostPosition),
                               closed As SortedSet(Of CostPosition),
                               order As BPlusTree(Of CostPosition),
                               newCur As CostPosition)
        ' クローズドリストに既に存在する場合は処理をスキップ
        If closed.Contains(newCur) Then
            Return
        End If

        ' オープンリストに同じ位置が存在するかチェック
        Dim oldCur As CostPosition = Nothing
        If open.TryGetValue(newCur, oldCur) Then
            ' 既存の位置より低いコストの場合のみ更新
            If newCur.TotalCost < oldCur.TotalCost Then
                ' 既存位置を削除
                open.Remove(oldCur)
                order.Remove(oldCur)

                ' 新しい位置を追加
                open.Add(newCur)
                order.Insert(newCur)
            End If
        Else
            ' 新しい位置の場合は追加
            open.Add(newCur)
            order.Insert(newCur)
        End If
    End Sub

    ''' <summary>
    ''' 新しいコスト位置を作成します（配列長ベース）
    ''' </summary>
    ''' <param name="current">現在の位置</param>
    ''' <param name="moveX">X軸の移動量</param>
    ''' <param name="moveY">Y軸の移動量</param>
    ''' <param name="arrivalCost">到達コスト</param>
    ''' <param name="srclen">ソース配列の長さ配列</param>
    ''' <param name="destlen">デスティネーション配列の長さ配列</param>
    ''' <param name="editChars">編集操作の詳細</param>
    ''' <returns>新しいコスト位置</returns>
    Private Function CreateNewPosition(current As CostPosition,
                                       moveX As Integer,
                                       moveY As Integer,
                                       arrivalCost As Integer,
                                       srclen As Integer(),
                                       destlen As Integer(),
                                       editChars() As EditChar) As CostPosition
        Dim x = current.X + moveX
        Dim y = current.Y + moveY
        Return New CostPosition(current,
                                current.X + moveX,
                                current.Y + moveY,
                                current.ArrivalCost + arrivalCost,
                                Math.Abs(srclen(x) - destlen(y)),
                                editChars)
    End Function

    ''' <summary>
    ''' 新しいコスト位置を作成します（残り長さ指定）
    ''' </summary>
    ''' <param name="current">現在の位置</param>
    ''' <param name="moveX">X軸の移動量</param>
    ''' <param name="moveY">Y軸の移動量</param>
    ''' <param name="arrivalCost">到達コスト</param>
    ''' <param name="srcLen">ソースの文字列長</param>
    ''' <param name="destLen">デスティネーションの文字列長</param>
    ''' <returns>新しいコスト位置</returns>
    Private Function CreateNewPosition(current As CostPosition,
                                       moveX As Integer,
                                       moveY As Integer,
                                       arrivalCost As Integer,
                                       srcLen As Integer,
                                       destLen As Integer) As CostPosition
        Return New CostPosition(current,
                                current.X + moveX,
                                current.Y + moveY,
                                current.ArrivalCost + arrivalCost,
                                Math.Abs((srcLen - current.X - moveX) - (destLen - current.Y - moveY)),
                                Nothing)
    End Function

    ''' <summary>
    ''' Myers差分アルゴリズムで使用される訪問位置を表すクラス
    ''' </summary>
    ''' <remarks>
    ''' Kライン（対角線）上での位置を追跡し、差分計算の経路を記録します
    ''' </remarks>
    Private NotInheritable Class VisitPosition
        Implements IComparable(Of VisitPosition)

        ''' <summary>前の位置への参照</summary>
        Public ReadOnly Property [From] As VisitPosition

        ''' <summary>X軸の位置（ソース配列のインデックス）</summary>
        Public ReadOnly Property X As Integer

        ''' <summary>Y軸の位置（デスティネーション配列のインデックス）</summary>
        Public ReadOnly Property Y As Integer

        ''' <summary>Kライン（X - Y の値）</summary>
        Public ReadOnly Property K As Integer
            Get
                Return Me.X - Me.Y
            End Get
        End Property

        ''' <summary>
        ''' 新しい訪問位置を初期化します
        ''' </summary>
        ''' <param name="from">前の位置</param>
        ''' <param name="x">X軸の位置</param>
        ''' <param name="y">Y軸の位置</param>
        Public Sub New([from] As VisitPosition, x As Integer, y As Integer)
            Me.From = [from]
            Me.X = x
            Me.Y = y
        End Sub

        ''' <summary>
        ''' 他の訪問位置との比較を行います
        ''' </summary>
        ''' <param name="other">比較対象</param>
        ''' <returns>比較結果</returns>
        Public Function CompareTo(other As VisitPosition) As Integer Implements IComparable(Of VisitPosition).CompareTo
            Dim r = Me.K - other.K
            If r <> 0 Then
                Return r
            End If
            Return other.X - Me.X
        End Function

        ''' <summary>
        ''' 訪問位置の文字列表現を返します
        ''' </summary>
        ''' <returns>位置情報の文字列</returns>
        Public Overrides Function ToString() As String
            If Me.From IsNot Nothing Then
                Return $"({Me.From.X}, {Me.From.Y}) -> ({Me.X}, {Me.Y}) k={Me.K}"
            Else
                Return $"({Me.X}, {Me.Y}) k={Me.K}"
            End If
        End Function

    End Class

    ''' <summary>
    ''' A*アルゴリズムで使用されるコスト付き位置を表すクラス
    ''' </summary>
    ''' <remarks>
    ''' 到達コストとヒューリスティックコストを含む完全なA*ノードです
    ''' </remarks>
    Private NotInheritable Class CostPosition
        Implements IComparable(Of CostPosition)

        ''' <summary>前の位置への参照</summary>
        Public ReadOnly Property [From] As CostPosition

        ''' <summary>X軸の位置</summary>
        Public ReadOnly Property X As Integer

        ''' <summary>Y軸の位置</summary>
        Public ReadOnly Property Y As Integer

        ''' <summary>この位置への到達コスト</summary>
        Public ReadOnly Property ArrivalCost As Integer

        ''' <summary>ゴールまでの推定コスト（ヒューリスティック）</summary>
        Public ReadOnly Property HeuristicCost As Integer

        ''' <summary>編集操作の詳細（主にA*アルゴリズムでの置換操作の内容を格納）</summary>
        Public ReadOnly Property EditChars As EditChar()

        ''' <summary>総コスト（到達コスト + ヒューリスティックコスト）</summary>
        Public ReadOnly Property TotalCost As Integer
            Get
                Return Me.ArrivalCost + Me.HeuristicCost
            End Get
        End Property

        ''' <summary>
        ''' 新しいコスト位置を初期化します
        ''' </summary>
        ''' <param name="from">前の位置</param>
        ''' <param name="x">X軸の位置</param>
        ''' <param name="y">Y軸の位置</param>
        ''' <param name="arrival">到達コスト</param>
        ''' <param name="heuristic">ヒューリスティックコスト</param>
        ''' <param name="editChars">編集操作の詳細を表す配列</param>
        Public Sub New([from] As CostPosition, x As Integer, y As Integer, arrival As Integer, heuristic As Integer, editChars As EditChar())
            Me.From = [from]
            Me.X = x
            Me.Y = y
            Me.ArrivalCost = arrival
            Me.HeuristicCost = heuristic
            Me.EditChars = editChars
        End Sub

        ''' <summary>
        ''' 他のコスト位置との比較を行います
        ''' </summary>
        ''' <param name="other">比較対象</param>
        ''' <returns>比較結果</returns>
        Public Function CompareTo(other As CostPosition) As Integer Implements IComparable(Of CostPosition).CompareTo
            Dim r = Me.Y - other.Y
            If r <> 0 Then
                Return r
            End If
            Return Me.X - other.X
        End Function

        ''' <summary>
        ''' コスト位置の文字列表現を返します
        ''' </summary>
        ''' <returns>位置とコスト情報の文字列</returns>
        Public Overrides Function ToString() As String
            If Me.From IsNot Nothing Then
                Return $"({Me.From.X}, {Me.From.Y}) -> ({Me.X}, {Me.Y}) TotalCost={Me.TotalCost}"
            Else
                Return $"({Me.X}, {Me.Y}) TotalCost={Me.TotalCost}"
            End If
        End Function

    End Class

    ''' <summary>
    ''' 文字レベルでの編集操作を表す構造体
    ''' </summary>
    ''' <remarks>
    ''' A*アルゴリズムによる文字単位の差分計算で使用され、
    ''' 個別の文字に対する編集操作（一致、挿入、削除）の詳細を格納します。
    ''' </remarks>
    Public Structure EditChar

        ''' <summary>
        ''' 編集操作の種類を取得します
        ''' </summary>
        ''' <value>
        ''' <see cref="EditTypeEnum"/>の値：
        ''' <list type="bullet">
        ''' <item><see cref="EditTypeEnum.Match"/>: 文字が一致</item>
        ''' <item><see cref="EditTypeEnum.Insert"/>: 文字の挿入</item>
        ''' <item><see cref="EditTypeEnum.Delete"/>: 文字の削除</item>
        ''' </list>
        ''' </value>
        Public ReadOnly Property EditType As EditTypeEnum

        ''' <summary>
        ''' 編集操作の対象となる文字を取得します
        ''' </summary>
        ''' <value>
        ''' 編集操作に関連する文字。
        ''' 挿入操作の場合は挿入される文字、削除操作の場合は削除される文字、
        ''' 一致操作の場合は一致する文字を表します。
        ''' </value>
        Public ReadOnly Property EditChar As Char

        ''' <summary>
        ''' 編集操作の種類と対象文字を指定してEditChar構造体を初期化します
        ''' </summary>
        ''' <param name="EditType">編集操作の種類</param>
        ''' <param name="EditChar">編集操作の対象となる文字</param>
        Public Sub New(EditType As EditTypeEnum, EditChar As Char)
            Me.EditType = EditType
            Me.EditChar = EditChar
        End Sub

        ''' <summary>
        ''' 編集操作の種類と対象文字を組み合わせた文字列表現を返します
        ''' </summary>
        ''' <returns>編集操作の文字列</returns>
        Public Overrides Function ToString() As String
            Return $"{Me.EditType}: {Me.EditChar}"
        End Function
    End Structure

    ''' <summary>
    ''' 差分操作の結果を表すクラス
    ''' </summary>
    ''' <remarks>
    ''' 編集の種類と対象となる文字列を含む差分操作の詳細情報です
    ''' </remarks>
    Public NotInheritable Class Answer

        ''' <summary>編集操作の種類</summary>
        Public ReadOnly Property EditType As EditTypeEnum

        ''' <summary>ソース文字列</summary>
        Public ReadOnly Property Source As LineInfo

        ''' <summary>デスティネーション文字列</summary>
        Public ReadOnly Property Destination As LineInfo

        ''' <summary>編集操作の詳細（主にA*アルゴリズムでの置換操作の内容を格納）</summary>
        Private ReadOnly editChars As EditChar()

        ''' <summary>
        ''' 編集操作の種類と対象文字列を組み合わせた差分表示用の文字列を取得します
        ''' </summary>
        ''' <returns>分表示用の文字列</returns>
        Public ReadOnly Property EditString As String
            Get
                ' 編集操作の詳細がない場合は空文字列を返す
                If Me.editChars Is Nothing OrElse Me.editChars.Length = 0 Then
                    Return String.Empty
                End If

                ' 文字列ビルダーを使用して編集操作の詳細を構築
                Dim res As New StringBuilder()
                Dim currentMode = EditTypeEnum.Match

                ' 編集操作の詳細を順に処理して、モードの切り替えや文字の追加を行う
                For Each edit In Me.editChars
                    If edit.EditType <> currentMode Then
                        CloseCurrentMode(res, currentMode)
                        OpenNewMode(res, edit.EditType, edit.EditChar)
                        currentMode = edit.EditType
                    Else
                        AppendCharacterForCurrentMode(res, edit.EditChar)
                    End If
                Next

                CloseCurrentMode(res, currentMode)
                Return res.ToString()
            End Get
        End Property

        ''' <summary>
        ''' 編集タイプと文字列を指定してAnswerを初期化します
        ''' </summary>
        ''' <param name="EditType">編集操作の種類</param>
        ''' <param name="src">ソース文字列</param>
        ''' <param name="dest">デスティネーション文字列</param>
        Public Sub New(EditType As EditTypeEnum, src As LineInfo, dest As LineInfo)
            Me.EditType = EditType
            Me.Source = src
            Me.Destination = dest
        End Sub

        ''' <summary>
        ''' 配列と位置情報からAnswerを初期化します
        ''' </summary>
        ''' <param name="source">ソース配列</param>
        ''' <param name="prevX">前のX位置</param>
        ''' <param name="nextX">次のX位置</param>
        ''' <param name="destination">デスティネーション配列</param>
        ''' <param name="prevY">前のY位置</param>
        ''' <param name="nextY">次のY位置</param>
        ''' <param name="editChars">編集操作の詳細</param>
        Public Sub New(source() As LineInfo,
                       prevX As Integer,
                       nextX As Integer,
                       destination() As LineInfo,
                       prevY As Integer,
                       nextY As Integer,
                       editChars As EditChar())
            ' 位置の変化から編集タイプを決定
            If nextX = prevX Then
                Me.EditType = EditTypeEnum.Insert
            ElseIf nextY = prevY Then
                Me.EditType = EditTypeEnum.Delete
            ElseIf editChars Is Nothing Then
                Me.EditType = EditTypeEnum.Match
            Else
                Me.EditType = EditTypeEnum.Diff
            End If

            ' 編集タイプに応じて文字列を設定
            Select Case Me.EditType
                Case EditTypeEnum.Match, EditTypeEnum.Diff
                    Me.Source = source(prevX)
                    Me.Destination = destination(prevY)
                Case EditTypeEnum.Insert
                    Me.Source = emptyLineInfo
                    Me.Destination = destination(prevY)
                Case EditTypeEnum.Delete
                    Me.Source = source(prevX)
                    Me.Destination = emptyLineInfo
            End Select

            ' 編集情報の並び替え
            If editChars IsNot Nothing Then
                Me.editChars = SortEditChars(editChars)
            End If
        End Sub

        ''' <summary>
        ''' 編集文字配列を編集種類別に並び替えます
        ''' </summary>
        ''' <param name="editChars">並び替え対象の編集文字配列</param>
        ''' <returns>並び替え済みの編集文字配列</returns>
        ''' <remarks>
        ''' 編集操作を種類別にグループ化し、削除操作を挿入操作よりも前に配置します。
        ''' 一致操作はそのままの位置を保持します。
        ''' </remarks>
        Private Shared Function SortEditChars(editChars As EditChar()) As EditChar()
            If editChars Is Nothing OrElse editChars.Length <= 1 Then
                Return editChars
            End If

            Dim result(editChars.Length - 1) As EditChar
            Array.Copy(editChars, result, editChars.Length)

            ' 一致操作以外の部分を種類別に安定ソート
            For i As Integer = 0 To result.Length - 1
                If result(i).EditType <> EditTypeEnum.Match Then
                    Dim groupStart = i

                    ' 連続する非一致操作の終了位置を見つける
                    Dim groupEnd = i
                    Do While groupEnd < result.Length AndAlso result(groupEnd).EditType <> EditTypeEnum.Match
                        groupEnd += 1
                    Loop
                    groupEnd -= 1

                    ' 該当範囲をソート（削除操作が挿入操作より前になるように）
                    If groupEnd > groupStart Then
                        SortEditGroup(result, groupStart, groupEnd)
                    End If
                    i = groupEnd
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' 指定された範囲の編集文字を挿入ソートで並び替えます
        ''' </summary>
        ''' <param name="editChars">編集文字配列</param>
        ''' <param name="startIndex">並び替え開始インデックス</param>
        ''' <param name="endIndex">並び替え終了インデックス</param>
        ''' <remarks>
        ''' 削除操作（Delete）を挿入操作（Insert）より前に配置する安定ソートを実行します
        ''' </remarks>
        Private Shared Sub SortEditGroup(editChars As EditChar(), startIndex As Integer, endIndex As Integer)
            For i As Integer = startIndex + 1 To endIndex
                Dim current = editChars(i)
                Dim j = i - 1

                ' 現在の要素を適切な位置に挿入（削除操作が挿入操作より前になるように）
                Do While j >= startIndex AndAlso editChars(j).EditType < current.EditType
                    editChars(j + 1) = editChars(j)
                    j -= 1
                Loop
                editChars(j + 1) = current
            Next
        End Sub

        ''' <summary>
        ''' 現在の編集モードを閉じる処理を行います
        ''' </summary>
        ''' <param name="sb">文字列ビルダー</param>
        ''' <param name="mode">現在のモード</param>
        Private Shared Sub CloseCurrentMode(sb As StringBuilder, mode As EditTypeEnum)
            If mode <> EditTypeEnum.Match Then
                sb.Append("}")
            End If
        End Sub

        ''' <summary>
        ''' 新しい編集モードを開始する処理を行います
        ''' </summary>
        ''' <param name="sb">文字列ビルダー</param>
        ''' <param name="newMode">新しいモード</param>
        ''' <param name="editChar">編集文字</param>
        Private Shared Sub OpenNewMode(sb As StringBuilder, newMode As EditTypeEnum, editChar As Char)
            Select Case newMode
                Case EditTypeEnum.Match
                    sb.Append(EscapeString(editChar))
                Case EditTypeEnum.Delete
                    sb.Append("{D:").Append(EscapeString(editChar))
                Case EditTypeEnum.Insert
                    sb.Append("{I:").Append(EscapeString(editChar))
            End Select
        End Sub

        ''' <summary>
        ''' 現在のモードに応じて文字を追加します
        ''' </summary>
        ''' <param name="sb">文字列ビルダー</param>
        ''' <param name="editChar">編集文字</param>
        Private Shared Sub AppendCharacterForCurrentMode(sb As StringBuilder, editChar As Char)
            sb.Append(EscapeString(editChar))
        End Sub

        ''' <summary>
        ''' 特殊文字をエスケープして安全な文字列表現に変換します
        ''' </summary>
        ''' <param name="s">エスケープ処理を行う文字</param>
        ''' <returns>
        ''' エスケープ処理された文字列。以下の変換が行われます：
        ''' <list type="bullet">
        ''' <item>バックスラッシュ（\）→ \\（二重バックスラッシュ）</item>
        ''' <item>左波括弧（{）→ \{（バックスラッシュエスケープ）</item>
        ''' <item>右波括弧（}）→ \}（バックスラッシュエスケープ）</item>
        ''' <item>その他の文字 → そのまま文字列として返す</item>
        ''' </list>
        ''' </returns>
        ''' <remarks>
        ''' この関数は<see cref="Answer.EditString"/>プロパティ内で使用され、
        ''' 差分表示における特殊文字の適切なエスケープ処理を提供します。
        ''' 波括弧（{ }）は差分表示フォーマットにおいて編集操作の境界を示すために
        ''' 使用されるため、これらの文字がリテラル文字として表示される際には
        ''' エスケープが必要になります。
        ''' 
        ''' <para>
        ''' <strong>使用例：</strong><br/>
        ''' バックスラッシュ文字 → "\\"<br/>
        ''' 左波括弧文字 → "\{"<br/>
        ''' 右波括弧文字 → "\}"<br/>
        ''' 通常の文字 → そのまま
        ''' </para>
        ''' </remarks>
        ''' <seealso cref="Answer.EditString"/>
        Private Shared Function EscapeString(s As Char) As String
            Select Case s
                Case "\"c
                    Return "\\"
                Case "{"c
                    Return "\{"
                Case "}"c
                    Return "\}"
                Case Else
                    Return s.ToString()
            End Select
        End Function

    End Class

    ''' <summary>
    ''' 行番号とその内容を表す構造体
    ''' </summary>
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

End Module
