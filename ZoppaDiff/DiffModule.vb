Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaDiff.Collections

''' <summary>
''' Myers差分アルゴリズムとA*アルゴリズムを使用して文字列配列間の差分を検出するモジュール
''' </summary>
Module DiffModule

    ''' <summary>
    ''' 編集操作の種類を表す列挙型
    ''' </summary>
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

    ''' <summary>
    ''' Myers差分アルゴリズムを使用して2つの文字列配列間の差分を計算します
    ''' </summary>
    ''' <param name="source">比較元の文字列配列</param>
    ''' <param name="destination">比較先の文字列配列</param>
    ''' <returns>差分操作のリスト</returns>
    ''' <remarks>
    ''' このアルゴリズムは最初にMyers差分アルゴリズムを使用して大まかな差分を計算し、
    ''' その後A*アルゴリズムを使用してより詳細な差分を計算します
    ''' </remarks>
    Public Function Diff(source() As String, destination() As String) As List(Of Answer)
        If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))
        If destination Is Nothing Then Throw New ArgumentNullException(NameOf(destination))

        Dim visits As New Dictionary(Of Integer, VisitPosition)
        Dim cur As New VisitPosition(Nothing, 0, 0)

        ' 最初の一致部分をスキップ
        Do While cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X) = destination(cur.Y)
            cur = New VisitPosition(cur, cur.X + 1, cur.Y + 1)
        Loop
        visits.Add(0, cur)

        ' Myers差分アルゴリズムのメインループ
        Dim answer As VisitPosition = Nothing
        For d As Integer = 1 To source.Length + destination.Length
            For k = -d To d Step 2
                If k = -d OrElse (k <> d AndAlso visits(k - 1).X < visits(k + 1).X) Then
                    Dim prev = visits(k + 1)
                    cur = New VisitPosition(prev, prev.X, prev.X - k)
                Else
                    Dim prev = visits(k - 1)
                    cur = New VisitPosition(prev, prev.X + 1, (prev.X + 1) - k)
                End If

                ' 一致する部分を進める
                Do While cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X) = destination(cur.Y)
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
        Dim j As Integer
        For i As Integer = firstStepAnswer.Count - 1 To 0 Step -1
            j = i
            Select Case firstStepAnswer(i).EditType
                Case EditTypeEnum.Match
                    ' 連続する一致をグループ化
                    Do While j >= 0 AndAlso firstStepAnswer(j).EditType = EditTypeEnum.Match
                        finalAnswer.Add(firstStepAnswer(j))
                        j -= 1
                    Loop
                    i = j + 1

                Case EditTypeEnum.Insert, EditTypeEnum.Delete
                    ' 挿入と削除のグループに対してA*差分を適用
                    Dim tmpSrc As New List(Of String)()
                    Dim tmpDest As New List(Of String)()
                    Do While j >= 0
                        Select Case firstStepAnswer(j).EditType
                            Case EditTypeEnum.Match
                                Exit Do
                            Case EditTypeEnum.Insert
                                tmpDest.Add(firstStepAnswer(j).Destination)
                            Case EditTypeEnum.Delete
                                tmpSrc.Add(firstStepAnswer(j).Source)
                        End Select
                        j -= 1
                    Loop
                    finalAnswer.AddRange(ASterDiff(tmpSrc.ToArray(), tmpDest.ToArray()))
                    i = j + 1
            End Select
        Next

        Return finalAnswer
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
    Private Function ASterDiff(source() As String, destination() As String) As List(Of Answer)
        ' ベースケース: 一方が空の場合
        If source.Length = 0 Then
            Dim lans As New List(Of Answer)(destination.Length)
            For Each s As String In destination
                lans.Add(New Answer(EditTypeEnum.Insert, "", s))
            Next
            Return lans
        End If
        If destination.Length = 0 Then
            Dim lans As New List(Of Answer)(source.Length)
            For Each s As String In source
                lans.Add(New Answer(EditTypeEnum.Delete, s, ""))
            Next
            Return lans
        End If

        ' 文字列長の配列を作成（ヒューリスティック計算用）
        Dim srclen = New Integer(source.Length) {}
        For i As Integer = source.Length - 1 To 0 Step -1
            srclen(i) = srclen(i + 1) + source(i).Length
        Next

        Dim destlen = New Integer(destination.Length) {}
        For i As Integer = destination.Length - 1 To 0 Step -1
            destlen(i) = destlen(i + 1) + destination(i).Length
        Next

        ' A*アルゴリズムの初期化
        Dim startPos As New CostPosition(Nothing, 0, 0, 0, 0, Nothing)

        Dim open As New SortedSet(Of CostPosition)()
        Dim closed As New SortedSet(Of CostPosition)()
        open.Add(New CostPosition(Nothing, 0, 0, 0, 0, Nothing))

        Dim order As New BPlusTree(Of CostPosition)(AddressOf CostPositionComparer)
        order.Insert(startPos)

        ' A*探索のメインループ
        Dim answer As CostPosition = Nothing
        Do While open.Count > 0
            Dim cur = order(0)
            order.Remove(cur)

            open.Remove(cur)
            closed.Add(cur)

            ' ゴール状態に到達したかチェック
            If cur.X >= source.Length AndAlso cur.Y >= destination.Length Then
                answer = cur
                Exit Do
            End If

            ' 隣接する状態を生成
            If cur.X < source.Length AndAlso cur.Y < destination.Length Then
                Dim editChars As New List(Of EditChar)()
                Dim editCost As Integer = ASterCharDiff(source(cur.X), destination(cur.Y), editChars)
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 1, editCost, srclen, destlen, editChars.ToArray()))
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, source(cur.X).Length, srclen, destlen, Nothing))
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, destination(cur.Y).Length, srclen, destlen, Nothing))
            Else
                If cur.X < source.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, source(cur.X).Length, srclen, destlen, Nothing))
                End If
                If cur.Y < destination.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, destination(cur.Y).Length, srclen, destlen, Nothing))
                End If
            End If
        Loop

        ' 結果をトレースバックして構築
        Dim ans As New List(Of Answer)()
        Do While answer.From IsNot Nothing
            ans.Add(New Answer(source, answer.From.X, answer.X, destination, answer.From.Y, answer.Y, answer.EditChars))
            answer = answer.From
        Loop
        ans.Reverse()
        Return ans
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
    Private Function ASterCharDiff(source As String, destination As String, editChars As List(Of EditChar)) As Integer
        Dim startPos As New CostPosition(Nothing, 0, 0, 0, 0, Nothing)

        Dim open As New SortedSet(Of CostPosition)()
        Dim closed As New SortedSet(Of CostPosition)()
        open.Add(startPos)

        Dim order As New BPlusTree(Of CostPosition)(AddressOf CostPositionComparer)
        order.Insert(startPos)

        Dim answer As CostPosition = Nothing
        Do While open.Count > 0
            Dim cur = order(0)
            order.Remove(cur)

            open.Remove(cur)
            closed.Add(cur)

            If cur.X >= source.Length AndAlso cur.Y >= destination.Length Then
                answer = cur
                Exit Do
            End If

            ' 文字が一致する場合はコスト0で移動
            If cur.X < source.Length AndAlso cur.Y < destination.Length AndAlso source(cur.X) = destination(cur.Y) Then
                UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 1, 0, source.Length - cur.X, destination.Length - cur.Y))
            Else
                ' 文字が一致しない場合は挿入/削除でコスト1
                If cur.Y < destination.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 0, 1, 1, source.Length - cur.X, destination.Length - cur.Y))
                End If
                If cur.X < source.Length Then
                    UpdatePosition(open, closed, order, CreateNewPosition(cur, 1, 0, 1, source.Length - cur.X, destination.Length - cur.Y))
                End If
            End If
        Loop

        ' 編集操作の数をカウント
        Dim ans As Integer = 0
        Do While answer.From IsNot Nothing
            Dim edit As EditChar
            If answer.From.X < answer.X AndAlso answer.From.Y = answer.Y Then
                ' 削除
                edit = New EditChar(EditTypeEnum.Delete, source(answer.From.X))
                ans += 1
            ElseIf answer.From.X = answer.X AndAlso answer.From.Y < answer.Y Then
                ' 挿入
                edit = New EditChar(EditTypeEnum.Insert, destination(answer.From.Y))
                ans += 1
            Else
                ' 一致
                edit = New EditChar(EditTypeEnum.Match, source(answer.From.X))
            End If
            editChars.Add(edit)

            answer = answer.From
        Loop

        editChars.Reverse()
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
        Dim r = a.TotalCost.CompareTo(b.TotalCost)
        If r <> 0 Then
            Return r
        End If

        r = a.Y.CompareTo(b.Y)
        If r <> 0 Then
            Return r
        End If

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
    Private Sub UpdatePosition(open As SortedSet(Of CostPosition), closed As SortedSet(Of CostPosition), order As BPlusTree(Of CostPosition), newCur As CostPosition)
        If Not closed.Contains(newCur) Then
            Dim oldCur As CostPosition = Nothing
            If open.TryGetValue(newCur, oldCur) Then
                If newCur.TotalCost < oldCur.TotalCost Then
                    open.Remove(oldCur)
                    open.Add(newCur)
                    order.Remove(oldCur)
                    order.Insert(newCur)
                End If
            Else
                open.Add(newCur)
                order.Insert(newCur)
            End If
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
    Private Function CreateNewPosition(current As CostPosition, moveX As Integer, moveY As Integer, arrivalCost As Integer, srclen As Integer(), destlen As Integer(), editChars() As EditChar) As CostPosition
        Dim x = current.X + moveX
        Dim y = current.Y + moveY
        Return New CostPosition(current, current.X + moveX, current.Y + moveY, current.ArrivalCost + arrivalCost, Math.Abs(srclen(x) - destlen(y)), editChars)
    End Function

    ''' <summary>
    ''' 新しいコスト位置を作成します（残り長さ指定）
    ''' </summary>
    ''' <param name="current">現在の位置</param>
    ''' <param name="moveX">X軸の移動量</param>
    ''' <param name="moveY">Y軸の移動量</param>
    ''' <param name="arrivalCost">到達コスト</param>
    ''' <param name="restX">X軸の残り長さ</param>
    ''' <param name="restY">Y軸の残り長さ</param>
    ''' <returns>新しいコスト位置</returns>
    Private Function CreateNewPosition(current As CostPosition, moveX As Integer, moveY As Integer, arrivalCost As Integer, restX As Integer, restY As Integer) As CostPosition
        Return New CostPosition(current, current.X + moveX, current.Y + moveY, current.ArrivalCost + arrivalCost, Math.Abs(restX - restY), Nothing)
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
        Public ReadOnly Property Source As String

        ''' <summary>デスティネーション文字列</summary>
        Public ReadOnly Property Destination As String

        ''' <summary>編集操作の詳細（主にA*アルゴリズムでの置換操作の内容を格納）</summary>
        Private editChars As EditChar()

        Public ReadOnly Property EditString As String
            Get
                Dim res As New StringBuilder()

                Dim mode = EditTypeEnum.Match
                For Each edit In Me.EditChars
                    If edit.EditType <> mode Then
                        If mode <> EditTypeEnum.Match Then
                            res.Append("}")
                        End If
                        Select Case edit.EditType
                            Case EditTypeEnum.Match
                                res.Append(EscapeString(edit.EditChar))
                            Case EditTypeEnum.Delete
                                res.Append("{D:" + EscapeString(edit.EditChar))
                            Case EditTypeEnum.Insert
                                res.Append("{I:" + EscapeString(edit.EditChar))
                        End Select
                        mode = edit.EditType
                    Else
                        res.Append(EscapeString(edit.EditChar))
                    End If
                Next
                If mode <> EditTypeEnum.Match Then
                    res.Append("}")
                End If
                Return res.ToString()
            End Get
        End Property

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

        ''' <summary>
        ''' 編集タイプと文字列を指定してAnswerを初期化します
        ''' </summary>
        ''' <param name="EditType">編集操作の種類</param>
        ''' <param name="src">ソース文字列</param>
        ''' <param name="dest">デスティネーション文字列</param>
        Public Sub New(EditType As EditTypeEnum, src As String, dest As String)
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
        Public Sub New(source() As String, prevX As Integer, nextX As Integer, destination() As String, prevY As Integer, nextY As Integer, editChars As EditChar())
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
                    Me.Source = ""
                    Me.Destination = destination(prevY)
                Case EditTypeEnum.Delete
                    Me.Source = source(prevX)
                    Me.Destination = ""
            End Select

            ' 編集情報
            Me.EditChars = editChars
        End Sub

        ''' <summary>
        ''' 差分操作の文字列表現を返します
        ''' </summary>
        ''' <returns>操作を表すフォーマットされた文字列</returns>
        ''' <remarks>
        ''' "  " : 一致、"+ " : 挿入、"- " : 削除、"R " : 置換
        ''' </remarks>
        Overrides Function ToString() As String
            Select Case Me.EditType
                Case EditTypeEnum.Match
                    Return $"  {Me.Source}"
                Case EditTypeEnum.Insert
                    Return $"+ {Me.Destination}"
                Case EditTypeEnum.Delete
                    Return $"- {Me.Source}"
                Case EditTypeEnum.Diff
                    Return $"R {Me.Source} - {Me.Destination}"
                Case Else
                    Return ""
            End Select
        End Function
    End Class

End Module
