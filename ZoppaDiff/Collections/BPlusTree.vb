Option Strict On
Option Explicit On

Namespace Collections

    ''' <summary>
    ''' B+木の実装クラス
    ''' 
    ''' B+木は、データベースやファイルシステムなどで使用されるデータ構造であり、
    ''' 高速な検索、挿入、削除を提供します。このクラスは、B+木の基本的な操作を実装しています。
    ''' 
    ''' このクラスは、IComparableインターフェイスを実装する型Tに対して動作します。
    ''' </summary>
    ''' <typeparam name="T">対象の型。</typeparam>
    Public NotInheritable Class BPlusTree(Of T)
        Implements IEnumerable(Of T)

        ''' <summary>挿入タイプ。</summary>
        Private Enum InsertType
            ' 挿入しない
            None
            ' 枝を挿入する
            Branch
            ' 葉を挿入する
            Leaf
        End Enum

        ' B+木の要素数
        Private ReadOnly M As Integer = 4

        ' B+木のルートノード
        Private _rootBranch As Branch

        ' B+木の最初の葉ノード
        Private ReadOnly _firstLeaf As Leaf

        ' 格納されているデータ数
        Private _count As Integer

        ' ロックオブジェクト
        Private ReadOnly _lock As New Object()

        ' 比較メソッド
        Private ReadOnly compareTo As Func(Of T, T, Integer)

        ''' <summary>
        ''' B木のデータ数を取得するプロパティ。
        ''' このプロパティは、B木に格納されている要素の数を返します。
        ''' </summary>
        ''' <returns>データ数。</returns>
        Public ReadOnly Property Count As Integer
            Get
                SyncLock Me._lock
                    Return Me._count
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' B+木の要素にアクセスするインデクサー。
        ''' このインデクサーは、指定されたインデックスに対応する要素を返します。
        ''' </summary>
        ''' <param name="index">取得する要素のインデックス。</param>
        ''' <returns>指定されたインデックスに対応する要素。</returns>
        Default Public ReadOnly Property Item(index As Integer) As T
            Get
                Dim current As Leaf = Me._firstLeaf
                Dim count As Integer = 0
                While current IsNot Nothing
                    If index < count + current.Count Then
                        Return current.values(index - count)
                    End If
                    count += current.Count
                    current = current.NextLeaf
                End While
                Throw New IndexOutOfRangeException()
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="compareTo">値の比較関数。</param>
        Public Sub New(compareTo As Func(Of T, T, Integer))
            Me.M = 4
            Me._rootBranch = Nothing
            Me._firstLeaf = New Leaf(Me.M)
            Me.compareTo = compareTo
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="blockSize">ブロックサイズ指定</param>
        ''' <param name="compareTo">値の比較関数。</param>
        Public Sub New(blockSize As UShort, compareTo As Func(Of T, T, Integer))
            Me.M = blockSize \ 2
            Me._rootBranch = Nothing
            Me._firstLeaf = New Leaf(Me.M)
            Me.compareTo = compareTo
        End Sub

        ''' <summary>
        ''' B+木に要素を挿入するメソッド。
        ''' </summary>
        ''' <param name="value">挿入する値。</param>
        Public Sub Insert(value As T)
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value), "挿入する値はnullにできません")
            End If

            ' 挿入操作の実装
            Dim update As New UpdateInfo With {
                .Done = False,
                .NewNode = Nothing
            }

            SyncLock Me._lock
                If Me._rootBranch Is Nothing Then
                    ' ルートノードがnullの場合、葉要素に値を追加
                    If Not InsertIntoLeaf(Me._firstLeaf, value, update).Done Then
                        Dim newNode = New Branch(Me.M) With {
                        .Count = 1
                    }
                        newNode.children(0) = Me._firstLeaf
                        newNode.children(1) = update.NewNode
                        newNode.Count = 2
                        Me._rootBranch = newNode
                    End If
                ElseIf Not InsertIntoBranch(Me._rootBranch, value, update).Done Then
                    ' ルートノードが存在する場合、枝に値を追加
                    Dim newNode = New Branch(Me.M) With {
                    .Count = 1
                }
                    newNode.children(0) = Me._rootBranch
                    newNode.children(1) = update.NewNode
                    newNode.Count = 2
                    Me._rootBranch = newNode
                End If

                ' 数を更新
                Me._count += 1
            End SyncLock
        End Sub

        ''' <summary>
        ''' B+木の葉ノードに値を挿入するメソッド。
        ''' このメソッドは、葉ノードに新しい値を挿入し、必要に応じて新しいノードを作成します。
        ''' </summary>
        ''' <param name="leaf">葉要素。</param>
        ''' <param name="value">挿入する値。</param>
        ''' <param name="update">更新情報。</param>
        ''' <returns>更新情報。</returns>
        Private Function InsertIntoLeaf(leaf As Leaf, value As T, update As UpdateInfo) As UpdateInfo
            If leaf.Count = 0 Then
                ' 最初の要素を挿入
                leaf.values(0) = value
                leaf.Count = 1
                update.Done = True
            Else
                ' 指定値が挿入される位置を取得
                Dim inc As Integer = 0
                While inc < leaf.Count AndAlso compareTo(leaf.values(inc), value) < 0
                    inc += 1
                End While

                If leaf.Count < leaf.values.Length Then
                    ' 葉要素の挿入位置に値を挿入する
                    If inc < leaf.Count Then
                        Array.ConstrainedCopy(leaf.values, inc, leaf.values, inc + 1, leaf.Count - inc)
                    End If
                    leaf.values(inc) = value
                    leaf.Count += 1
                    update.Done = True
                Else
                    ' 葉要素が満杯の場合、分割する
                    Dim newLeaf = Me.SplitLeaf(leaf, value, inc)
                    newLeaf.NextLeaf = leaf.NextLeaf
                    newLeaf.PrevLeaf = leaf
                    If leaf.NextLeaf IsNot Nothing Then
                        leaf.NextLeaf.PrevLeaf = newLeaf
                    End If
                    leaf.NextLeaf = newLeaf

                    update.Done = False
                    update.NewNode = newLeaf
                End If
            End If
            Return update
        End Function

        ''' <summary>
        ''' 葉要素を分割して新しい葉要素を生成します。
        ''' </summary>
        ''' <param name="leaf">分割する葉要素。</param>
        ''' <param name="value">挿入する値。</param>
        ''' <param name="inc">挿入位置のインデックス。</param>
        ''' <returns>新しい葉要素。</returns>
        Private Function SplitLeaf(leaf As Leaf, value As T, inc As Integer) As Leaf
            Dim newLeaf As New Leaf(Me.M)
            If inc < Me.M + 1 Then
                ' 後半部をコピー
                newLeaf.Count = Me.M + 1
                Array.Copy(leaf.values, Me.M, newLeaf.values, 0, leaf.values.Length - Me.M)

                ' 前半部に値を挿入する
                leaf.Count = M + 1
                If inc < Me.M Then
                    Array.ConstrainedCopy(leaf.values, inc, leaf.values, inc + 1, Me.M - inc)
                End If
                leaf.values(inc) = value
            Else
                ' 後半部に値を挿入する
                Dim splitIndex As Integer = inc - (Me.M + 1)
                newLeaf.Count = Me.M + 1

                ' 挿入位置までの値をコピー
                If splitIndex > 0 Then
                    Array.Copy(leaf.values, Me.M + 1, newLeaf.values, 0, splitIndex)
                End If

                ' 挿入位置に値を挿入
                newLeaf.values(splitIndex) = value

                ' 挿入位置以降の値をコピー
                Array.Copy(leaf.values, inc, newLeaf.values, splitIndex + 1, leaf.values.Length - inc)

                ' 前半部分は変更なし
                leaf.Count = M + 1
            End If
#If DEBUG Then
            For d As Integer = M + 1 To leaf.values.Length - 1
                leaf.values(d) = CType(Nothing, T)
            Next
#End If
            Return newLeaf
        End Function

        ''' <summary>
        ''' B+木の枝に値を挿入するメソッド。
        ''' このメソッドは、指定された枝に値を挿入し、必要に応じて新しい枝や葉を作成します。
        ''' </summary>
        ''' <param name="branch">枝要素。</param>
        ''' <param name="value">挿入する値。</param>
        ''' <param name="update">更新情報。</param>
        ''' <returns>更新情報。</returns>
        Private Function InsertIntoBranch(branch As Branch, value As T, update As UpdateInfo) As UpdateInfo
            ' 指定値が挿入される位置を取得
            Dim inc As Integer = 0
            While inc < branch.Count - 1 AndAlso compareTo(TraverseHead(branch.children(inc + 1)), value) < 0
                inc += 1
            End While

            ' 枝、葉の挿入位置に値を挿入する
            If TypeOf branch.children(inc) Is Branch Then
                Me.InsertIntoBranch(CType(branch.children(inc), Branch), value, update)
            Else
                Me.InsertIntoLeaf(CType(branch.children(inc), Leaf), value, update)
            End If

            If Not update.Done Then
                If branch.Count < branch.children.Length Then
                    ' 枝要素の挿入位置に値を挿入する
                    If inc + 1 < branch.Count Then
                        Array.ConstrainedCopy(branch.children, inc + 1, branch.children, inc + 2, branch.Count - (inc + 1))
                    End If
                    branch.children(inc + 1) = update.NewNode
                    branch.Count += 1
                    update.Done = True
                Else
                    ' 枝要素が満杯の場合、分割する
                    Dim newBranch = Me.SplitBranch(branch, update.NewNode, inc + 1)
                    update.Done = False
                    update.NewNode = newBranch
                End If
            End If

            Return update
        End Function

        ''' <summary>
        ''' 紐付く値の最初の要素を取得するヘルパーメソッド。
        ''' このメソッドは、B+木の枝をたどって最初の葉ノードに到達し、その値を返します。
        ''' </summary>
        ''' <param name="parts">要素のインターフェイス。</param>
        ''' <returns>最初の値。</returns>
        Private Shared Function TraverseHead(parts As IParts) As T
            While TypeOf parts Is Branch
                parts = CType(parts, Branch).children(0)
            End While
            Return CType(parts, Leaf).values(0)
        End Function

        ''' <summary>
        ''' 紐付く値の最後の要素を取得するヘルパーメソッド。
        ''' このメソッドは、B+木の枝をたどって最後の葉ノードに到達し、その値を返します。
        ''' </summary>
        ''' <param name="parts">要素のインターフェイス。</param>
        ''' <returns>最後の値。</returns>
        Private Shared Function TraverseLast(parts As IParts) As T
            While TypeOf parts Is Branch
                parts = CType(parts, Branch).children(parts.Count - 1)
            End While
            Return CType(parts, Leaf).values(parts.Count - 1)
        End Function

        ''' <summary>
        ''' 枝要素を分割して新しい葉要素を生成します。
        ''' </summary>
        ''' <param name="branch">分割する枝要素。</param>
        ''' <param name="insertParts">挿入する要素。</param>
        ''' <param name="inc">挿入位置のインデックス。</param>
        ''' <returns>新しい枝要素。</returns>
        Private Function SplitBranch(branch As Branch, insertParts As IParts, inc As Integer) As Branch
            Dim newBranch As New Branch(Me.M)
            If inc < Me.M + 1 Then
                ' 後半部をコピー
                newBranch.Count = Me.M + 1
                Array.Copy(branch.children, Me.M, newBranch.children, 0, branch.children.Length - Me.M)

                ' 前半部に値を挿入する
                branch.Count = M + 1
                If inc < Me.M Then
                    Array.ConstrainedCopy(branch.children, inc, branch.children, inc + 1, Me.M - inc)
                End If
                branch.children(inc) = insertParts
            Else
                ' 後半部に値を挿入する
                Dim splitIndex As Integer = inc - (Me.M + 1)
                newBranch.Count = Me.M + 1

                ' 挿入位置までの値をコピー
                If splitIndex > 0 Then
                    Array.Copy(branch.children, Me.M + 1, newBranch.children, 0, splitIndex)
                End If

                ' 挿入位置に値を挿入
                newBranch.children(splitIndex) = insertParts

                ' 挿入位置以降の値をコピー
                Array.Copy(branch.children, inc, newBranch.children, splitIndex + 1, branch.children.Length - inc)

                ' 前半部分は変更なし
                branch.Count = M + 1
            End If
#If DEBUG Then
            For d As Integer = M + 1 To branch.children.Length - 1
                branch.children(d) = Nothing
            Next
#End If
            Return newBranch
        End Function

        ''' <summary>
        ''' B+木に指定された値が存在するかどうかを確認するメソッド。
        ''' </summary>
        ''' <param name="value">確認する値。</param>
        ''' <returns>値が存在する場合は True、存在しない場合は False。</returns>
        Public Function Contains(value As T) As Boolean
            SyncLock Me._lock
                Dim lf = Me.SearchLeaf(value)
                Dim i As Integer = 0
                While i < lf.Count
                    Dim cmp = compareTo(lf.values(i), value)
                    If cmp = 0 Then
                        ' 値が見つかった場合、その値を返す
                        Return True
                    ElseIf cmp > 0 Then
                        ' 値が大きい場合、探索を終了
                        Exit While
                    End If
                    i += 1
                End While
                Return False
            End SyncLock
        End Function

        ''' <summary>
        ''' B+木から指定された値を検索するメソッド。
        ''' このメソッドは、B+木から指定された値を検索し、見つかった場合はその値を返します。
        ''' </summary>
        ''' <param name="value">検索する値。</param>
        ''' <returns>見つかった値。見つからない場合は Nothing。</returns>
        Public Function Search(value As T) As T
            SyncLock Me._lock
                Dim lf = Me.SearchLeaf(value)
                Dim i As Integer = 0
                While i < lf.Count
                    Dim cmp = compareTo(lf.values(i), value)
                    If cmp = 0 Then
                        ' 値が見つかった場合、その値を返す
                        Return lf.values(i)
                    ElseIf cmp > 0 Then
                        ' 値が大きい場合、探索を終了
                        Exit While
                    End If
                    i += 1
                End While
                Return CType(Nothing, T)
            End SyncLock
        End Function

        ''' <summary>
        ''' ルート枝から指定された値を検索し、対応する葉を返すメソッド。
        ''' このメソッドは、B+木のルート枝から指定された値を検索し、最終的に対応する葉を返します。
        ''' </summary>
        ''' <param name="value">検索する値。</param>
        ''' <returns>葉要素。</returns>
        Private Function SearchLeaf(value As T) As Leaf
            ' ルートノードがnullの場合、最初の葉を返す
            If Me._rootBranch Is Nothing Then
                Return Me._firstLeaf
            End If

            ' ルート枝から値を検索
            Dim ptr As IParts = Me._rootBranch
            While TypeOf ptr Is Branch
                ' 指定値が枝に存在するか確認
                Dim brh As Branch = CType(ptr, Branch)
                Dim p As Integer = brh.Count - 1
                While p > 0 AndAlso compareTo(TraverseLast(brh.children(p - 1)), value) >= 0
                    p -= 1
                End While

                ' 枝の子要素をたどる
                ptr = brh.children(p)
                If TypeOf ptr Is Leaf Then
                    Exit While
                End If
            End While

            Return CType(ptr, Leaf)
        End Function

        ''' <summary>
        ''' B+木から指定された値を削除するメソッド。
        ''' このメソッドは、B+木から指定された値を削除し、必要に応じてノードの再構築を行います。
        ''' </summary>
        ''' <param name="value">削除する値。</param>
        ''' <remarks>現在は未実装</remarks>
        Public Sub Remove(value As T)
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value), "挿入する値はnullにできません")
            End If
            ' 削除操作の実装
            Dim delete As New DeleteInfo With {
                .Done = False,
                .NeedRebuild = False
            }
            SyncLock Me._lock
                If Me._rootBranch Is Nothing Then
                    ' ルートノードがnullの場合、葉要素から値を削除
                    Me.DeleteIntoLeaf(Me._firstLeaf, value, delete)

                ElseIf Me.DeleteIntoBranch(Me._rootBranch, value, delete).Done Then
                    ' ルートノードが空になった場合は、ルートを子要素で置き換える
                    If Me._rootBranch.Count <= 1 Then
                        Me._rootBranch = TryCast(Me._rootBranch.children(0), Branch)
                    End If
                End If

                If delete.Done Then
                    ' 削除が成功した場合は、カウントを減らす
                    Me._count -= 1
                Else
                    ' 要素がなく、削除できなかった場合はエラーを返す
                    Throw New BtreeException("指定された値は存在せず、削除できません")
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' B+木の葉に値を削除するメソッド。
        ''' このメソッドは、指定された葉から値を削除し、必要に応じてノードの再構築を行います。
        ''' </summary>
        ''' <param name="leaf">葉要素。</param>
        ''' <param name="value">削除する値。</param>
        ''' <param name="delete">削除情報。</param>
        ''' <returns>削除情報。</returns>
        Private Function DeleteIntoLeaf(leaf As Leaf, value As T, delete As DeleteInfo) As DeleteInfo
            ' 削除する値の位置を探す
            Dim del As Integer = 0
            While del < leaf.Count AndAlso compareTo(leaf.values(del), value) < 0
                del += 1
            End While

            If del < leaf.Count AndAlso compareTo(leaf.values(del), value) = 0 Then
                ' 値を削除し、右側の値を左にシフト
                Array.ConstrainedCopy(leaf.values, del + 1, leaf.values, del, leaf.Count - del - 1)

                ' ノードの値の数を減らす、小さくなりすぎたら needRebuild フラグを立てる
                leaf.Count -= 1
#If DEBUG Then
                leaf.values(leaf.Count) = CType(Nothing, T) ' デバッグ用に値をクリア
#End If
                delete.Done = True
                delete.NeedRebuild = leaf.Count <= M
            End If
            Return delete
        End Function

        ''' <summary>
        ''' B+木の枝から値を削除するメソッド。
        ''' このメソッドは、指定された枝から値を削除し、必要に応じてノードの再構築を行います。
        ''' </summary>
        ''' <param name="branch">ルート枝要素。</param>
        ''' <param name="value">削除する値。</param>
        ''' <param name="delete">削除情報。</param>
        ''' <returns>削除情報。</returns>
        Private Function DeleteIntoBranch(branch As Branch, value As T, delete As DeleteInfo) As DeleteInfo
            ' 指定値が挿入される位置を取得
            Dim del As Integer = branch.Count - 1
            While del > 0 AndAlso compareTo(TraverseLast(branch.children(del - 1)), value) >= 0
                del -= 1
            End While

            ' 枝、葉の削除位置から値を削除する
            If TypeOf branch.children(del) Is Branch Then
                If Me.DeleteIntoBranch(CType(branch.children(del), Branch), value, delete).NeedRebuild Then
                    Me.BalanceBranch(branch, del, delete)
                End If
            Else
                If Me.DeleteIntoLeaf(CType(branch.children(del), Leaf), value, delete).NeedRebuild Then
                    Me.BalanceLeaf(branch, del, delete)
                End If
            End If

            Return delete
        End Function

        ''' <summary>
        ''' B+木の枝をバランスさせるメソッド。
        ''' このメソッドは、指定された枝の左右の枝をバランスさせます。
        ''' </summary>
        ''' <param name="branch">枝要素。</param>
        ''' <param name="del">削除位置のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub BalanceBranch(branch As Branch, del As Integer, delete As DeleteInfo)
            If del > 0 AndAlso branch.children(del - 1).Count + branch.children(del).Count > Me.M * 2 + 1 Then
                ' 削除する枝と左の枝とをバランス
                Me.BalanceBranch(CType(branch.children(del - 1), Branch), CType(branch.children(del), Branch))
                delete.NeedRebuild = False
            ElseIf del < branch.Count - 1 AndAlso branch.children(del).Count + branch.children(del + 1).Count > Me.M * 2 + 1 Then
                ' 削除する枝と右の枝とをバランス
                Me.BalanceBranch(CType(branch.children(del), Branch), CType(branch.children(del + 1), Branch))
                delete.NeedRebuild = False
            ElseIf del > 0 Then
                ' 削除する枝と左の枝とをマージ
                Dim leftBranch = CType(branch.children(del - 1), Branch)
                Dim rightBranch = CType(branch.children(del), Branch)
                Array.Copy(rightBranch.children, 0, leftBranch.children, leftBranch.Count, rightBranch.Count)
                leftBranch.Count += rightBranch.Count

                ' マージした枝分を削除して、移動
                If del + 1 < branch.Count Then
                    Array.ConstrainedCopy(branch.children, del + 1, branch.children, del, branch.Count - (del + 1))
                End If
                branch.Count -= 1
#If DEBUG Then
                branch.children(branch.Count) = Nothing
#End If
            Else
                ' 削除する枝と右の枝とをマージ
                Dim leftBranch = CType(branch.children(del), Branch)
                Dim rightBranch = CType(branch.children(del + 1), Branch)
                Array.Copy(rightBranch.children, 0, leftBranch.children, leftBranch.Count, rightBranch.Count)
                leftBranch.Count += rightBranch.Count

                ' マージした枝分を削除して、移動
                If del + 2 < branch.Count Then
                    Array.ConstrainedCopy(branch.children, del + 2, branch.children, del + 1, branch.Count - (del + 2))
                End If
                branch.Count -= 1
#If DEBUG Then
                branch.children(branch.Count) = Nothing
#End If
            End If
        End Sub

        ''' <summary>
        ''' B+木の葉を持つ枝をバランスさせるメソッド。
        ''' このメソッドは、指定された枝の左右の枝をバランスさせます。
        ''' </summary>
        ''' <param name="branch">枝要素。</param>
        ''' <param name="del">削除位置のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub BalanceLeaf(branch As Branch, del As Integer, delete As DeleteInfo)
            If del > 0 AndAlso branch.children(del - 1).Count + branch.children(del).Count > Me.M * 2 + 1 Then
                ' 削除する葉と左の葉とをバランス
                Me.BalanceLeaf(CType(branch.children(del - 1), Leaf), CType(branch.children(del), Leaf))
                delete.NeedRebuild = False
            ElseIf del < branch.Count - 1 AndAlso branch.children(del).Count + branch.children(del + 1).Count > Me.M * 2 + 1 Then
                ' 削除する葉と右の葉とをバランス
                Me.BalanceLeaf(CType(branch.children(del), Leaf), CType(branch.children(del + 1), Leaf))
                delete.NeedRebuild = False
            ElseIf del > 0 Then
                ' 削除する葉と左の葉とをマージ
                Dim leftLeaf As Leaf = CType(branch.children(del - 1), Leaf)
                Dim rightLeaf As Leaf = CType(branch.children(del), Leaf)
                Array.Copy(rightLeaf.values, 0, leftLeaf.values, leftLeaf.Count, rightLeaf.Count)
                leftLeaf.Count += rightLeaf.Count
                leftLeaf.NextLeaf = rightLeaf.NextLeaf

                ' マージした葉分を削除して、移動
                If del + 1 < branch.Count Then
                    Array.ConstrainedCopy(branch.children, del + 1, branch.children, del, branch.Count - (del + 1))
                End If
                branch.Count -= 1
#If DEBUG Then
                branch.children(branch.Count) = Nothing
#End If
            Else
                ' 削除する葉と右の葉とをマージ
                Dim leftLeaf = CType(branch.children(del), Leaf)
                Dim rightLeaf = CType(branch.children(del + 1), Leaf)
                Array.Copy(rightLeaf.values, 0, leftLeaf.values, leftLeaf.Count, rightLeaf.Count)
                leftLeaf.Count += rightLeaf.Count
                leftLeaf.NextLeaf = rightLeaf.NextLeaf
                If rightLeaf.NextLeaf IsNot Nothing Then
                    rightLeaf.NextLeaf.PrevLeaf = leftLeaf
                End If

                ' マージした葉分を削除して、移動
                If del + 2 < branch.Count Then
                    Array.ConstrainedCopy(branch.children, del + 2, branch.children, del + 1, branch.Count - (del + 2))
                End If
                branch.Count -= 1
#If DEBUG Then
                branch.children(branch.Count) = Nothing
#End If
            End If
        End Sub

        ''' <summary>
        ''' B+木の葉をバランスさせるメソッド。
        ''' このメソッドは、指定された葉の左右の葉をバランスさせます。
        ''' </summary>
        ''' <param name="leftLeaf">左の葉要素。</param>
        ''' <param name="rightLeaf">右の葉要素。</param>
        Private Sub BalanceLeaf(leftLeaf As Leaf, rightLeaf As Leaf)
            Dim split = (leftLeaf.Count + rightLeaf.Count) \ 2
            If leftLeaf.Count > split Then
                ' 左の葉が大きい場合、右に値を移動
                Dim moveCount = leftLeaf.Count - split

                ' 右の葉の値を左に移動
                Array.ConstrainedCopy(rightLeaf.values, 0, rightLeaf.values, moveCount, rightLeaf.Count)

                ' 左の葉の値を移動した数分、右に移動
                Array.Copy(leftLeaf.values, split, rightLeaf.values, 0, moveCount)

                ' 個数を更新
                rightLeaf.Count += moveCount
                leftLeaf.Count = split
#If DEBUG Then
                For d As Integer = split To leftLeaf.values.Length - 1
                    leftLeaf.values(d) = CType(Nothing, T)
                Next
#End If
            ElseIf rightLeaf.Count > split Then
                ' 右の葉が大きい場合、左に値を移動
                Dim moveCount = rightLeaf.Count - split

                ' 右の葉の値を左に移動
                Array.Copy(rightLeaf.values, 0, leftLeaf.values, leftLeaf.Count, moveCount)

                ' 右の葉の値を移動した数分、左に移動
                Array.ConstrainedCopy(rightLeaf.values, moveCount, rightLeaf.values, 0, rightLeaf.Count - moveCount)

                ' 個数を更新
                leftLeaf.Count += moveCount
                rightLeaf.Count = split
#If DEBUG Then
                For d As Integer = split To rightLeaf.values.Length - 1
                    rightLeaf.values(d) = CType(Nothing, T)
                Next
#End If
            End If
        End Sub

        ''' <summary>
        ''' B+木の枝をバランスさせるメソッド。
        ''' このメソッドは、指定された枝の左右の枝をバランスさせます。
        ''' </summary>
        ''' <param name="leftBranch">左の枝要素。</param>
        ''' <param name="rightBranch">右の枝要素。</param>
        Private Sub BalanceBranch(leftBranch As Branch, rightBranch As Branch)
            Dim split = (leftBranch.Count + rightBranch.Count) \ 2
            If leftBranch.Count > split Then
                ' 左の葉が大きい場合、右に値を移動
                Dim moveCount = leftBranch.Count - split

                ' 右の枝の子要素を左に移動
                Array.ConstrainedCopy(rightBranch.children, 0, rightBranch.children, moveCount, rightBranch.Count)

                ' 左の枝の子要素を移動した数分、右に移動
                Array.Copy(leftBranch.children, split, rightBranch.children, 0, moveCount)

                ' 個数を更新
                rightBranch.Count += moveCount
                leftBranch.Count = split
#If DEBUG Then
                For d As Integer = split To leftBranch.children.Length - 1
                    leftBranch.children(d) = Nothing
                Next
#End If
            ElseIf rightBranch.Count > split Then
                ' 右の枝が大きい場合、左に値を移動
                Dim moveCount = rightBranch.Count - split

                ' 右の枝の子要素を左に移動
                Array.Copy(rightBranch.children, 0, leftBranch.children, leftBranch.Count, moveCount)

                ' 右の枝の子要素を移動した数分、左に移動
                Array.ConstrainedCopy(rightBranch.children, moveCount, rightBranch.children, 0, rightBranch.Count - moveCount)

                ' 個数を更新
                leftBranch.Count += moveCount
                rightBranch.Count = split
#If DEBUG Then
                For d As Integer = split To rightBranch.children.Length - 1
                    rightBranch.children(d) = Nothing
                Next
#End If
            End If
        End Sub

        ''' <summary>
        ''' B+木の要素を列挙するためのメソッド。
        ''' このメソッドは、B+木の要素を列挙するための IEnumerator を返します。
        ''' </summary>
        ''' <returns>B+木の要素を列挙する IEnumerator。</returns>
        Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
            Return New BtreeEnumerator(Me)
        End Function

        ''' <summary>
        ''' B+木の要素を列挙するための非ジェネリックな IEnumerator を返すメソッド。
        ''' </summary>
        ''' <returns>B+木の要素を列挙する非ジェネリックな IEnumerator。</returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        ''' <summary>
        ''' B+木の列挙子を表す内部クラス。
        ''' このクラスは、B+木の要素を列挙するための IEnumerator を実装します。
        ''' </summary>
        Private NotInheritable Class BtreeEnumerator
            Implements IEnumerator(Of T)

            ' 対象の B+木
            Private ReadOnly _source As BPlusTree(Of T)

            ' 参照している葉
            Private _currentLeaf As Leaf

            ' インデックス
            Private _index As Integer = -1

            ' カレントのデータ
            Private _current As T

            ''' <summary>
            ''' 列挙子の現在の要素を取得します。
            ''' </summary>
            ''' <returns>現在の要素。</returns>
            Public ReadOnly Property Current As T Implements IEnumerator(Of T).Current
                Get
                    Return _current
                End Get
            End Property

            ''' <summary>
            ''' 列挙子の現在の要素を取得します（非ジェネリック）。
            ''' </summary>
            ''' <returns>現在の要素。</returns>
            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return _current
                End Get
            End Property

            ''' <summary>
            ''' コンストラクタ。
            ''' このコンストラクタは、指定された B+木を参照する列挙子を初期化します。
            ''' </summary>
            ''' <param name="source">対象の B+木。</param>
            Public Sub New(source As BPlusTree(Of T))
                Me._source = source
                Me._currentLeaf = source._firstLeaf
                Me._index = -1
            End Sub

            ''' <summary>列挙子を初期化します。</summary>
            Public Sub Reset() Implements IEnumerator.Reset
                Me._currentLeaf = Me._source._firstLeaf
                Me._index = -1
            End Sub

            ''' <summary>
            ''' 列挙子を現在の位置から次の位置に進めます。
            ''' </summary>
            ''' <returns>次の位置に進めた場合は True、それ以外は False。</returns>
            Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
                If Me._currentLeaf IsNot Nothing Then
                    Me._index += 1
                    If Me._index < Me._currentLeaf.Count Then
                        ' 現在の葉の値を返す
                        Me._current = Me._currentLeaf.values(Me._index)
                        Return True
                    Else
                        ' 次の葉に移動
                        Me._currentLeaf = Me._currentLeaf.NextLeaf
                        Me._index = 0
                        If Me._currentLeaf IsNot Nothing Then
                            Me._current = Me._currentLeaf.values(Me._index)
                            Return True
                        End If
                    End If
                End If
                Return False
            End Function

            ''' <summary>リソースを解放する。</summary>
            Public Sub Dispose() Implements IDisposable.Dispose
                ' 空実装
            End Sub

        End Class

        ''' <summary>
        ''' B+木の枝と葉を表すインターフェイス。
        ''' このインターフェイスは、B木のノードが持つべき基本的な機能を定義します。
        ''' </summary>
        Private Interface IParts

            ''' <summary>格納数を取得する。</summary>
            ''' <returns>格納数。</returns>
            Property Count As Integer

        End Interface

        ''' <summary>
        ''' B木の枝と葉を表す内部クラス。
        ''' このクラスは、B木のノードを表し、枝と葉の両方を格納します。
        ''' </summary>
        Private NotInheritable Class Branch
            Implements IParts

            ' 格納データ数
            Public Property Count As Integer Implements IParts.Count

            ' 枝、葉を格納するリスト
            Public children() As IParts

            ''' <summary>
            ''' コンストラクタ。
            ''' このコンストラクタは、指定されたブロックサイズに基づいてノードを初期化します。
            ''' </summary>
            ''' <param name="M">ブロックサイズ。</param>
            Public Sub New(M As Integer)
                Me.children = New IParts(M * 2) {}
            End Sub

        End Class

        ''' <summary>
        ''' B木の葉ノードを表す内部クラス。
        ''' このクラスは、B木の葉ノードを表し、データを格納します。
        ''' </summary>
        Private NotInheritable Class Leaf
            Implements IParts

            ' 格納データ数
            Public Property Count As Integer Implements IParts.Count

            ' 格納データ
            Public values() As T

            ' 前の葉ノードへの参照
            Public Property PrevLeaf As Leaf = Nothing

            ' 次の葉ノードへの参照
            Public Property NextLeaf As Leaf = Nothing

            ''' <summary>
            ''' コンストラクタ。
            ''' このコンストラクタは、指定されたブロックサイズに基づいて葉ノードを初期化します。
            ''' </summary>
            ''' <param name="M">ブロックサイズ。</param>
            Public Sub New(M As Integer)
                Me.values = New T(M * 2) {}
            End Sub
        End Class

        ''' <summary>
        ''' B+木の更新情報を表す内部クラス。
        ''' このクラスは、B+木の挿入操作の結果として生成される更新情報を格納します。
        ''' </summary>
        Private NotInheritable Class UpdateInfo

            ' 更新が行われたかどうか。
            Public Property Done As Boolean

            ' 更新後の新しく追加となったノード。
            Public Property NewNode As IParts

        End Class

        ''' <summary>
        ''' B+木の削除情報を表す内部クラス。
        ''' このクラスは、B+木の削除操作の結果として生成される削除情報を格納します。
        ''' </summary>
        Private NotInheritable Class DeleteInfo

            ' 削除が行われたかどうか。
            Public Property Done As Boolean

            ' 削除によりノードの再構築が必要かどうか。
            Public Property NeedRebuild As Boolean

        End Class

    End Class

End Namespace