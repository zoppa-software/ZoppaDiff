Option Strict On
Option Explicit On

Namespace ApplicationSwitch

    ''' <summary>コマンドラインスイッチを表します。</summary>
    Public NotInheritable Class CommandLineSwitch

        ''' <summary>スイッチの名前です。</summary>
        Public ReadOnly Property Name As String

        ''' <summary>短いスイッチです。</summary>
        Public Property SwitchChar As Char?

        ''' <summary>長いスイッチです。</summary>
        Public Property SwitchName As String

        ''' <summary>引数の数です（マイナス値なら引数の数は無制限です）</summary>
        Public Property ValueCount As Integer

        ''' <summary>引数の名前です（Helpで表示）</summary>
        Public Property ValueName As String

        ''' <summary>引数が必須ならば真です。</summary>
        Public Property Required As Boolean

        ''' <summary>スイッチの説明です。</summary>
        Public Property Description As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="name">スイッチ名。</param>
        Public Sub New(name As String)
            Me.Name = name
        End Sub

        ''' <summary>等価比較を行います。</summary>
        ''' <param name="obj">比較対象。</param>
        ''' <returns>比較結果。</returns>
        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            If obj IsNot Nothing AndAlso Me.GetType() Is obj.GetType() Then
                Return Me.Name = CType(obj, CommandLineSwitch).Name
            Else
                Return False
            End If
        End Function

        ''' <summary>ハッシュコード値を取得します。</summary>
        ''' <returns>ハッシュコード値。</returns>
        Public Overrides Function GetHashCode() As Integer
            Return Me.Name.GetHashCode()
        End Function

    End Class

End Namespace