Option Strict On
Option Explicit On

Namespace Collections

    ''' <summary>
    ''' B木の例外クラス。
    ''' 
    ''' このクラスは、B木の操作中に発生する可能性のある例外を表します。
    ''' 例えば、挿入や削除操作で不正な状態が検出された場合に使用されます。
    ''' </summary>
    Public NotInheritable Class BtreeException
        Inherits Exception

        ''' <summary>
        ''' BtreeException の新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="message">例外のメッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace
