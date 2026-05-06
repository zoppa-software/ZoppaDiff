Option Strict On
Option Explicit On

Namespace IOHelper

    ''' <summary>入出力モジュール。</summary>
    Public Module OutputMethods

        ''' <summary>標準出力にメッセージを出力する。</summary>
        ''' <param name="message">メッセージ。</param>
        Public Sub Write(message As String)
            If Not IsSilent Then
                Console.Out.Write(message)
            End If
        End Sub

        ''' <summary>標準出力にメッセージを出力する（改行付き）</summary>
        ''' <param name="message">メッセージ。</param>
        Public Sub WriteLine(message As String)
            If Not IsSilent Then
                Console.Out.WriteLine(message)
            End If
        End Sub

        ''' <summary>標準出力に改行を出力する。</summary>
        Public Sub WriteLine()
            If Not IsSilent Then
                Console.Out.WriteLine()
            End If
        End Sub

        ''' <summary>標準エラー出力にメッセージを出力する。</summary>
        ''' <param name="message">メッセージ。</param>
        Public Sub WriteError(message As String)
            Console.Error.Write(message)
        End Sub

        ''' <summary>標準エラー出力にメッセージを出力する（改行付き）</summary>
        ''' <param name="message">メッセージ。</param>
        Public Sub WriteErrorLine(message As String)
            Console.Error.WriteLine(message)
        End Sub

        ''' <summary>標準エラー出力に改行を出力する。</summary>
        Public Sub WriteErrorLine()
            Console.Error.WriteLine()
        End Sub

        ''' <summary>メッセージの出力を停止する。</summary>
        ''' <returns>メッセージを停止していたら真。</returns>
        Public Property IsSilent() As Boolean = False

    End Module

End Namespace
