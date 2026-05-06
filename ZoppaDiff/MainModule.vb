Option Strict On
Option Explicit On

Imports System.IO
Imports ZoppaDiff.IOHelper
Imports ZoppaDiff.ApplicationSwitch

''' <summary>アプリケーションのエントリーポイントを提供します。</summary>
Public Module MainModule

    ''' <summary>エンコードスイッチ名です。</summary>
    Public Const ENCODE_SW As String = "encode"

    ''' <summary>
    ''' エントリポイント。
    ''' </summary>
    Sub Main()
        Dim analysis = SwitchAnalyzer.Create().
            SetDescription("シンプルな Diffコマンド。").
            SetRemark("簡単な Diffを行います").
            SetAuthor("zoppa software").
            SetMailAddress("zoppa@ab.auone-net.jp").
            SetequiredParameter(True).
            SetValueName("source file path] [destination file path").
            SetSwitch(ENCODE_SW, "e"c, "encode", valueCount:=1, valueName:="encode", description:="文字コードを指定する(shift-jis,UTF-8など)").
            Parse()

        ' Help、Version指定、もしくはエラーならば終了
        If analysis.IsHelp OrElse analysis.IsVersion OrElse analysis.IsError Then
            System.Environment.Exit(0)
        End If

        ' エンコード指定を読み込む
        Dim encode As Text.Encoding = Text.Encoding.Default
        If analysis.HasSwitch(ENCODE_SW) Then
            encode = GetEncoder(analysis.GetSwitchParameter(ENCODE_SW)(0))
        End If

        ' 引数を読み込む
        Dim parameters = analysis.GetParameters()
        If parameters.Length < 2 Then
            WriteErrorLine("引数が足りません。")
            System.Environment.Exit(1)
        End If
        Dim sourceStrs = ReadLines(parameters(0), encode, "比較元ファイルの読込に失敗しました")
        Dim destinationStrs = ReadLines(parameters(1), encode, "比較先ファイルの読込に失敗しました")

        ' Diffを実行
        Dim diff = DiffModule.Diff(sourceStrs, destinationStrs)
        For Each ln In diff
            Select Case ln.EditType
                Case DiffModule.EditTypeEnum.Insert
                    Console.WriteLine("+ :" & ln.Destination)
                Case DiffModule.EditTypeEnum.Delete
                    Console.WriteLine("- :" & ln.Source)
                Case DiffModule.EditTypeEnum.Match
                    Console.WriteLine("   " & ln.Source)
                Case DiffModule.EditTypeEnum.Diff
                    Console.WriteLine("Rs:" & ln.Source)
                    Console.WriteLine(" d:" & ln.Destination)
            End Select
        Next
    End Sub

    ''' <summary>
    ''' 指定されたファイルから全ての行を読み込み、文字列配列として返します。
    ''' </summary>
    ''' <param name="path">読み込むファイルのパス。</param>
    ''' <param name="encode">ファイルの文字エンコーディング。Nothing の場合はデフォルトエンコーディングが使用されます。</param>
    ''' <param name="errMessage">読み込みに失敗した場合に例外に含めるエラーメッセージ。</param>
    ''' <returns>ファイルから読み込まれた各行を格納した文字列配列。</returns>
    ''' <exception cref="IOException">ファイルの読み込み中にエラーが発生した場合にスローされます。</exception>
    Private Function ReadLines(path As String, encode As Text.Encoding, errMessage As String) As String()
        Try
            Using reader = If(encode IsNot Nothing, New StreamReader(path, encode), New StreamReader(path))
                Dim lines As New List(Of String)
                While Not reader.EndOfStream
                    lines.Add(reader.ReadLine())
                End While
                Return lines.ToArray()
            End Using
        Catch ex As Exception
            Throw New IOException(errMessage, ex)
        End Try
    End Function

End Module
