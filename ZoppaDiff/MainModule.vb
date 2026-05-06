Option Strict On
Option Explicit On

Imports System.IO
Imports ZoppaDiff.IOHelper
Imports ZoppaDiff.ApplicationSwitch

''' <summary>アプリケーションのエントリーポイントを提供します。</summary>
Public Module MainModule

    ''' <summary>エンコードスイッチ名です。</summary>
    Public Const ENCODE_SW As String = "encode"

    ''' <summary>行番号スイッチ名です。</summary>
    Public Const LINE_SW As String = "line"

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
            SetSwitch(LINE_SW, "l"c, "line", valueCount:=0, valueName:="line", description:="行番号を表示する").
            Parse()

        ' Help、Version指定、もしくはエラーならば終了
        If analysis.IsHelp OrElse analysis.IsVersion OrElse analysis.IsError Then
            System.Environment.Exit(0)
        End If

        ' エンコード指定を読み込む
        Dim encode As Text.Encoding = Text.Encoding.UTF8
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
        WriteLine("比較元ファイル(s): " & parameters(0))
        WriteLine("比較先ファイル(d): " & parameters(1))
        WriteLine(" (s) <-> (d)")
        WriteLine("--- 比較結果 ---")
        Dim diff = DiffModule.Diff(sourceStrs, destinationStrs)
        If analysis.HasSwitch(LINE_SW) Then
            ' 行番号を表示する場合
            For Each ln In diff
                Select Case ln.EditType
                    Case DiffModule.EditTypeEnum.Insert
                        WriteLine(String.Format("<- (d){0:00000}|{1}", ln.Destination.Line, ln.Destination.Str))
                    Case DiffModule.EditTypeEnum.Delete
                        WriteLine(String.Format("-> (s){0:00000}|{1}", ln.Source.Line, ln.Source.Str))
                    Case DiffModule.EditTypeEnum.Match
                        WriteLine(String.Format("{0:00000}-{1:00000}|{2}", ln.Source.Line, ln.Destination.Line, ln.Source.Str))
                    Case DiffModule.EditTypeEnum.Diff
                        WriteLine(String.Format("Replace    |{0}", ln.EditString))
                        WriteLine(String.Format("   (s){0:00000}|{1}", ln.Source.Line, ln.Source.Str))
                        WriteLine(String.Format("   (d){0:00000}|{1}", ln.Destination.Line, ln.Destination.Str))
                End Select
            Next
        Else
            ' 行番号を表示しない場合（省略表示）
            For Each ln In diff
                Select Case ln.EditType
                    Case DiffModule.EditTypeEnum.Insert
                        WriteLine(String.Format("<-|{0}", ln.Destination.Str))
                    Case DiffModule.EditTypeEnum.Delete
                        WriteLine(String.Format("->|{0}", ln.Source.Str))
                    Case DiffModule.EditTypeEnum.Match
                        WriteLine(String.Format("   {0}", ln.Source.Str))
                    Case DiffModule.EditTypeEnum.Diff
                        WriteLine(String.Format("R |{0}", ln.EditString))
                        WriteLine(String.Format(" s|{0}", ln.Source.Str))
                        WriteLine(String.Format(" d|{0}", ln.Destination.Str))
                End Select
            Next
        End If
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
            Using reader = New StreamReader(path, encode)
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
