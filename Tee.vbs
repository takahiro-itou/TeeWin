Option Explicit

Dim ioMode
Dim i
Dim arg
Dim targetFileName

ioMode = 2

For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments.Item(i)
    If Left(arg, 1) = "/" Or Left(arg, 1) = "-" Then
        Select Case arg
        Case "/a", "-a", "/append", "--append"
            ioMode = 8
        End Select
    Else
        targetFileName = arg
    End If
Next

WScript.StdErr.WriteLine("出力ファイル :" & targetFileName)
WScript.StdErr.WriteLine("ioMode = " & ioMode)
