Option Explicit

Dim ioMode
Dim i
Dim arg
Dim targetFileName
Dim fso
Dim file

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


Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

    Set file = fso.OpenTextFile(targetFileName, ioMode, true)
    If Err.Number <> 0 Then
        WScript.StdErr.WriteLine("Error : " & Err.Number & _
            Err.Description & ":" & _
            fso.GetAbsolutePathName(targetFileName))
    End If

On Error Goto 0

Dim stdin, line
Set stdin = WScript.Stdin
Do Until stdin.AtEndOfStream
    line = stdin.ReadLine()
    Call  WScript.Echo(line)
    Call  file.WriteLine(line)
Loop
