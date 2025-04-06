Option Explicit

Dim ioMode
Dim i
Dim arg
Dim bufSize
Dim targetFileName
Dim fso
Dim file

Const IO_MODE_WRITE  = 2
Const IO_MODE_APPEND = 8

ioMode = IO_MODE_WRITE
bufSize = 1

For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments.Item(i)
    If Left(arg, 1) = "/" Or Left(arg, 1) = "-" Then
        Select Case arg
        Case "/a", "-a", "/append", "--append"
            ioMode = IO_MODE_APPEND
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
    line = stdin.Read(bufSize)
    Call  WScript.StdOut.Write(line)
    Call  file.WriteLine(line)
Loop
