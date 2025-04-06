Option Explicit

Dim i
Dim arg

For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments.Item(i)
    WScript.StdErr.WriteLine("ˆø”:" & arg)
Next
