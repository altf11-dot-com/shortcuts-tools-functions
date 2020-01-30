Attribute VB_Name = "sandbox"

Option Compare Database
Option Explicit

Sub test000()
    Dim x As String
    On Error GoTo err000_:
    Stop
    test001
    Stop
    x = 5 / 0
    Stop
    Exit Sub
err000_:
    Debug.Print Err.Description
    Resume Next
End Sub

Sub test001()
    Dim x As String
    On Error GoTo err001_:
    Stop
    Stop
    x = 5 / 0
    Stop
    Exit Sub
err001_:
    Debug.Print Err.Description
    'On Error Resume Next
    On Error GoTo 0
    Resume
End Sub

Sub OnErrorStatementDemo()
    Dim objRef As Object, msg As String
    On Error GoTo ErrorHandler ' Enable error-handling routine.
    Open "TESTFILE" For Output As #1 ' Open file for output.
    Kill "TESTFILE" ' Attempt to delete open
    ' file.
