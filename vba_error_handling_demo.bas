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
    On Error GoTo 0 ' Turn off error trapping.
    On Error Resume Next ' Defer error trapping.
    objRef = GetObject("MyWord.Basic") ' Try to start nonexistent
    ' object, then test for
    'Check for likely Automation errors.
    If Err.Number = 440 Or Err.Number = 432 Or Err.Number = -2147221020 Then
    ' Tell user what happened. Then clear the Err object.
    msg = "There was an error attempting to open the Automation object!"
    MsgBox msg, , "Deferred Error Test"
    Err.Clear ' Clear Err object fields
    End If
    Exit Sub ' Exit to avoid handler.
ErrorHandler:     ' Error-handling routine.
    Select Case Err.Number ' Evaluate error number.
    Case 55 ' "File already open" error.
    Close #1 ' Close open file.
    Case Else
    ' Handle other situations here...
    End Select
    Resume ' Resume execution at same line
    ' that caused the error.
End Sub
