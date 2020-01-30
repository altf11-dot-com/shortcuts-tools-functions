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

