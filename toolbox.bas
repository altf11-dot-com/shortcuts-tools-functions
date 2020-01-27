Attribute VB_Name = "Shortcuts_Tools"
Option Explicit

Public Sub KS_CtrlShiftD()
    MoveCurrentRowDownOne
End Sub


Private Sub MoveCurrentRowDownOne()
    Rows(ActiveCell.row).Cut
    Rows(ActiveCell.row + 2).Insert Shift:=xlDown
    Cells(ActiveCell.row + 1, ActiveCell.Column).Activate
End Sub

Public Sub KS_CtrlShiftU()
    MoveCurrentRowUp
End Sub

Private Sub MoveCurrentRowUp()
    Dim currRow As Integer, currCol As Integer, TargetRow As Integer, MinRow As Integer
    Dim AutomaticMode, NextUpwardValue As String, MatchString As String
    currRow = ActiveCell.row
    currCol = ActiveCell.Column
    MatchString = ""
    MinRow = 3
    If currRow = MinRow Then
        MsgBox "Top of column ..."
        Exit Sub
    End If
    TargetRow = currRow - 1
    NextUpwardValue = Cells(TargetRow, currCol).Value
    If NextUpwardValue = "" Then
            While TargetRow > MinRow And NextUpwardValue = ""
                TargetRow = TargetRow - 1
                NextUpwardValue = Cells(TargetRow, currCol).Value
            Wend
            TargetRow = TargetRow + 1
    End If
    If MatchString <> "" Then
        If NextUpwardValue <> MatchString Then
            While TargetRow > MinRow And NextUpwardValue <> MatchString
                TargetRow = TargetRow - 1
                NextUpwardValue = Cells(TargetRow, currCol).Value
            Wend
            TargetRow = TargetRow + 1
        End If
    End If
    Rows(currRow).Cut
    Rows(TargetRow).Insert Shift:=xlDown
    Cells(TargetRow, currCol).Activate
End Sub

Public Sub KS_CtrlShiftT()
    DateTimeStamp
End Sub



Sub DateTimeStamp()
    Dim okToWrite As Long
    With ActiveCell
        If .Formula = vbNullString Then
            okToWrite = vbYes
        Else
            Beep
            okToWrite = MsgBox("Active cell is not empty, over-write?", vbYesNo, "Date/Time Stamp")
            If okToWrite = vbNo Then Exit Sub
        End If
        If okToWrite = vbYes Then .Value = Now
        .NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@"
        Columns(.Column).EntireColumn.AutoFit
    End With
End Sub

Public Sub KS_CtrlShiftF()
    AutofitAllColumnsAllRows
End Sub

Sub AutofitAllColumnsAllRows()
    If MsgBox("OK = Autofit all rows and columns", vbOKCancel, "") = vbCancel Then end
    Columns.EntireColumn.AutoFit
    Rows.EntireRow.AutoFit
End Sub

