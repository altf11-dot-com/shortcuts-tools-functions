Attribute VB_Name = "Module1"

Sub KS_CtrlShiftD()
    MoveCurrentRowDownOne
End Sub


Sub MoveCurrentRowDownOne()
    Rows(ActiveCell.row).Cut
    Rows(ActiveCell.row + 2).Insert Shift:=xlDown
    Cells(ActiveCell.row + 1, ActiveCell.Column).Activate
End Sub

Sub KS_CtrlShiftU()
    MoveCurrentRowUp
End Sub

Sub MoveCurrentRowUp()
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
    NextUpwardValue = Cells(TargetRow, currCol).value
    If NextUpwardValue = "" Then
            While TargetRow > MinRow And NextUpwardValue = ""
                TargetRow = TargetRow - 1
                NextUpwardValue = Cells(TargetRow, currCol).value
            Wend
            TargetRow = TargetRow + 1
    End If
    If MatchString <> "" Then
        If NextUpwardValue <> MatchString Then
            While TargetRow > MinRow And NextUpwardValue <> MatchString
                TargetRow = TargetRow - 1
                NextUpwardValue = Cells(TargetRow, currCol).value
            Wend
            TargetRow = TargetRow + 1
        End If
    End If
    Rows(currRow).Cut
    Rows(TargetRow).Insert Shift:=xlDown
    Cells(TargetRow, currCol).Activate
End Sub



