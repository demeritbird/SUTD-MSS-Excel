Attribute VB_Name = "Sequences"
Option Explicit


Sub expand()

    Dim x As Range
    Set x = Application.ActiveCell


        If InStr(Range("F6").Formula, "x") > 0 Then

            Range("F7").Formula = Range("F6").Formula
            Range("F7").Formula = "'" & Range("F7").Formula
            
            Range("F10").Formula = Range("F6").Formula
            
            Range("F10").Replace What:="exp", Replacement:="z"
            Range("F10").Replace What:="x", Replacement:="$E$3"
            Range("F10").Replace What:="z", Replacement:="exp"
            
            
            Range("F10").Replace What:="k", Replacement:="E10"
            Range("F10").AutoFill Destination:=Range("F10:F1009")
            
                    
        ElseIf InStr(Range("F6").Formula, "k") > 0 Then
        'Executes if there is no x-value in the formula entry.
    
            Range("F7").Formula = Range("F6").Formula
            Range("F7").Formula = "'" & Range("F7").Formula
        
            Range("F10").Formula = Range("F6").Formula
            Range("F10").Replace What:="k", Replacement:="E10"
            Range("F10").AutoFill Destination:=Range("F10:F1009")
            
        
        Else
        'Executes if there is no k-value in the formula entry.
    

            MsgBox ("Please fill in desired values")
               

    
    End If
       
    
    x.Select
    
    

End Sub



Sub clear()

    Dim x As Range
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    Set x = Application.ActiveCell
    
    'Asking for Confirmation
    MsgConfirm = MsgBox("All entries will be cleared!" & vbNewLine & "Continue?", vbOKCancel + vbDefaultButton2, "Clear Contents?")
    If MsgConfirm = vbCancel Then Exit Sub
    
    Range("F6:F7", "F10:F1009").ClearContents
    
    x.Select



End Sub

Sub show()

    Range("F9:F1009").EntireRow.Hidden = False
    
End Sub

Sub hide()

    Range("F9:F1009").EntireRow.Hidden = True
    
End Sub
