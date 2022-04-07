Attribute VB_Name = "Euler"
Option Explicit
    


Sub ExpandFunction()
'
'

    Dim x As Range
    Set x = Application.ActiveCell


    If InStr(Range("I4").Formula, "h") = 0 Then
    
        MsgBox ("Please fill in desired values.")
    


    ElseIf InStr(Range("I4").Formula, "x") > 0 Then
    
    'With Xn Expansion
            
            
        Range("I5").Formula = "'" & Range("I4").Formula
        
        Range("I4").Replace What:="h", Replacement:="$B$4"
        
        Range("I4").Replace What:="exp", Replacement:="z"
        Range("I4").Replace What:="x", Replacement:="E4"
        Range("I4").Replace What:="z", Replacement:="exp"
        
        Range("I4").Replace What:="y", Replacement:="F4"
        
        'Transferring Data
            
        Range("I4").Copy
        Range("F9").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("F9").Replace What:="B9", Replacement:="E8"
        Range("F9").Replace What:="C9", Replacement:="F8"

            
            
        Range("F9").AutoFill Destination:=Range("F9:F1008")
        Application.CutCopyMode = False

        'Replace Back

        Range("I4").Replace What:="$B$4", Replacement:="h"
        Range("I4").Replace What:="E4", Replacement:="x"
        Range("I4").Replace What:="F4", Replacement:="y"

    Else
        
    'Without Xn Expansion
            
            
        Range("I5").Formula = "'" & Range("I4").Formula
        
        Range("I4").Replace What:="h", Replacement:="$B$4"
        Range("I4").Replace What:="y", Replacement:="F4"
        
        'Transferring Data
        
        Range("I4").Copy
        Range("F9").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("F9").Replace What:="C9", Replacement:="F8"
        
        

        
        Range("F9").AutoFill Destination:=Range("F9:F1008")
        Application.CutCopyMode = False
        
        'Replace Back
        
        Range("I4").Replace What:="$B$4", Replacement:="h"
        Range("I4").Replace What:="F4", Replacement:="y"
            

        
    End If
    
    x.Select
    
End Sub


Sub ClearFunction()
    
    Dim x As Range
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    Set x = Application.ActiveCell
    
    'Asking for Confirmation
    MsgConfirm = MsgBox("All entries will be cleared!" & vbNewLine & "Continue?", vbOKCancel + vbDefaultButton2, "Clear Contents?")
    If MsgConfirm = vbCancel Then Exit Sub
    
    Range("I4:I5").ClearContents
    Range("E4:F4").ClearContents
    Range("F9:F1008").ClearContents
    
    x.Select
    
End Sub
Sub HideValues()
'
' HideValues Macro
'

'
    Rows("7:1009").EntireRow.Hidden = True
End Sub
Sub ShowValues()
'
' ShowValues Macro
'

'
    Rows("7:1009").EntireRow.Hidden = False
End Sub
