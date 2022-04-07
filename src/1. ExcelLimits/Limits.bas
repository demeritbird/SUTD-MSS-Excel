Attribute VB_Name = "Limits"
Option Explicit

Sub Expand()

    Dim x As Range
    Set x = Application.ActiveCell
    
    If InStr(Range("I3").Formula, "x") > 0 Then

    Range("I3").Replace What:="exp", Replacement:="z"
    Range("I3").Replace What:="x", Replacement:="H3"
    Range("I3").Replace What:="z", Replacement:="exp"

    Range("I3:J3").Copy
    Range("I4:J13").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Range("J2").Formula = "'" & Range("I3").Formula
    Range("J2").Replace What:="H3", Replacement:="x"
    Range("I3").Replace What:="H3", Replacement:="x"
    
    Application.CutCopyMode = False
    
    Else
        MsgBox ("Please fill in x variable.")
        
    End If
    
        
    x.Select
    
End Sub


Sub Clear()

    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    
    'Asking for Confirmation
    MsgConfirm = MsgBox("All entries will be cleared!" & vbNewLine & "Continue?", vbOKCancel + vbDefaultButton2, "Clear Contents?")
    If MsgConfirm = vbCancel Then Exit Sub

    Range("J2,I3:J18").ClearContents


End Sub

