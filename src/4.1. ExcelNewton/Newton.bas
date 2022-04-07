Attribute VB_Name = "Newton"
Sub ExpandFormulaNewton()
Attribute ExpandFormulaNewton.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Dim x As Range
    Set x = Application.ActiveCell


    If InStr(1, (Range("D4").Formula), "x") > 0 And InStr(1, (Range("E4").Formula), "x") > 0 Then
            Range("D5").Formula = ""
            Range("E5").Formula = ""
    
            Range("D4:E4").Copy
            Range("D8:E8").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            
            Range("D8:E8").Replace What:="exp", Replacement:="z"
            Range("D8:E8").Replace What:="x", Replacement:="C8"
            Range("D8:E8").Replace What:="z", Replacement:="exp"
            
            
            Range("D8").AutoFill Destination:=Range("D8:D508")
            Range("E8").AutoFill Destination:=Range("E8:E508")
            
            Range("D5").Formula = "'" & Range("D4").Formula
            Range("E5").Formula = "'" & Range("E4").Formula
            
            'Range("D5:E5").Replace What:="C4", Replacement:="x", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            With Range("D5:E5")
                .Font.Bold = True
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
             End With
        
        
    Else:
        MsgBox "Error: Fill in Blanks with X-values."

    
    End If
    
    
    x.Select
End Sub
Sub ClearFormulaNewton()
'
' Macro6 Macro
'

    Dim x As Range
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    Set x = Application.ActiveCell
    
    'Asking for Confirmation
    MsgConfirm = MsgBox("All entries will be cleared!" & vbNewLine & "Continue?", vbOKCancel + vbDefaultButton2, "Clear Contents?")
    If MsgConfirm = vbCancel Then Exit Sub
    
    Range("D4:E5").ClearContents
    Range("D8:E508").ClearContents
    
    x.Select
End Sub
Sub HideValuesNewton()
'
' HideValues Macro
'

    Dim x As Range
    Set x = Application.ActiveCell

    Rows("7:508").EntireRow.Hidden = True
    
    x.Select
End Sub
Sub ExpandValuesNewton()
'
' ExpandValues Macro
'

    Dim x As Range
    Set x = Application.ActiveCell

    Rows("6:509").EntireRow.Hidden = False
    
    x.Select
End Sub
