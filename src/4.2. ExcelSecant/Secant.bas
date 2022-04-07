Attribute VB_Name = "Secant"
Sub ExpandFormulaSecant()
Attribute ExpandFormulaSecant.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ExpandFormulaSecant Macro
'

'
    Dim x As Range
    Set x = Application.ActiveCell


    If InStr(1, (Range("D4").Formula), "x") > 0 Then
        
        Range("D7").Formula = ""
        
        Range("D4").Copy
        Range("D10").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        Range("D10").Replace What:="exp", Replacement:="z"
        Range("D10").Replace What:="x", Replacement:="C10"
        Range("D10").Replace What:="z", Replacement:="exp"
        
        
        Range("D10").AutoFill Destination:=Range("D10:D511")
        
        Range("D7").Formula = "'" & Range("D4").Formula
        
        'Range("D7").Replace What:="C4", Replacement:="x"
    
        With Range("D7")
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
        
    Else
    
        MsgBox "Error: Fill in Blanks with X-values."
        
    End If
    

        
    x.Select
    
End Sub
Sub ClearFormulaSecant()
Attribute ClearFormulaSecant.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearFormulaSecant Macro
'
    Dim x As Range
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    Set x = Application.ActiveCell
    
    'Asking for Confirmation
    MsgConfirm = MsgBox("All entries will be cleared!" & vbNewLine & "Continue?", vbOKCancel + vbDefaultButton2, "Clear Contents?")
    If MsgConfirm = vbCancel Then Exit Sub
    
    Range("D10:D511").ClearContents
    Range("D7,D4").ClearContents
    
    x.Select
End Sub
Sub HideValuesSecant()
Attribute HideValuesSecant.VB_ProcData.VB_Invoke_Func = " \n14"
'
' HideValuesSecant Macro
'

'
    Rows("9:511").EntireRow.Hidden = True
End Sub
Sub ShowValuesSecant()
Attribute ShowValuesSecant.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ShowValuesSecant Macro
'

'
    Rows("9:512").EntireRow.Hidden = False
End Sub
