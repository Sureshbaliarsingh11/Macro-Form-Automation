# Macro-Form-Automation
Order Management Form Automation 
Sub CheckRequiredCells()
    Dim ws As Worksheet
    Dim addresses As Variant
    Dim i As Long
    Dim r As Range
    
    Set ws = ActiveSheet
    
    addresses = Array("F4", "C7", "F7", "F13", "H7", "F31", "C32", "C35", _
                "C38", "C39", "C40", "E40", "G40", "J32", "M31", "N40", "L40", "J38", _
                 "J39", "J40", "F56", "I56", "M13", _
                 "C16", "J16", _
                 "J35", "C49", "C51", "C53", "F51", "I51", "L51", "I59", "F59", "C59", "C72", _
                 "F72", "F74", "C74", "J72", "J74", "M72", "M74", "C84", "F84", "H84")
    
    ' Optional: clear previous highlighting
    Dim addr As Variant
    For Each addr In addresses
        ws.Range(addr).Interior.ColorIndex = xlNone
    Next addr

    For i = LBound(addresses) To UBound(addresses)
        Set r = ws.Range(addresses(i))
        If IsBlankLike(r) Then
            r.Interior.Color = vbRed
            Application.Goto r, True   ' Visually go to the cell
            MsgBox "Cell " & r.Address(0, 0) & " is empty.", vbExclamation, "Empty Check"
            Exit Sub                    ' Stop at first blank
        End If
    Next i
    
    MsgBox "All required cells have values.", vbInformation, "Check Complete"
    ' Next steps…
    SubmitMacro
End Sub

Private Function IsBlankLike(ByVal r As Range) As Boolean
    ' Treat empty or formula returning "" as blank. Also trims spaces.
    IsBlankLike = (Len(Trim$(CStr(r.Value))) = 0)
End Function


Sub SubmitMacro()
'
' SubmitMacro Macro
'

'
    Range("A4").Select
    Selection.End(xlDown).Select
    Range("C1048574").Select
    Selection.End(xlUp).Select
    Range("C1048443").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Data").Select
    Range("A2").Select
    Selection.Insert Shift:=xlDown
    Sheets("New Order-LTL Form").Select
    Range("C1048574").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("C13").Select
    Selection.ClearContents
    Range("C14:F14").Select
    Selection.ClearContents
    Range("C15").Select
    Selection.ClearContents
    Range("C16").Select
    Selection.ClearContents
    Range("C17").Select
    Selection.ClearContents
    Range("C18").Select
    Selection.ClearContents
    Range("C19").Select
    Selection.ClearContents
    Range("E18").Select
    Selection.ClearContents
    Range("E17").Select
    Selection.ClearContents
    Range("E16").Select
    Selection.ClearContents
    Range("F13").Select
    Selection.ClearContents
    Range("G16").Select
    Selection.ClearContents
    Range("J13").Select
    Selection.ClearContents
    Range("J14:M14").Select
    Selection.ClearContents
    Range("J15").Select
    Selection.ClearContents
    Range("J16").Select
    Selection.ClearContents
    Range("J17").Select
    Selection.ClearContents
    Range("J18").Select
    Selection.ClearContents
    Range("J19").Select
    Selection.ClearContents
    Range("L18").Select
    Selection.ClearContents
    Range("L17").Select
    Selection.ClearContents
    Range("L16").Select
    Selection.ClearContents
    Range("M13").Select
    Selection.ClearContents
    Range("N16").Select
    Selection.ClearContents
    Range("C7").Select
    ActiveWindow.SmallScroll Down:=8
    Range("C22").Select
    Selection.ClearContents
    Range("C23:F23").Select
    Selection.ClearContents
    Range("C24").Select
    Selection.ClearContents
    Range("C74").Select
    Selection.ClearContents
    Range("F74").Select
    Selection.ClearContents
    Range("F72").Select
    Selection.ClearContents
    Range("J74").Select
    Selection.ClearContents
    Range("M74").Select
    Selection.ClearContents
    Range("M72").Select
    Selection.ClearContents
    Range("C57").Select
    Selection.ClearContents
    Range("C59").Select
    Selection.ClearContents
    Range("F59").Select
    Selection.ClearContents
    Range("C40").Select
    Selection.ClearContents
    Range("E40").Select
    Selection.ClearContents
    Range("G40").Select
    Selection.ClearContents
    Range("J40").Select
    Selection.ClearContents
    Range("L40").Select
    Selection.ClearContents
    Range("N40").Select
    Selection.ClearContents
    Range("J36").Select
    Selection.ClearContents
    Range("C7").Select
    Selection.ClearContents
    Range("C25").Select
    Selection.ClearContents
    Range("C26").Select
    Selection.ClearContents
    Range("C27").Select
    Selection.ClearContents
    Range("C28").Select
    Selection.ClearContents
    Range("E27").Select
    Selection.ClearContents
    Range("E26").Select
    Selection.ClearContents
    Range("E25").Select
    Selection.ClearContents
    Range("F22").Select
    Selection.ClearContents
    Range("G25").Select
    Selection.ClearContents
    Range("C32").Select
    ActiveWindow.SmallScroll Down:=11
    Range("C31").Select
    Selection.ClearContents
    Range("C32").Select
    Selection.ClearContents
    Range("C33:F33").Select
    Selection.ClearContents
    Range("C34").Select
    Selection.ClearContents
    Range("C35").Select
    Selection.ClearContents
    Range("C36").Select
    Selection.ClearContents
    Range("C37").Select
    Selection.ClearContents
    Range("C41").Select
    Selection.ClearContents
    Range("C42").Select
    Selection.ClearContents
    Range("C38").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("C38").Select
    Selection.Copy
    Range("C39").Select
    ActiveSheet.Paste
    Range("E35").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("E36").Select
    Selection.ClearContents
    Range("E37").Select
    Selection.ClearContents
    Range("E38").Select
    Selection.ClearContents
    Range("E39").Select
    Selection.ClearContents
    Range("G31").Select
    Selection.ClearContents
    Range("F31").Select
    Selection.ClearContents
    Range("G35").Select
    Selection.ClearContents
    Range("J31").Select
    Selection.ClearContents
    Range("J32").Select
    Selection.ClearContents
    Range("J33:M33").Select
    Selection.ClearContents
    Range("J34").Select
    Selection.ClearContents
    Range("J35").Select
    Selection.ClearContents
    Range("J37").Select
    Selection.ClearContents
    Range("J41").Select
    Selection.ClearContents
    Range("J42").Select
    Selection.ClearContents
    Range("J38").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("J38").Select
    Selection.Copy
    Range("J39").Select
    ActiveSheet.Paste
    Range("L35").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("L36").Select
    Selection.ClearContents
    Range("L37").Select
    Selection.ClearContents
    Range("L38").Select
    Selection.ClearContents
    Range("L39").Select
    Selection.ClearContents
    Range("N35").Select
    Selection.ClearContents
    Range("N40").Select
    ActiveSheet.Shapes.Range(Array("Rectangle 6")).Select
    Range("L44").Select
    Selection.End(xlToLeft).Select
    Range("B45").Select
    ActiveWindow.SmallScroll Down:=11
    Range("C51").Select
    Selection.ClearContents
    Range("C53").Select
    Selection.ClearContents
    Range("C55").Select
    Selection.ClearContents
    Range("E53").Select
    Selection.ClearContents
    Range("F51").Select
    Selection.ClearContents
    Range("D55").Select
    ActiveCell.FormulaR1C1 = "L"
    Range("E55").Select
    ActiveCell.FormulaR1C1 = "W"
    Range("F55").Select
    ActiveCell.FormulaR1C1 = "H"
    Range("G55").Select
    ActiveWindow.SmallScroll Down:=3
    Range("F57").Select
    Selection.ClearContents
    Range("C59").Select
    ActiveWindow.SmallScroll Down:=6
    Range("I57").Select
    Selection.ClearContents
    Range("I59").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-6
    Range("I51").Select
    ActiveCell.FormulaR1C1 = "FAK"
    Range("I52").Select
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll Down:=6
    Range("M52").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("L59").Select
    Selection.ClearContents
    Range("M59").Select
    Selection.ClearContents
    Range("L51").Select
    ActiveCell.FormulaR1C1 = "FRIEGHT OF ALL KIND*"
    Range("L52").Select
    ActiveWindow.SmallScroll Down:=9
    Range("K61").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("L63").Select
    Selection.ClearContents
    Range("L65").Select
    Selection.ClearContents
    Range("I63").Select
    Selection.ClearContents
    Range("I65").Select
    Selection.ClearContents
    ActiveWindow.LargeScroll ToRight:=-1
    Range("F63").Select
    Selection.ClearContents
    Range("F65").Select
    Selection.ClearContents
    Range("C65").Select
    Selection.ClearContents
    Range("C63").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=12
    Range("C72").Select
    Selection.ClearContents
    Range("C76").Select
    Selection.ClearContents
    Range("J72").Select
    Selection.ClearContents
    Range("J76").Select
    Selection.ClearContents
    Range("B80").Select
    ActiveWindow.SmallScroll Down:=8
    Range("C80").Select
    Selection.ClearContents
    Range("C82").Select
    Selection.ClearContents
    Range("C84").Select
    Selection.ClearContents
    Range("F84").Select
    Selection.ClearContents
    Range("H84").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=11
    Range("C92").Select
    Selection.ClearContents
    Range("C96").Select
    Selection.ClearContents
    Range("C97").Select
    Selection.ClearContents
    Range("C98").Select
    Selection.ClearContents
    Range("C99").Select
    Selection.ClearContents
    Range("C57").Select
    ActiveCell.FormulaR1C1 = "Yes"
    Range("F98").Select
    Selection.ClearContents
    Range("F96").Select
    Selection.ClearContents
    Range("G96").Select
    ActiveWindow.SmallScroll Down:=-79
    Range("D1").Select
End Sub

Sub SaveMacro()
'
' SaveMacro Macro
'

'
    Range("A7").Select
    Selection.End(xlDown).Select
    Range("C1048574").Select
    Selection.End(xlUp).Select
    Range("C1048443").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet3").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Save").Select
    Range("A2").Select
    Selection.Insert Shift:=xlDown
    Sheets("New Order-LTL Form").Select
    Range("D1048567").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-22
    Range("D7").Select
End Sub
Sub ResetMacro()
'
' ResetMacro Macro
'

'
    Range("C4").Select
    Selection.ClearContents
    Range("C13").Select
    Selection.ClearContents
    Range("C14:F14").Select
    Selection.ClearContents
    Range("C15").Select
    Selection.ClearContents
    Range("C16").Select
    Selection.ClearContents
    Range("C17").Select
    Selection.ClearContents
    Range("C18").Select
    Selection.ClearContents
    Range("C19").Select
    Selection.ClearContents
    Range("E19").Select
    Selection.ClearContents
    Range("E18").Select
    Selection.ClearContents
    Range("E17").Select
    Selection.ClearContents
    Range("E16").Select
    Selection.ClearContents
    Range("F13").Select
    Selection.ClearContents
    Range("G16").Select
    Selection.ClearContents
    Range("J13").Select
    Selection.ClearContents
    Range("J14").Select
    Selection.ClearContents
    Range("J15").Select
    Selection.ClearContents
    Range("J16").Select
    Selection.ClearContents
    Range("J17").Select
    Selection.ClearContents
    Range("J18").Select
    Selection.ClearContents
    Range("J19").Select
    Selection.ClearContents
    Range("L18").Select
    Selection.ClearContents
    Range("C24").Select
    Selection.ClearContents
    Range("C74").Select
    Selection.ClearContents
    Range("F74").Select
    Selection.ClearContents
    Range("F72").Select
    Selection.ClearContents
    Range("J74").Select
    Selection.ClearContents
    Range("M74").Select
    Selection.ClearContents
    Range("M72").Select
    Selection.ClearContents
    Range("C57").Select
    Selection.ClearContents
    Range("C59").Select
    Selection.ClearContents
    Range("F59").Select
    Selection.ClearContents
    Range("C40").Select
    Selection.ClearContents
    Range("E40").Select
    Selection.ClearContents
    Range("G40").Select
    Selection.ClearContents
    Range("J40").Select
    Selection.ClearContents
    Range("L40").Select
    Selection.ClearContents
    Range("N40").Select
    Selection.ClearContents
    Range("J36").Select
    Selection.ClearContents
    Range("C7").Select
    Selection.ClearContents
    Range("L17").Select
    Selection.ClearContents
    Range("L16").Select
    Selection.ClearContents
    Range("M13").Select
    Selection.ClearContents
    Range("N16").Select
    Selection.ClearContents
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    ActiveWindow.SmallScroll Down:=14
    Range("C22").Select
    Selection.ClearContents
    Range("C23:F23").Select
    Selection.ClearContents
    Range("C24").Select
    Selection.ClearContents
    Range("C25").Select
    Selection.ClearContents
    Range("C26").Select
    Selection.ClearContents
    Range("C27").Select
    Selection.ClearContents
    Range("C28").Select
    Selection.ClearContents
    Range("E27").Select
    Selection.ClearContents
    Range("E26").Select
    Selection.ClearContents
    Range("E25").Select
    Selection.ClearContents
    Range("F22").Select
    Selection.ClearContents
    Range("G25").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=9
    Range("C31").Select
    Selection.ClearContents
    Range("C32").Select
    Selection.ClearContents
    Range("C33:F33").Select
    Selection.ClearContents
    Range("C34").Select
    Selection.ClearContents
    Range("C35").Select
    Selection.ClearContents
    Range("C36").Select
    Selection.ClearContents
    Range("C37").Select
    Selection.ClearContents
    Range("C41").Select
    Selection.ClearContents
    Range("C42").Select
    Selection.ClearContents
    Range("E39").Select
    Selection.ClearContents
    Range("E38").Select
    Selection.ClearContents
    Range("E37").Select
    Selection.ClearContents
    Range("E36").Select
    Selection.ClearContents
    Range("E35").Select
    Selection.ClearContents
    Range("G35").Select
    Selection.ClearContents
    Range("J31").Select
    Selection.ClearContents
    Selection.ClearContents
    Range("J32").Select
    Selection.ClearContents
    Range("J33:M33").Select
    Selection.ClearContents
    Range("J34").Select
    Selection.ClearContents
    Range("J35").Select
    Selection.ClearContents
    Range("J36").Select
    Selection.ClearContents
    Range("J37").Select
    Selection.ClearContents
    Range("J41").Select
    Selection.ClearContents
    Range("J42").Select
    Selection.ClearContents
    Range("L39").Select
    Selection.ClearContents
    Range("L38").Select
    Selection.ClearContents
    Range("L37").Select
    Selection.ClearContents
    Range("L36").Select
    Selection.ClearContents
    Range("L35").Select
    Selection.ClearContents
    Range("M31").Select
    Selection.ClearContents
    Range("N35").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=14
    Range("C51").Select
    Selection.ClearContents
    Range("C53").Select
    Selection.ClearContents
    Range("C55").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=3
    Range("E53").Select
    Selection.ClearContents
    Range("F51").Select
    Selection.ClearContents
    Range("I57").Select
    Selection.ClearContents
    Range("I59").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=4
    Range("M59").Select
    Selection.ClearContents
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll Down:=8
    Range("C63").Select
    Selection.ClearContents
    Range("C65").Select
    Selection.ClearContents
    Range("F65").Select
    Selection.ClearContents
    Range("F63").Select
    Selection.ClearContents
    Range("I65").Select
    Selection.ClearContents
    Range("I63").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=2
    Range("L65").Select
    Selection.ClearContents
    Range("L63").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=2
    Range("M59").Select
    Selection.ClearContents
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll Down:=12
    Range("C72").Select
    Selection.ClearContents
    Range("C76").Select
    Selection.ClearContents
    Range("C74").Select
    ActiveWindow.SmallScroll Down:=9
    Range("C80").Select
    Selection.ClearContents
    Range("C82").Select
    Selection.ClearContents
    Range("C84").Select
    Selection.ClearContents
    Range("F84").Select
    Selection.ClearContents
    Range("C57").Select
    ActiveCell.FormulaR1C1 = "Yes"
    Range("H84").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=13
    Range("C92").Select
    Selection.ClearContents
    Range("C96").Select
    Selection.ClearContents
    Range("C97").Select
    Selection.ClearContents
    Range("C98").Select
    Selection.ClearContents
    Range("C99").Select
    Selection.ClearContents
    Range("F98").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-5
    Range("F96").Select
    ActiveWindow.SmallScroll Down:=8
    Range("F96").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-45
    Range("D55").Select
    ActiveCell.FormulaR1C1 = "L"
    Range("E55").Select
    ActiveCell.FormulaR1C1 = "W"
    Range("F55").Select
    ActiveCell.FormulaR1C1 = "H"
    Range("I51").Select
    ActiveCell.FormulaR1C1 = "FAK"
    Range("L57").Select
    ActiveWindow.SmallScroll Down:=-40
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll Down:=0
    ActiveWindow.SmallScroll ToRight:=5
    ActiveWindow.SmallScroll Down:=74
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll Down:=-26
    Range("D57").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("D4").Select
    ActiveWindow.SmallScroll Down:=-7
End Sub

