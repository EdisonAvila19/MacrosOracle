Attribute VB_Name = "Módulo2"

Sub Main()
    
    'For Your Action => ForActionCo
    'Major Issues => MajorIssuesCo
    'Minor Issues => MinorIssuesCo
    'Excess Issue% => LastCo
        
    'Image Issue (Q1) => Q1Co
    'Content Issue (Q2) => Q2Co
    'Link Issue (Q3) => Q3Co
    'Template or Format Issue (Q4) => Q4Co
    'Tracking Issue (Q5) => Q5Co
    'Process Issue (Q6) => Q6Co
     
    
    'Last Row => EndRow

    SheetName = ActiveSheet.Name 'Nombre de la hoja de la base de datos
    
    Cells.Find(What:="For Your Action", MatchCase:=True).Activate  'Busca columna con titulo For Your Action
    ForActionCo = Selection.Column '# de columna "For Your Action"
    'MsgBox (ForActionCo)
    Cells.Find(What:="Major Issues", LookAt:=xlWhole).Activate
    MajorIssuesCo = Selection.Column '# de columna "Major Issues"
    'MsgBox (MajorIssuesCo)
    Cells.Find(What:="Minor Issues", LookAt:=xlWhole).Activate
    MinorIssuesCo = Selection.Column '# de columna "Minor Issues"
    'MsgBox (MinorIssuesCo)
    Cells.Find(What:="Excess Issue%", LookAt:=xlWhole).Activate
    LastCo = Selection.Column '# de columna "Last Column"
    'MsgBox (LastCo)
    Cells.Find(What:="Image Issue (Q1)", MatchCase:=True).Activate
    Q1Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q1Co)
    Cells.Find(What:="Content Issue (Q2)", MatchCase:=True).Activate
    Q2Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q2Co)
    Cells.Find(What:="Link Issue (Q3)", MatchCase:=True).Activate
    Q3Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q3Co)
    Cells.Find(What:="Template or Format Issue (Q4)", MatchCase:=True).Activate
    Q4Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q4Co)
    Cells.Find(What:="Tracking Issue (Q5)", MatchCase:=True).Activate
    Q5Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q5Co)
    Cells.Find(What:="Process Issue (Q6)", MatchCase:=True).Activate
    Q6Co = Selection.Column '# de columna "Minor Issues"
    'MsgBox (Q6Co)
      
    Selection.Cells(1, 1).Select 'Ir a la celda 1,1
    Selection.End(xlDown).Select 'Ir a la ultima fila de la primera columna
    EndRow = Selection.Row 'Guarda ultima fila de todos los registros
    
    'For Your Action
    Range(Cells(2, ForActionCo), Cells(EndRow, ForActionCo)).Select
    Selection.Copy
    Sheets.Add.Name = "Order"
    Range("I1").Select
    ActiveSheet.Paste
    
    'Texto en Columnas
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=Chr(10), FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
    'Conteo
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=+COUNTA(RC[1]:RC[25])"
    'Range("H1").Select
    Selection.Copy
    Range(Cells(2, 8), Cells(EndRow - 1, 8)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Copiar Q# a nueva hoja para contar total
    Sheets(SheetName).Select
    Range(Cells(2, Q1Co), Cells(EndRow, Q6Co)).Select
    Selection.Copy
    Sheets("Order").Select
    Range("B1").Select
    ActiveSheet.Paste
    Range(Cells(1, 8), Cells(EndRow - 1, 8)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Sumar cantidad de espacios
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+SUM(RC[1]:RC[7])"
    Range("A1").Select
    Selection.Copy
    Range(Cells(2, 1), Cells(EndRow - 1, 1)).Select
    ActiveSheet.Paste
    Range(Cells(1, 1), Cells(EndRow - 1, 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Max = 9 + WorksheetFunction.Max(Range(Cells(1, 8), Cells(EndRow - 1, 8)))
    'MsgBox (Max)
        
    'MajorIssuesCo
    Sheets(SheetName).Select
    Range(Cells(2, MajorIssuesCo - 1), Cells(EndRow, MajorIssuesCo)).Select
    Selection.Copy
    Sheets("Order").Select
    Cells(1, Max).Select
    ActiveSheet.Paste
        For i = 1 To EndRow - 1
           If Cells(i, Max) = 0 Then
               Cells(i, Max + 1) = ""
           End If
        Next i
    Range(Cells(1, Max + 1), Cells(EndRow - 1, Max + 1)).Select
    Selection.TextToColumns Destination:=Range(Cells(1, Max + 1), Cells(1, Max + 1)), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=Chr(10), FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
    
    Max2 = Max + 1 + WorksheetFunction.Max(Range(Cells(1, Max), Cells(EndRow - 1, Max)))
    'MsgBox (Max2)
    
    'MinorIssuesCo
    Sheets(SheetName).Select
    Range(Cells(2, MinorIssuesCo - 1), Cells(EndRow, MinorIssuesCo)).Select
    Selection.Copy
    Sheets("Order").Select
    Cells(1, Max2).Select
    ActiveSheet.Paste
        For i = 1 To EndRow - 1
           If Cells(i, Max2) = 0 Then
               Cells(i, Max2 + 1) = ""
           End If
        Next i
    Range(Cells(1, Max2 + 1), Cells(EndRow - 1, Max2 + 1)).Select
    
    Selection.TextToColumns Destination:=Range(Cells(1, Max2 + 1), Cells(1, Max2 + 1)), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=Chr(10), FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
    
    'Insert Rows
    EndInsRow = EndRow
    For i = -EndRow + 1 To -1
        
        Cells(-i, Max).Select
        If Cells(-i, Max) = 0 Then
            Cells(-i, Max + 1) = "No issues"
        End If
        Cells(-i, Max2).Select
        If Cells(-i, Max2) = 0 Then
            Cells(-i, Max2 + 1) = "No issues"
        End If
        
        Cells(-i, 1).Select
        If Cells(-i, 1) > 1 Then
            
            Total = Cells(-i, 1) - 1
            Range(Cells(-i + 1, 1), Cells(-i + Total, 1)).Select
            Selection.EntireRow.insert
            Sheets(SheetName).Select
            Range(Cells(-i + 2, 1), Cells(-i + 1 + Total, 1)).Select
            Selection.EntireRow.insert
            Sheets("Order").Select
            
            EndInsRow = EndInsRow + Total
            
        'Ordenar 1.High 2.Low 3.FYA 4.LastCo
        'Columnas Max(13) Max2(16) 8
        
            Cells(-i, Max).Select 'High
            Sequence = 0
            If Cells(-i, Max) = 1 Then
                Sequence = Cells(-i, Max)
                Cells(-i, Max + 1).Copy
                Cells(-i, 2).Select
                ActiveSheet.Paste
            ElseIf Cells(-i, Max) > 1 Then
                Var = Cells(-i, Max)
                Sequence = Var
                Range(Cells(-i, Max + 2), Cells(-i, Max + Var)).Select
                Selection.Copy
                Cells(-i + 1, Max + 1).Select
                    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                Range(Cells(-i, Max + 1), Cells(-i, Max + Var)).Copy
                Cells(-i, 2).Select
                    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                Range(Cells(-i, Max + 2), Cells(-i, Max + Var)).Clear
            End If
            
            
                'MsgBox (Sequence)
                
            Cells(-i, Max2).Select 'Low
            If Cells(-i, Max2) > 0 Then
                Var = Cells(-i, Max2)
                If Sequence = 0 And Var > 1 Then
                    Range(Cells(-i, Max2 + 2), Cells(-i, Max2 + Var)).Select
                    Selection.Copy
                    Cells(-i + 1, Max2 + 1).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, Max2 + 1), Cells(-i, Max2 + Var)).Copy
                    Cells(-i, 2).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, Max2 + 2), Cells(-i, Max2 + Var)).Clear
                ElseIf Sequence = 0 And Var = 1 Then
                    Cells(-i, Max2 + 1).Copy
                    Cells(-i, 2).Select
                    ActiveSheet.Paste
                Else
                    Range(Cells(-i, Max2 + 1), Cells(-i, Max2 + Var)).Select
                    Selection.Copy
                    Cells(-i + Sequence, Max2 + 1).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, Max2 + 1), Cells(-i, Max2 + Var)).Copy
                    Cells(-i + Sequence, 2).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, Max2 + 1), Cells(-i, Max2 + Var)).Clear
                End If
                Sequence = Sequence + Var
            End If
            
            Cells(-i, 8).Select 'FYA
            If Cells(-i, 8) > 0 Then
                Var = Cells(-i, 8)
                If Sequence = 0 And Var > 1 Then
                    Range(Cells(-i, 10), Cells(-i, 8 + Var)).Select
                    Selection.Copy
                    Cells(-i + 1, 9).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, 9), Cells(-i, 8 + Var)).Copy
                    Cells(-i, 2).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, 10), Cells(-i, 8 + Var)).Clear
                ElseIf Sequence = 0 And Var = 1 Then
                    Cells(-i, 9).Copy
                    Cells(-i, 2).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                Else
                    Range(Cells(-i, 9), Cells(-i, 8 + Var)).Select
                    Selection.Copy
                    Cells(-i + Sequence, 9).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, 9), Cells(-i, 8 + Var)).Copy
                    Cells(-i + Sequence, 2).Select
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                            False, Transpose:=True
                    Range(Cells(-i, 9), Cells(-i, 8 + Var)).Clear
                End If
            End If
        End If
        If Cells(-i, 1) = 1 Then
            If Cells(-i, Max) = 1 Then
                Cells(-i, Max + 1).Copy
                Cells(-i, 2).Select
                ActiveSheet.Paste
            End If
            If Cells(-i, Max2) = 1 Then
                Cells(-i, Max2 + 1).Copy
                Cells(-i, 2).Select
                ActiveSheet.Paste
            End If
            If Cells(-i, 8) = 1 Then
                Cells(-i, 9).Copy
                Cells(-i, 2).Select
                ActiveSheet.Paste
            End If
        End If
    'MsgBox (EndInsRow)
    Next i
    
    Range(Cells(1, Max + 1), Cells(EndInsRow - 1, Max + 1)).Select
    Selection.Copy
    Sheets(SheetName).Select
    Cells(2, MajorIssuesCo).Select
    ActiveSheet.Paste
    
    Sheets("Order").Select
    
    Range(Cells(1, Max2 + 1), Cells(EndInsRow - 1, Max2 + 1)).Select
    Selection.Copy
    Sheets(SheetName).Select
    Cells(2, MinorIssuesCo).Select
    ActiveSheet.Paste
    
    Sheets("Order").Select
    
    Range(Cells(1, 9), Cells(EndInsRow - 1, 9)).Select
    Selection.Copy
    Sheets(SheetName).Select
    Cells(2, ForActionCo).Select
    ActiveSheet.Paste
    
    Sheets("Order").Select
    
    Range(Cells(1, 2), Cells(EndInsRow - 1, 2)).Select
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.FormulaR1C1 = "No issues"
    Range(Cells(1, 2), Cells(EndInsRow - 1, 2)).Copy
    Sheets(SheetName).Select
    Cells(2, LastCo + 1).Select
    ActiveSheet.Paste
    Cells(1, LastCo + 1) = "Issues"

    Sheets(SheetName).Select
    
    Range("A2", Cells(EndInsRow, MajorIssuesCo - 2)).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=+R[-1]C"
    ActiveWindow.SmallScroll Down:=-30
    Range("A2", Cells(EndInsRow, MajorIssuesCo - 2)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    
    Sheets("Order").Delete
    
End Sub
