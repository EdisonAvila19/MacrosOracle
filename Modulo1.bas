Attribute VB_Name = "Modulo1"
Sub Inicio()
Attribute Inicio.VB_ProcData.VB_Invoke_Func = "q\n14"
' Inicio
' Elimina espacios vacios en el rango seleccionado; Establece Variables iniciales; Llama funci�n Proceso
' Ctrl + q

    'Selection.End(xlDown).Select
    'Range("L2:L1240").Select
    
    Ini = Selection.Row
    
    'Agrega "-" a todas las celdas vacias
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "-"

    'Delimita Inicio y Final de datos
    Fi = Selection.End(xlDown).Row
    Co = Selection.End(xlDown).Column
    Ini = Selection.Row
    
    
    'Range("M1") = Fi
    'Range("N1") = Ini
    'Range("O1") = Co
    
    Cells(1, Co + 1) = Cells(1, Co)
    
    'Llama Funcion Proceso
    Proceso Co, Fi, Ini
End Sub


Sub Proceso(Co, Fi, Ini)
Attribute Proceso.VB_ProcData.VB_Invoke_Func = "q\n14"
' Proceso
' Separa informaci�n de cada selda por salto de linea, Agrega Filas necesarias y transpone de columnas a filas la informaci�n

    ColMax = Co + 28 'Columna U
    
    'Selecciona los datos a tratar
    Range(Cells(Ini, Co), Cells(Fi, Co)).Select
           
    
    'Separa texto inicial en columnas 'N2
    Selection.TextToColumns Destination:=Range(Cells(Ini, Co + 2), Cells(Ini, Co + 2)), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=Chr(10), FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
    
    'Conteo de celdas llenas
    Range(Cells(Ini, Co + 2), Cells(Ini, ColMax)).Select 'M2:V2
    Cells(Ini, ColMax).Activate 'V2
    ActiveCell.FormulaR1C1 = "=+COUNTA(RC[-" & ColMax - Co - 2 & "]:RC[-1])-1"
    'ActiveCell.FormulaR1C1 = "=+COUNTA(RC[-28]:RC[-1])-1"
    Cells(Ini, ColMax).Select 'V2
    Selection.Copy
    Range(Cells(Ini, ColMax), Cells(Fi, ColMax)).Select 'V2:V100
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Copia Columna N2 a M2
    Range(Cells(Ini, Co + 2), Cells(Fi, Co + 2)).Select
    Selection.Copy
    Range(Cells(Ini, Co + 1), Cells(Ini, Co + 1)).Select
    ActiveSheet.Paste

    'Ciclo de creaci�n de filas
    For i = -Fi To -Ini
        
        If Cells(-i, ColMax) > 0 Then
            Cont = Cells(-i, ColMax)
            
            Range(Cells(-i + 1, ColMax), Cells(-i + Cont, ColMax)).Select
            Selection.EntireRow.insert
            
            'Transponer datos
            
            Range(Cells(-i, Co + 2), Cells(-i, Co + 2 + Cont)).Select
            Selection.Copy
            Range(Cells(-i, Co + 1), Cells(-i, Co + 1)).Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
            
            'Copiar Datos antes de seleccion
            
            Range(Cells(-i, 1), Cells(-i, Co - 1)).Select
            Selection.Copy
            Range(Cells(-i + 1, 1), Cells(-i + Cont, Co - 1)).Select
            ActiveSheet.Paste
            
        End If
        
    Next i
    
    'Borrar Datos Extras
    
    Range(Cells(1, Co + 2), Cells(5000, ColMax)).Clear
    
End Sub





