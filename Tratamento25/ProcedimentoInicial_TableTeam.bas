Attribute VB_Name = "Módulo1"
Public a() As Variant, b() As Variant, c() As Variant
Public cels As Range
Public i As Double
Public rowmin As Double
Public rowmax As Double
Public colmin As Integer
Public colmax As Integer
Sub Procedimento_Geral()
    
    Application.ScreenUpdating = False
    Dim wk As Worksheet, z As Double
    Dim contador As Double
    Set wb = ActiveWorkbook
    'Paste special
    Cells.Copy
    Range("A1").PasteSpecial xlPasteValues
    
    'rename the sheet to '2.5'
    ActiveSheet.Name = "2.5"
    
    'Delete all worksheets that are not the "2.5"
    Application.DisplayAlerts = False
    For Each wk In wb.Worksheets
        If wk.Name <> "2.5" Then
            wk.Delete
        End If
    Next wk
    Application.DisplayAlerts = True
    
'Search and replace all "-" for ""
    Cells.Select
    Cells.Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, _
        FormulaVersion:=xlReplaceFormula2 ', after:=Range("A1")
        
    
'Show up hidden row/colunms
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Selection.UnMerge
    Cells.EntireColumn.AutoFit
        
    Application.ScreenUpdating = True
    
    'Exclui colunas que não tem necessidade
    Set cels = Application.InputBox(prompt:="A matriz começa em qual célula?", Title:="Começo da matriz", Type:=8)
    
    Application.ScreenUpdating = False
    
    rowmin = cels.row
    colmin = cels.Column
    rowmax = WorksheetFunction.CountA(ActiveSheet.Columns(colmin))
    colmax = WorksheetFunction.CountA(ActiveSheet.Rows(rowmin))
    On Error GoTo errohandle
    i = colmin
    
    Do While ActiveSheet.Cells(rowmin, i).Value <> ""
            If ((ActiveSheet.Cells(rowmin, i).End(xlDown).Value = "") Or ((WorksheetFunction.Sum(ActiveSheet.Columns(i)) = 0) And IsNumeric(ActiveSheet.Cells(rowmin, i).End(xlDown).Value))) Then
                If ((ActiveSheet.Cells(rowmin, i) <> "VALOR EXCEDENTE") And (ActiveSheet.Cells(rowmin, i) <> "PRAZO(DIAS ÚTEIS)") And (ActiveSheet.Cells(rowmin, i) <> "CEPI") And (ActiveSheet.Cells(rowmin, i) <> "CEPF")) And Not (IsNumeric(ActiveSheet.Cells(rowmin, i))) Then
                    ActiveSheet.Columns(i).Delete
                    i = i - 1
                End If
            End If
        i = i + 1
    Loop
    
    'Mostra as colunas com categorias erradas
      Call declaração_variavel
      
      colmax = WorksheetFunction.CountA(ActiveSheet.Rows(rowmin))
      i = colmin
      Do While ActiveSheet.Cells(rowmin, i).Value <> ""
        For z = 0 To UBound(a)
            If ((ActiveSheet.Cells(rowmin, i).Value = a(z)) Or (ActiveSheet.Cells(rowmin, i).Value = b(z) Or (IsNumeric(ActiveSheet.Cells(rowmin, i).Value)))) Then
                Exit For
            ElseIf z = 67 Then
                ActiveSheet.Cells(rowmin, i).Interior.Color = vbRed
            End If
        Next z
        i = i + 1
      Loop
      
    Application.ScreenUpdating = True
    'Add the Info sheet
    'Sheets.Add
    'ActiveSheet.Name = "Info"
    'With Range("A1")
        '.Interior.Color = 4697456
        '.Font.Color = 167772155
        '.Value = "CHAMADO"
   ' End With
   ' Range("A2") = InputBox("Digite o número do chamado.", "Número do chamado")
    
   ' With Range("A4")
        '.Interior.Color = 4697456
       ' .Font.Color = 167772155
        '.Value = "ID CLIENTE"
   ' End With
   ' Range("A5") = InputBox("Digite o número do ID do cliente.", "ID do cliente")
    
    'With Range("A7")
        '.Interior.Color = 4697456
        '.Font.Color = 167772155
        '.Value = "MÉTODO"
    'End With
    'Range("A8") = InputBox("Digite o ID do método.", "ID do método")
    
    'With Range("A10")
        '.Interior.Color = 4697456
        '.Font.Color = 167772155
        '.Value = "CD"
   ' End With
    
    'Range("A11") = InputBox("Digite o CD (se for CD de origem nacional, deixe o '1')", "CD", "1")
    'Cells.Columns.AutoFit
Exit Sub

errohandle:

MsgBox ("Há um problema na categoria " & ActiveSheet.Cells(rowmin, i) & "!")
End
End Sub


    
Sub Tirar_Porcentagem()
Dim cel As Range
Set cel = Application.InputBox(prompt:="A partir de qual célula você quer transformar pra decimal?", Title:="Tirar Porcentagem", Type:=8)
cel.Activate


'insert a new column to the right
ActiveCell.Offset(0, 1).Activate
ActiveCell.EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'Add the formula to turn the percentage into a decimal
cel.Offset(0, 1).Activate
ActiveCell.FormulaR1C1 = "=RC[-1]*100"

'Find out what the final row is and after that paste the formulas in the cells down the first
cel.Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Activate
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

'Save to make sure that the formulas will be correctly applied
ActiveWorkbook.Save

'Copy, paste transform the type of data, and delete the column
    Selection.Copy
    cel.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "General"
    cel.Offset(0, 1).Activate
    ActiveCell.EntireColumn.Activate
    Selection.Delete Shift:=xlToLeft

End Sub

Sub declaração_variavel()

a = Array("CEPI", "CEPF", "PRAZO(DIAS ÚTEIS)", "FRETE TOTAL MÍNIMO", "VALOR EXCEDENTE", _
"FRETE VALOR SOBRE A NOTA(%)", "FRETE MÍNIMO", "% SOBRE A NF", "VALOR POR KG", _
"GRIS MÍNIMO", "GRIS MÁXIMO", "GRIS(%)", "FAIXA INICIAL DE GRIS", "FAIXA FINAL DE GRIS", _
"FAIXA VIGENTE SOBRE(NF ou Peso)", "GRIS(%)", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "TAS VALOR FIXO", _
"FAIXA INICIAL DE TAS", "FAIXA FINAL DE TAS", "FAIXA VIGENTE SOBRE(NF ou Peso)", _
"TAS VALOR FIXO", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "TRT MÍNIMO", "TRT MÁXIMO", "TRT(%)", _
"TRT VALOR FIXO", "TDE MÍNIMO", "TDE MÁXIMO", "TDE(%)", "TDE VALOR FIXO", "TDA MÍNIMO", _
"TDA MÁXIMO", "TDA(%)", "TDA VALOR FIXO", "TSB(%)", "TSB VALOR FIXO", "SUFRAMA VALOR FIXO", _
"SEGURO FLUVIAL MÍNIMO", "SEGURO FLUVIAL MÁXIMO", "SEGURO FLUVIAL(%)", _
"SEGURO FLUVIAL VALOR FIXO", "PEDÁGIO VALOR FIXO", "PEDÁGIO FRAÇÃO A CADA x KG", _
"FAIXA INÍCIAL DE PEDÁGIO", "FAIXA FINAL DE PEDÁGIO", "FAIXA VIGENTE SOBRE(NF ou Peso)", _
"PEDÁGIO VALOR FIXO", "PEDÁGIO FRAÇÃO A CADA x KG", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "COLETA VALOR FIXO", _
"FAIXA INICIAL DE COLETA", "FAIXA FINAL DE COLETA", "FAIXA VIGENTE SOBRE(NF ou Peso)", _
"COLETA VALOR FIXO", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "ENTREGA VALOR FIXO", _
"FAIXA INICIAL DE ENTREGA", "FAIXA FINAL DE ENTREGA", "FAIXA VIGENTE SOBRE(NF ou Peso)", _
"ENTREGA VALOR FIXO", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", ".")

b = Array("SEGURO MÁXIMO", "SEGURO(%)", "SEGURO VALOR FIXO", "FAIXA INICIAL DE SEGURO", _
"FAIXA FINAL DE SEGURO", "FAIXA VIGENTE SOBRE(NF ou Peso)", "SEGURO VALOR FIXO", _
"SEGURO(%)", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", _
"ADEME MÍNIMO", "ADEME MÁXIMO", "ADEME(%)", "ADEME VALOR FIXO", "FAIXA INICIAL DE ADEME", _
"FAIXA FINAL DE ADEME", "FAIXA VIGENTE SOBRE(NF ou Peso)", "ADEME VALOR FIXO", "ADEME(%)", _
"VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "EMEX MÍNIMO", "EMEX MÁXIMO", _
"EMEX(%)", "EMEX VALOR FIXO", "EMEX FRAÇÃO A CADA x KG", "FAIXA INICIAL DE EMEX", _
"FAIXA FINAL DE EMEX", "FAIXA VIGENTE SOBRE(NF ou Peso)", "EMEX(%)", "EMEX VALOR FIXO", _
"EMEX FRAÇÃO A CADA x KG", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "CTE MÍNIMO", "CTE MÁXIMO", "CTE(%)", _
"CTE VALOR FIXO", "CTE FRAÇÃO A CADA x KG", "FAIXA INICIAL DE CTE", "FAIXA FINAL DE CTE", _
"FAIXA VIGENTE SOBRE(NF ou Peso)", "CTE(%)", "CTE VALOR FIXO", "CTE FRAÇÃO A CADA x KG", _
"VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", _
"OUTRA TAXA MÍNIMO", "OUTRA TAXA MÁXIMO", "OUTRA TAXA(%)", "OUTRA TAXA VALOR FIXO", _
"OUTRA TAXA FRAÇÃO A CADA x KG", "FAIXA INICIAL DE OUTRA TAXA", "FAIXA FINAL DE OUTRA TAXA", _
"FAIXA VIGENTE SOBRE(NF ou Peso)", "OUTRA TAXA VALOR FIXO", "OUTRA TAXA FRAÇÃO A CADA x KG", _
"OUTRA TAXA(%)", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "SEGURO MÍNIMO", _
"VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", _
"SECCAT VALOR FIXO", "FAIXA FINAL DE SECCAT", "FAIXA VIGENTE SOBRE(NF ou Peso)", _
"FAIXA INICIAL DE SECCAT", "SECCAT VALOR FIXO")

c = Array("ICMS Incluso?(S/N)", "CUBAGEM(kg/m³)", "ISENÇÃO DE CUBAGEM(kg)", "LIMITE DE ALTURA(cm)", "LIMITE DE LARGURA(cm)", "LIMITE DE COMPRIMENTO(cm)", "CEP ORIGEM")
End Sub
