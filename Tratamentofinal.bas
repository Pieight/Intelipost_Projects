Attribute VB_Name = "Módulo2"
'Public a() As Variant, b() As Variant
Public cels As Range
Public i As Double
Public rowmin As Double
Public rowmax As Double
Public colmin As Integer
Public colmax As Integer
Public colcepi As Double
Public colcepf As Double
Public coldel As Double
Public colval As Double

Sub TratamentoFinal()
Attribute TratamentoFinal.VB_ProcData.VB_Invoke_Func = "t\n14"
Dim categorias As Variant, z As Double, contador As Integer, valorexists As Boolean, cabecalho As Boolean, celcabecalho As Range, comparacao As Boolean
Set cels = Cells.Find("CEPI")
rowmin = cels.row
colmin = cels.Column
rowmax = WorksheetFunction.CountA(ActiveSheet.Columns(colmin))
colmax = WorksheetFunction.CountA(ActiveSheet.Rows(rowmin))
coldel = Cells.Find("PRAZO(DIAS ÚTEIS)").Column
colval = Cells.Find("PRAZO(DIAS ÚTEIS)").Offset(0, 1).Column
    If Cells(rowmin, colval) = "FRETE TOTAL MÍNIMO" Then
        colval = colval + 1
    End If

' Verifica se as colunas estão certas
comparacao = False
valorexists = False
cabecalho = False
Set celcabecalho = Range("A3")
Call declaração_variavel
      i = colmin
     Do While ActiveSheet.Cells(rowmin, i).Value <> ""
        For z = 0 To UBound(a)
            If ((ActiveSheet.Cells(rowmin, i).Offset(-1, 0).Value = "FAIXAS DE PESO (KG)")) Then
                cabecalho = True
                Set celcabecalho = ActiveSheet.Cells(rowmin, i).Offset(-1, 0)
            End If
            If ActiveSheet.Cells(rowmin, i).Offset(-3, 0).Value = "TABELA DE FRETE POR COMPARAÇÃO" Then
                comparacao = True
            End If
            If (ActiveSheet.Cells(rowmin, i).Value = a(z)) Or (ActiveSheet.Cells(rowmin, i).Value = b(z) Or (IsNumeric(ActiveSheet.Cells(rowmin, i).Value))) Then
                ActiveSheet.Cells(rowmin, i).Interior.Color = 4697456
                ActiveSheet.Cells(rowmin, i).Font.Color = 16777215
                If ActiveSheet.Cells(rowmin, i).Value = "VALOR EXCEDENTE" Then
                    valorexists = True
                End If
                Exit For
            ElseIf z = 67 Then
                ActiveSheet.Cells(rowmin, i).Interior.Color = vbRed
                MsgBox ("A categoria " & ActiveSheet.Cells(rowmin, i) & " está errada!")
                End
            End If
        Next z
        i = i + 1
      Loop
If Not (valorexists) Then
    MsgBox ("Está faltando o 'VALOR EXCEDENTE' na tabela!!!")
End If
      
If ((Not (cabecalho)) Or (celcabecalho.Offset(-2, 0) <> "TABELA DE FRETE POR PESO") Or (Not (IsNumeric(celcabecalho.Offset(1, 0))))) Then
    If Not (comparacao) Then
        MsgBox ("Há um problema no cabeçalho da tabela!!!")
        End
    End If
End If

'Aplica um tratamento específico para cada tipo de categoria
categorias = Range(ActiveSheet.Cells(rowmin, colmin), ActiveSheet.Cells(rowmin, colmax + 1))
i = 0
Do While cels.Offset(0, i).Value <> ""
    Select Case cels.Offset(0, i)
    Case "FRETE VALOR SOBRE A NOTA(%)", "GRIS(%)", "TRT(%)", "TDA(%)", "TSB(%)", "TDE(%)", "SEGURO FLUVIAL(%)", "SEGURO(%)", "ADEME(%)", "EMEX(%)", "CTE(%)", "OUTRA TAXA(%)", "% SOBRE A NF"
        Call achar_notinteger
        Call preencher_zero
        
        For z = 1 To rowmax - 1
            If cels.Offset(z, i) <> 0 Then
                If ((cels.Offset(z, i).NumberFormat = "0.00%") Or (cels.Offset(z, i).NumberFormat = "0%") _
                Or (cels.Offset(z, i).NumberFormat = "0.000%") Or (cels.Offset(z, i).NumberFormat = "0.0%") Or (cels.Offset(z, i).NumberFormat = "0.0000%") _
                Or (cels.Offset(z, i).NumberFormat = "0.00000%") Or (cels.Offset(z, i).NumberFormat = "0.000000%")) Then
                    Call Tirar_Porcentagem_function(cels.Offset(1, i))
                Else
                    Exit For
                End If
            End If
        Next z

        Call geral
        
    Case "CEPI", "CEPF", "PRAZO(DIAS ÚTEIS)", "PEDÁGIO FRAÇÃO A CADA x KG", "EMEX FRAÇÃO A CADA x KG", "CTE FRAÇÃO A CADA x KG", "OUTRA TAXA FRAÇÃO A CADA x KG"
        
        Call achar_vazio
        Call achar_notinteger
        Call geral
    Case "GRIS MÍNIMO", "GRIS MÁXIMO", "TAS VALOR FIXO", "TRT MÍNIMO", "TRT MÁXIMO", "TRT VALOR FIXO", "TDE MÍNIMO", "TDE MÁXIMO", "TDE VALOR FIXO", "TDA MÍNIMO", "TDA MÁXIMO", "TDA VALOR FIXO"
        Call preencher_zero
        Call achar_notinteger
        Call moeda
    Case "FRETE TOTAL MÍNIMO", "VALOR EXCEDENTE", "FRETE MÍNIMO", "TSB VALOR FIXO", "SUFRAMA VALOR FIXO", "SEGURO FLUVIAL MÍNIMO", "SEGURO FLUVIAL MÁXIMO", "SEGURO FLUVIAL VALOR FIXO", "PEDÁGIO VALOR FIXO"
        Call preencher_zero
        Call achar_notinteger
        Call moeda
    Case "COLETA VALOR FIXO", "ENTREGA VALOR FIXO", "SEGURO MÁXIMO", "SEGURO VALOR FIXO", "ADEME MÍNIMO", "ADEME MÁXIMO", "ADEME VALOR FIXO", "EMEX MÍNIMO", "EMEX MÁXIMO", "EMEX VALOR FIXO", "VALOR POR KG"
        Call preencher_zero
        Call achar_notinteger
        Call moeda
    Case "CTE MÍNIMO", "CTE MÁXIMO", "CTE VALOR FIXO", "OUTRA TAXA VALOR FIXO", "OUTRA TAXA MÁXIMO", "OUTRA TAXA MÍNIMO", "OUTRA TAXA VALOR FIXO", "SEGURO MÍNIMO", "SECCAT VALOR FIXO", "SECCAT VALOR FIXO"
        Call preencher_zero
        Call achar_notinteger
        Call moeda
    Case "FAIXA VIGENTE SOBRE(NF ou Peso)", "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)", "VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)"
        Call achar_integer
        Call conformity(cels.Offset(0, i).Value)
        Call geral
    End Select
    If IsNumeric(cels.Offset(0, i)) Then
        Call achar_notinteger
        Call moeda
    End If
    If (cels.Offset(0, i) = "CEPI") Or (cels.Offset(0, i) = "CEPF") Then
        Call tirar_espaco
    End If
    Call achar_vazio
    i = i + 1
Loop

'Procura por erros na coluna 1



For i = 1 To 40
    If (Not (IsNumeric(ActiveSheet.Cells(i, 1))) And (ActiveSheet.Cells(i, 1) <> "S") And (ActiveSheet.Cells(i, 1) <> "N")) Then
        For z = 0 To UBound(c)
            If c(z) = ActiveSheet.Cells(i, 1) Then
                Exit For
            ElseIf z = UBound(c) Then
                MsgBox ("O campo " & ActiveSheet.Cells(i, 1) & " está incorreto!")
                ActiveSheet.Cells(i, 1).Interior.Color = vbRed
                End
            End If
        Next z
    End If
Next i



For i = 1 To 34
    If ActiveSheet.Cells(i, 1) = "ICMS Incluso?(S/N)" Then
        contador = contador + 1
        If (ActiveSheet.Cells(i, 1).Offset(1, 0) <> "S") And (ActiveSheet.Cells(i, 1).Offset(1, 0) <> "N") Then
            MsgBox ("Há um erro no campo ICMS Incluso?(S/N)!")
            End
        End If
    ElseIf ActiveSheet.Cells(i, 1) = "CUBAGEM(kg/m³)" Then
        contador = contador + 1
        If Not (IsNumeric(ActiveSheet.Cells(i, 1).Offset(1, 0).Value)) Or ActiveSheet.Cells(i, 1).Offset(1, 0) = "" Then
            MsgBox ("Há um erro no campo CUBAGEM(kg/m³)!")
            End
        End If
    ElseIf (ActiveSheet.Cells(i, 1) = "ISENÇÃO DE CUBAGEM(kg)") Or (ActiveSheet.Cells(i, 1) = "LIMITE DE ALTURA(cm)") Or (ActiveSheet.Cells(i, 1) = "LIMITE DE LARGURA(cm)") Or (ActiveSheet.Cells(i, 1) = "LIMITE DE COMPRIMENTO(cm)") Or (ActiveSheet.Cells(i, 1) = "CEP ORIGEM") Then
        If Not (IsNumeric(ActiveSheet.Cells(i, 1).Offset(1, 0).Value)) Or ActiveSheet.Cells(i, 1).Offset(1, 0) = "" Then
            MsgBox ("Há um erro em algum campo da primeira coluna!")
        End
        End If
    End If
Next i
If contador <> 2 Then
    MsgBox ("Falta campos na primeira coluna!")
End If


'Trocar os CEPs maiores
colcepi = Cells.Find("CEPI").Column
colcepf = Cells.Find("CEPF").Column

For z = 1 To rowmax - 1
    If ActiveSheet.Cells(rowmin, colcepi).Offset(z, 0) > ActiveSheet.Cells(rowmin, colcepf).Offset(z, 0) Then
        tempcep = ActiveSheet.Cells(rowmin, colcepi).Offset(z, 0)
        ActiveSheet.Cells(rowmin, colcepi).Offset(z, 0) = ActiveSheet.Cells(rowmin, colcepf).Offset(z, 0)
        ActiveSheet.Cells(rowmin, colcepf).Offset(z, 0) = tempcep
    End If
Next z

If ActiveSheet.AutoFilterMode = True Then
   ActiveSheet.AutoFilterMode = False
End If

colr = Cells.Find("CEPI").End(xlToRight).Column
rowr = Cells.Find("CEPI").End(xlDown).row

Range(Cells.Find("CEPI"), Cells(rowr, colr)).Select

Selection.Sort key1:=Range("C4:C" & rowmax), key2:=Range("D4:D" & rowmax), Order1:=xlAscending, Order2:=xlAscending, Header:=x1yes




'Cells.Find("CEPI").Offset(1, 0).Select
'Selection.AutoFilter

'ActiveWorkbook.Worksheets("2.5").AutoFilter.Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("2.5").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("C4:C" & rowmax), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    'With ActiveWorkbook.Worksheets("2.5").AutoFilter.Sort
        '.Header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
       ' .Apply
    'End With

Application.ScreenUpdating = False

Set wk25 = ActiveSheet
Range("A4:A30").Select
Selection.Copy
Worksheets.Add
Set wka = ActiveSheet
wka.Range("A4").Select
ActiveSheet.Paste

wk25.Activate

For z = 1 To rowmax - 1
    If Cells(rowmin + z + 1, colcepf) = "" Then Exit For
    If Cells(rowmin + z, colcepi) = Cells(rowmin + z + 1, colcepi) Then
        If Cells(rowmin + z, colcepf) = Cells(rowmin + z + 1, colcepf) Then
            'Cells(rowmin + z + 1, colcepf).Select
            If Cells(rowmin + z, coldel) > Cells(rowmin + z + 1, coldel) Then
                Rows((rowmin + z + 1) & ":" & (rowmin + z + 1)).Delete
                z = z - 1
            ElseIf Cells(rowmin + z, coldel) < Cells(rowmin + z + 1, coldel) Then
                Rows((rowmin + z) & ":" & (rowmin + z)).Delete
                z = z - 1
            Else
                If Cells(rowmin + z, colval) >= Cells(rowmin + z + 1, colval) Then
                    Rows((rowmin + z + 1) & ":" & (rowmin + z + 1)).Delete
                    z = z - 1
                ElseIf Cells(rowmin + z, colval) < Cells(rowmin + z + 1, colval) Then
                    Rows((rowmin + z) & ":" & (rowmin + z)).Delete
                    z = z - 1
                End If
            End If
        End If
    End If
Next z

wka.Activate
wka.Range("A4:A30").Select
Selection.Copy
wk25.Activate
wk25.Range("A4").Select
ActiveSheet.Paste

Application.DisplayAlerts = False
wka.Delete
Application.DisplayAlerts = True

ActiveSheet.AutoFilterMode = False
Range("A1").Select

Cells.EntireColumn.AutoFit
Application.ScreenUpdating = True

End Sub


Sub achar_vazio()
On Error GoTo errorhandl
For z = 1 To rowmax - 1
    If cels.Offset(z, i) = "" Then
        MsgBox ("Há células vazias na categoria " & cels.Offset(0, i) & "!")
        End
    End If
Next z
Exit Sub

errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub geral()
On Error GoTo errorhandl
Range(cels.Offset(1, i), cels.Offset(rowmax - 1, i)).Select
Selection.NumberFormat = "General"
Exit Sub
errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub moeda()
On Error GoTo errorhandl
Range(cels.Offset(1, i), cels.Offset(rowmax - 1, i)).Select
Selection.NumberFormat = "$ #,##0.00"
Exit Sub
errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub achar_notinteger()
On Error GoTo errorhandl
For z = 1 To rowmax - 1
    'cels.Offset(z, i) = Val(cels.Offset(z, i))
    If Not (IsNumeric(cels.Offset(z, i)) Or (cels.Offset(z, i) = "#N/D")) Then
        MsgBox ("Há células que não são números na categoria " & cels.Offset(0, i) & "!")
        End
    End If
Next z
Exit Sub
errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub achar_integer()
On Error GoTo errorhandl
For z = 1 To rowmax - 1
    If ((IsNumeric(cels.Offset(z, i))) Or (cels.Offset(z, i) = "#N/D")) Then
        MsgBox ("Há células que são números ou não condizem com a coluna na categoria " & cels.Offset(0, i) & "!")
        End
    End If
Next z
Exit Sub
errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub preencher_zero()
On Error GoTo errorhandl
For z = 1 To rowmax - 1
    If cels.Offset(z, i) = "" Then
        cels.Offset(z, i) = 0
    End If
Next z
Exit Sub

errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub Tirar_Porcentagem_function(cel As Range)
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

Sub conformity(category As Variant)
On Error GoTo errorhandl

If category = "FAIXA VIGENTE SOBRE(NF ou Peso)" Then
    For z = 1 To rowmax - 1
        If ((cels.Offset(z, i) <> "NF") And (cels.Offset(z, i) <> "Peso")) Then
            MsgBox ("Há células na categoria " & category & " que estão erradas.")
            End
        End If
    Next z
ElseIf category = "VALOR DE FAIXA SOMA COM VALOR GERAL?(S/N)" Then
    For z = 1 To rowmax - 1
        If ((cels.Offset(z, i) <> "S") And (cels.Offset(z, i) <> "N")) Then
            MsgBox ("Há células na categoria " & category & " que estão erradas.")
            End
        End If
    Next z
ElseIf category = "VALOR SOMADO VIGENTE SOBRE FAIXA OU VALOR COMPLETO(F/VC)" Then
    For z = 1 To rowmax - 1
        If ((cels.Offset(z, i) <> "F") And (cels.Offset(z, i) <> "VC")) Then
            MsgBox ("Há células na categoria " & category & " que estão erradas.")
            End
        End If
    Next z
End If
Exit Sub

errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub


Sub tirar_espaco()
Dim palavra As String
On Error GoTo errorhandl

For z = 1 To rowmax - 1
    palavra = cels.Offset(z, i)
    For h = 1 To Len(palavra)
        If Not (IsNumeric(Mid(palavra, h, 1))) Then
            cels.Offset(z, i).Select
            palavra = Replace(palavra, Mid(palavra, h, 1), "")
            cels.Offset(z, i) = palavra
        End If
    Next h
Next z

Exit Sub

errorhandl:
   MsgBox ("Há células com algum problema na categoria " & cels.Offset(0, i) & "!")
        End
End Sub



