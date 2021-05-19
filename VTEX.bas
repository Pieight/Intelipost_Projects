Attribute VB_Name = "Macro_VTEX"
Dim cepicol As Double, cepfcol As Double, rowmax As Double, wscol As Double, wecol As Double, minimumvalcol As Double
Dim moneycostcol As Double, pricepercol As Double, exccol As Double, maxvolcol As Double, timecostcol As Double
Dim firstrange As Range, secondrange As Range, thirdrange As Range, fourthrange As Range, rangepeso() As Variant, array25() As Variant, rowmax25 As Double, divisor As Double, wb25 As Workbook

Public ceporigem1, icms, limitc, limitl, limita, cubagem, isencao

Sub main100()

'ActiveWorkbook.Save
Sheets("Dados").Activate

'Função que irá declarar e definir todas as variáveis necessárias para as funções
Call declarar

'Realiza o processo de classificação de cada um dos campos em ordem
Call classificar

'Cria uma matriz com todos os pesos possíveis da VTEX
Call verificacao

'Verifica a formatcao do prazo
Call format_prazo

'Faz um sort da matriz de pesos
Call sorting

'Funcao que cria uma matrix com todas informações necessárias de uma 2.5
Call construcao

'Criacao da 25 a partir da VTEX
Call cria_25

'Criar a versão final da tabela a partir da 2.5
Call go(wb25)

End Sub



Sub declarar()

rowmax = Range("A1").End(xlDown).Row
cepicol = Cells.Find("ZipCodeStart").Column
cepfcol = Cells.Find("ZipCodeEnd").Column
wscol = Cells.Find("WeightStart").Column
wecol = Cells.Find("WeightEnd").Column
moneycostcol = Cells.Find("AbsoluteMoneyCost").Column
pricepercol = Cells.Find("PricePercent").Column
exccol = Cells.Find("PriceByExtraWeight").Column
timecostcol = Cells.Find("TimeCost").Column

On Error Resume Next
maxvolcol = Cells.Find("MaxVolume").Column
minimumvalcol = Cells.Find("MinimumValueInsurance").Column

Set firstrange = Range(Cells(1, cepicol), Cells(rowmax, cepicol))
Set secondrange = Range(Cells(1, cepfcol), Cells(rowmax, cepfcol))
Set thirdrange = Range(Cells(1, wscol), Cells(rowmax, wscol))
Set fourthrange = Range(Cells(1, wecol), Cells(rowmax, wecol))

rowmax25 = 1
For i = 2 To rowmax - 1
    If Cells(i, cepicol) <> Cells(i + 1, cepicol) Then
        rowmax25 = rowmax25 + 1
    End If
Next i
End Sub

Sub classificar()
endcol = Range("A1").End(xlToRight).Column

ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=firstrange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=secondrange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=thirdrange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=fourthrange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range(Cells(1, 1), Cells(rowmax, endcol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Sub verificacao()
Dim pesoexists As Boolean
Dim limit As Double: limit = 0

'Verifica quais são todos os pesos existentes na VTEX
ReDim Preserve rangepeso(0)

'Linha para cuidar do caso dos pesos que estão em toneladas ao invés de kg
divisor = 1
tonelada = MsgBox("O peso da tabela está em gramas?", vbYesNo, "PESO")


If tonelada = 6 Then
    divisor = 1000
End If

For i = 2 To rowmax
    pesoexists = False
    For k = 0 To UBound(rangepeso)
        If rangepeso(k) = Cells(i, wecol) Then
            pesoexists = True
            Exit For
        End If
    Next k
    If Not (pesoexists) Then
            ReDim Preserve rangepeso(limit)
            rangepeso(limit) = Cells(i, wecol)
            limit = limit + 1
        End If
Next i


'FAzer um algoritmo para saber a procedencia dos pesos

End Sub



Sub format_prazo()

must_format_prazo = False

Set prazo_range = Range(Cells(2, timecostcol), Cells(rowmax, timecostcol))

prazo_range.Select

If Not (prazo_range.Cells.Find(".") Is Nothing) Then
    For i = 2 To rowmax
        word = Cells(i, timecostcol)
        Cells(i, timecostcol) = Replace(Cells(i, timecostcol), ".00:00:00", "")
    Next i
End If


End Sub



Sub sorting()

For i = 0 To UBound(rangepeso) - 1
        For j = 1 + i To UBound(rangepeso)
            If Val(rangepeso(j)) < Val(rangepeso(i)) Then
                pesotemp = rangepeso(j)
                rangepeso(j) = rangepeso(i)
                rangepeso(i) = pesotemp
            End If
        Next j
    Next i

End Sub
Sub construcao()
Dim pesoexists As Boolean: pesoexists = False
Dim i As Double: i = 0
Dim same_cep As Range


'Parte para verificar a a validação da range peso
'For k = 2 To rowmax
   ' For i = 0 To UBound(rangepeso)
        'If rangepeso(i) = Cells(k, wecol) Then Exit For
        'elseif rangepeso(i) <> Cells(k, wecol) and
    'Next

ReDim Preserve array25(0 To 5 + UBound(rangepeso), 0 To rowmax25)
array25(0, 0) = "CEPI": array25(1, 0) = "CEPF": array25(2, 0) = "PRAZO(DIAS ÚTEIS)": array25(UBound(array25, 1) - 1, 0) = "VALOR EXCEDENTE": array25(UBound(array25, 1), 0) = "FRETE VALOR SOBRE A NOTA(%)"

For i = 3 To UBound(rangepeso) + 3
    array25(i, 0) = rangepeso(i - 3)
Next i

linha_matriz = 1
inicial = 2

'tempoinicial = Time

For k = 1 To UBound(array25, 2)
    
    If Cells(inicial, cepicol) = "" Then Exit For
    
    linha_final = range_cep(Cells(inicial, cepicol), inicial)
    Set same_cep = Range(Cells(inicial, cepicol), Cells(linha_final, cepicol))
    
    array25(0, linha_matriz) = Cells(inicial, cepicol).Value: array25(1, linha_matriz) = Cells(inicial, cepfcol).Value: array25(2, linha_matriz) = Cells(inicial, timecostcol).Value
    array25(UBound(array25, 1) - 1, linha_matriz) = (Cells(linha_final, exccol).Value * divisor):  array25(UBound(array25, 1), linha_matriz) = Cells(linha_final, pricepercol).Value
    
    m = inicial
    For j = 3 To UBound(array25, 1) - 2
        For i = m To linha_final
            If Cells(i, wecol) = array25(j, 0) Then
                pesoexists = True
                array25(j, linha_matriz) = Cells(i, moneycostcol)
                m = i + 1
                Exit For
            End If
        Next i
        If pesoexists = False Then
            array25(j, linha_matriz) = Val("0.01")
        End If
        pesoexists = False
    Next j
    
    inicial = linha_final + 1
    linha_matriz = linha_matriz + 1
Next k

'MsgBox Minute(Time - tempoinicial) & " minutos " & Second(Time - tempoinicial) & " segundos"

For j = 3 To UBound(array25, 1) - 2
    array25(j, 0) = array25(j, 0) / divisor
Next j



End Sub




Sub cria_25()

Workbooks.Add

Set wb25 = ActiveWorkbook

end_row = UBound(array25, 2) + 1
end_col = UBound(array25, 1) + 1
Range(Cells(4, 3), Cells(end_row + 3, end_col + 2)).Select

Selection = WorksheetFunction.Transpose(array25)

With Range(Cells(4, 3), Cells(4, end_col + 2))
    .Interior.Color = 4697456
    .Font.Color = 16777215
End With

''

DADOS.Show

With Range("A4")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "ICMS Incluso?(S/N)"
End With


With Range("A7")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "CUBAGEM(kg/m³)"
End With

If ceporigem1 <> "" Then
    With Range("A10")
        .Interior.Color = 4697456
        .Font.Color = 16777215
        .Value = "CEP ORIGEM"
    End With
    Cells(11, 1) = ceporigem1
End If

If isencao <> "" Then
    Cells(14, 1) = isencao
    With Range("A13")
        .Interior.Color = 4697456
        .Font.Color = 16777215
         .Value = "ISENÇÃO DE CUBAGEM(kg)"
    End With

End If
With Range("A16")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "LIMITE DE ALTURA(cm)"
End With

With Range("A19")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "LIMITE DE LARGURA(cm)"
End With

With Range("A22")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "LIMITE DE COMPRIMENTO(cm)"
End With

Range("F1") = "TABELA DE FRETE POR PESO"
Range("F3") = "FAIXAS DE PESO (KG)"
   
With Range("F1:F3")
    .Interior.Color = 4697456
    .Font.Color = 16777215
End With


Cells(5, 1) = icms
Cells(23, 1) = limitc
Cells(20, 1) = limitl
Cells(17, 1) = limita
Cells(8, 1) = cubagem



collumexc = Cells.Find("VALOR EXCEDENTE").Column


If WorksheetFunction.Sum(Cells.Columns(collumexc)) > 0 Then
    dif = Cells(4, collumexc).Offset(0, -1) - Cells(4, collumexc).Offset(0, -2)
    If dif > 1500 Then
        Cells.Columns(collumexc - 1).Delete
    End If
End If

'end_money_col = end_col
'sub_end_money_col = end_col - 1

'final_row = Range(Cells(4, 3), Cells(4, 3)).End(xlDown).Row
'erro_ultima_coluna = True
'For i = 5 To final_row
    'If Cells(i, end_money_col) <> Cells(i, sub_end_money_col) Then
        'erro_ultima_coluna = False
        'Exit For
    'End If
'Next i

'If erro_ultima_coluna Then
    'Columns(end_money_col).Delete
'End If


Cells.EntireColumn.AutoFit

ActiveSheet.Name = "2.5"


End Sub


Function range_cep(ceptemp, linha_inicial)

'Função para achar a range do cep em específico
linha_final = 0
p = 0


Do While Cells(linha_inicial + p, cepicol) = ceptemp
    linha_final = linha_final + 1
    p = p + 1
Loop

range_cep = linha_inicial + linha_final - 1

End Function









