Attribute VB_Name = "M�dulo1"
Dim cepicol As Double, cepfcol As Double, rowmax As Double, wscol As Double, wecol As Double, minimumvalcol As Double
Dim moneycostcol As Double, pricepercol As Double, exccol As Double, maxvolcol As Double, timecostcol As Double
Dim firstrange As range, secondrange As range, thirdrange As range, fourthrange As range, rangepeso() As Variant, array25() As Variant, rowmax25 As Double
Sub main()

ActiveWorkbook.Save

'Fun��o que ir� declarar e definir todas as vari�veis necess�rias para as fun��es
Call declarar

'Realiza o processo de classifica��o de cada um dos campos em ordem
Call classificar

'Cria uma matriz com todos os pesos poss�veis da VTEX
Call verificacao

'Faz um sort da matriz de pesos
Call sorting

'Funcao que cria uma matrix com todas informa��es necess�rias de uma 2.5
Call construcao

'Criacao da 25 a partir da VTEX
Call cria_25


End Sub



Sub declarar()

rowmax = range("A1").End(xlDown).Row
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

Set firstrange = range(Cells(1, cepicol), Cells(rowmax, cepicol))
Set secondrange = range(Cells(1, cepfcol), Cells(rowmax, cepfcol))
Set thirdrange = range(Cells(1, wscol), Cells(rowmax, wscol))
Set fourthrange = range(Cells(1, wecol), Cells(rowmax, wecol))

rowmax25 = 1
For i = 2 To rowmax - 1
    If Cells(i, cepicol) <> Cells(i + 1, cepicol) Then
        rowmax25 = rowmax25 + 1
    End If
Next i
End Sub

Sub classificar()


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
        .SetRange range(Cells(1, 1), Cells(rowmax, minimumvalcol))
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

'Verifica quais s�o todos os pesos existentes na VTEX
ReDim Preserve rangepeso(0)

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
Dim same_cep As range


'Parte para verificar a a valida��o da range peso
'For k = 2 To rowmax
   ' For i = 0 To UBound(rangepeso)
        'If rangepeso(i) = Cells(k, wecol) Then Exit For
        'elseif rangepeso(i) <> Cells(k, wecol) and
    'Next

ReDim Preserve array25(0 To 5 + UBound(rangepeso), 0 To rowmax25)
array25(0, 0) = "CEPI": array25(1, 0) = "CEPF": array25(2, 0) = "PRAZO(DIAS �TEIS)": array25(UBound(array25, 1) - 1, 0) = "VALOR EXCEDENTE": array25(UBound(array25, 1), 0) = "FRETE VALOR SOBRE A NOTA(%)"

For i = 3 To UBound(rangepeso) + 3
    array25(i, 0) = rangepeso(i - 3)
Next i

linha_matriz = 1
inicial = 2
For k = 1 To UBound(array25, 2)
    
    linha_final = range_cep(Cells(inicial, cepicol), inicial)
    Set same_cep = range(Cells(inicial, cepicol), Cells(linha_final, cepicol))
    
    array25(0, linha_matriz) = Cells(inicial, cepicol).Value: array25(1, linha_matriz) = Cells(inicial, cepfcol).Value: array25(2, linha_matriz) = Cells(inicial, timecostcol).Value
    array25(UBound(array25, 1) - 1, linha_matriz) = Cells(linha_final, exccol).Value:  array25(UBound(array25, 1), linha_matriz) = Cells(linha_final, pricepercol).Value
    
    For j = 3 To UBound(array25, 1) - 2
        For i = inicial To linha_final
            If Cells(i, wecol) = array25(j, 0) Then
                pesoexists = True
                array25(j, linha_matriz) = Cells(i, moneycostcol)
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








End Sub

Sub cria_25()

Workbooks.Add

Set wb25 = ActiveWorkbook

end_row = UBound(array25, 2) + 1
end_col = UBound(array25, 1) + 1
range(Cells(4, 3), Cells(end_row + 3, end_col + 2)).Select

Selection = WorksheetFunction.Transpose(array25)

With range(Cells(4, 3), Cells(4, end_col + 2))
    .Interior.Color = 4697456
    .Font.Color = 16777215
End With

With range("A4")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "ICMS Incluso?(S/N)"
End With
range("A5") = "N"

With range("A7")
    .Interior.Color = 4697456
    .Font.Color = 16777215
     .Value = "CUBAGEM(kg/m�)"
End With
range("A8") = 0

range("F1") = "TABELA DE FRETE POR PESO"
range("F3") = "FAIXAS DE PESO (KG)"
   
With range("F1:F3")
    .Interior.Color = 4697456
    .Font.Color = 16777215
End With

Cells.EntireColumn.AutoFit

End Sub


Function range_cep(ceptemp, linha_inicial)

'Fun��o para achar a range do cep em espec�fico
linha_final = 0
p = 0

Do While Cells(linha_inicial + p, cepicol) = ceptemp
    linha_final = linha_final + 1
    p = p + 1
Loop

range_cep = linha_inicial + linha_final - 1

End Function
