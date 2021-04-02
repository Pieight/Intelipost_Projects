Attribute VB_Name = "Módulo4"
Public mfil0
Public FSO1 As New FileSystemObject
Public Folder_client() As Variant
Public list_of_ids() As Variant
Public word() As Variant
Public Folder_files() As Variant
Public files_to_throw() As Variant
Public path() As Variant

Sub arrebatar()

Set FSO1 = CreateObject("Scripting.FileSystemObject")


Cells(1, 5) = Time

'Função para preencher todas as pastas existentes
Call array_folder("I:\03 Clientes")

'Função para retornar uma matriz apenas com os nomes necessários
Call ids

'Função para preencher a matriz  com os ids
Call preencher_matriz(word)

'Função para fazer o sorting da matriz de ids
Call bubbling_sort(list_of_ids)

'Função para jogar os arquivos nas respectivas pastas
Call throw


End Sub

Sub array_folder(hostfolder As String)
Dim F, i As Double

F = Dir(hostfolder & "\*", vbDirectory)
i = 0

'Criar a lista de diretórios disponíveis
Do While F <> ""
    If IsNumeric(Left(F, 1)) Then
        ReDim Preserve Folder_client(i)
        Folder_client(i) = F
        i = i + 1
    End If
    F = Dir
Loop
End Sub


Sub ids()
Dim id_count As Double, cdtrue As Boolean, v2true As Boolean, mfil0 As Variant

mfil0 = Dir("C:\Users\Paulo Henrique\Desktop\Resultados\*.xlsx")
j = 0
k = 0
Do While mfil0 <> ""
    v2true = False
    cdtrue = False
    und_count = 0
    For i = 1 To Len(mfil0)
        If Mid(mfil0, i, 1) = "_" Then
            und_count = und_count + 1
            If Not (IsNumeric(Mid(mfil0, i + 1, 1))) Then
                Exit For
            End If
        End If
        If Mid(mfil0, i, 1) = "C" Then
            If Mid(mfil0, i + 1, 1) = "D" Then
                cdtrue = True
            End If
        End If
        If Mid(mfil0, i, 1) = "v" Then
            If Mid(mfil0, i + 1, 1) = "2" Then
                If Mid(mfil0, i + 2, 1) = "_" Then
                    v2true = True
                End If
            End If
        End If
    Next i
    If (und_count = 3) And (v2true) And (cdtrue) Then
        ReDim Preserve word(j)
        word(j) = mfil0
        j = j + 1
    End If
    mfil0 = Dir
Loop

End Sub

Sub preencher_matriz(matrix)
Dim id_client As String, id_transp As String, id_call As String, first_letter As Double, last_letter As Double

ReDim Preserve list_of_ids(0 To UBound(matrix), 0 To 2)
For i = 0 To UBound(matrix)
    first_letter = 1
    k = 0
    For j = 1 To Len(matrix(i))
        If Mid(word(i), j, 1) = " " Then
            list_of_ids(i, k) = Mid(matrix(i), first_letter, j - first_letter)
            Exit For
        End If
        If Mid(word(i), j, 1) = "_" Then
            list_of_ids(i, k) = Mid(matrix(i), first_letter, j - first_letter)
            first_letter = j + 1
            k = k + 1
        End If
    Next j
Next i

End Sub

Sub bubbling_sort(matrix)
Dim i As Double, mintemp As Variant, j As Double, matriztemp() As Variant
ReDim matriztemp(0 To 0, 0 To 2)

    For i = 0 To UBound(matrix, 1) - 1
        For j = 1 + i To UBound(matrix, 1)
            If Val(matrix(j, 1)) < Val(matrix(i, 1)) Then
                matriztemp(0, 0) = matrix(i, 0)
                matriztemp(0, 1) = matrix(i, 1)
                matriztemp(0, 2) = matrix(i, 2)
        
                matrix(i, 0) = matrix(j, 0)
                matrix(i, 1) = matrix(j, 1)
                matrix(i, 2) = matrix(j, 2)
        
                matrix(j, 0) = matriztemp(0, 0)
                matrix(j, 1) = matriztemp(0, 1)
                matrix(j, 2) = matriztemp(0, 2)
    
            End If
        Next j
    Next i

End Sub

Sub throw()
Dim id_client As Variant, id_chamado As Variant, id_tp As Variant, mfil0 As Variant, moved_file As Boolean, found_client_folder As Boolean, tabelas_de_frete As Boolean
moved_file = False
index_files_moved = 0
For i = 0 To UBound(list_of_ids, 1)
client_folder_created:
    found_client_folder = False
    For j = 0 To UBound(Folder_client)
        If (list_of_ids(i, 1) = Left(Folder_client(j), Len(list_of_ids(i, 1))) And ((Mid(Folder_client(j), Len(list_of_ids(i, 1)) + 1, 1) = " ") Or ((Mid(Folder_client(j), Len(list_of_ids(i, 1)) + 1, 1) = "")))) Then
            found_client_past = True
            'MsgBox list_of_ids(i, 1) & " " & Folder_client(j)
            mfil0 = Dir("I:\03 Clientes\" & Folder_client(j) & "\", vbDirectory)
            tabelas_de_frete = False
            Do While mfil0 <> ""
            
                If Left(mfil0, 3) = "2 T" Then
same_id:
method_created:
                    tabelas_de_frete = True
                    mfil1 = Dir("I:\03 Clientes\" & Folder_client(j) & "\" & mfil0 & "\", vbDirectory)
                    moved_file = False
                    Do While mfil1 <> ""
                        If (list_of_ids(i, 2) = Left(mfil1, Len(list_of_ids(i, 2)))) Then
                            If (Mid(mfil1, Len(list_of_ids(i, 2)) + 1, 1) = " ") Or ((Mid(mfil1, Len(list_of_ids(i, 2)) + 1, 1) = "")) Then
                                'MsgBox list_of_ids(i, 2) & " " & mfil1
                                Call files_to_move(list_of_ids(i, 0))
                                For s = 0 To UBound(files_to_throw)
                                    If files_to_throw(0) <> "" Then
                                        ReDim Preserve path(index_files_moved)
                                        path(index_files_moved) = "I:\03 Clientes\" & Folder_client(j) & "\" & mfil0 & "\" & mfil1 & "\" & files_to_throw(s)
                                        FSO1.MoveFile "C:\Users\Paulo Henrique\Desktop\Resultados\" & files_to_throw(s), "I:\03 Clientes\" & Folder_client(j) & "\" & mfil0 & "\" & mfil1 & "\" & files_to_throw(s)
                                        moved_file = True
                                        index_files_moved = index_files_moved + 1
                                    End If
                                Next s
                                If i <> UBound(list_of_ids, 1) Then
                                    If list_of_ids(i, 1) = list_of_ids(i + 1, 1) Then
                                        i = i + 1
                                        GoTo same_id
                                     Else: GoTo next_id
                                    End If
                                Else
                                    Cells(2, 5) = Time
                                    Call printing_paths(path)
                                    ActiveWorkbook.Save
                                    End
                                End If
                            End If
                        End If
                        mfil1 = Dir
                    Loop
                    If moved_file = False Then
                        Call crate_new_method_folder(CDbl(list_of_ids(i, 2)), CStr(Sheets("Arrebatador").Cells.Find(list_of_ids(i, 2), lookat:=xlWhole).Offset(0, 1)), CStr(mfil0), j)
                        GoTo method_created
                    End If
                End If
                mfil0 = Dir
            Loop
            If tabelas_de_frete = False Then
                GoTo next_id
            End If
        End If
    Next j
    If found_client_folder = False Then
        create_new_client_folder (CDbl(list_of_ids(i, 1)))
        Call array_folder("I:\03 Clientes")
        GoTo client_folder_created
    End If
next_id:
Next i

End Sub

Sub files_to_move(chamado)
Call preencher_folder
y = 0
ReDim files_to_throw(y)
files_to_throw(y) = ""
For h = 0 To UBound(Folder_files)
    If chamado = Left(Folder_files(h), Len(chamado)) Then
        ReDim Preserve files_to_throw(y)
        files_to_throw(y) = Folder_files(h)
        y = y + 1
    End If
Next

End Sub

Sub preencher_folder()
Dim k_folder As Double
mfil3 = Dir("C:\Users\Paulo Henrique\Desktop\Resultados\*.xl*")
k_folder = 0
Do While mfil3 <> ""
    ReDim Preserve Folder_files(k_folder)
    Folder_files(k_folder) = mfil3
    k_folder = k_folder + 1
    mfil3 = Dir
Loop
If k_folder = 0 Then
    ReDim Preserve Folder_files(k_folder)
    Folder_files(k_folder) = ""
End If
End Sub

Sub crate_new_method_folder(id_method As Double, name_method As String, path As String, index)
 FSO1.CopyFolder "C:\Users\Paulo Henrique\Desktop\Resultados\Modelo", "I:\03 Clientes\" & Folder_client(index) & "\" & path & "\" & id_method & " - " & name_method
End Sub

Sub create_new_client_folder(id_method As Double)
client_name = InputBox("Não encontramos a pasta do cliente que possui o id: " & id_method & ". Digite o nome dele para que criemos uma pasta no Intelidrive")
FSO1.CopyFolder "C:\Users\Paulo Henrique\Desktop\Resultados\00 - Modelo", "I:\03 Clientes\" & id_method & " - " & client_name
End Sub

Sub printing_paths(path_vector)
Sheets("Historico").Activate
For g = 0 To UBound(path_vector)
    Range("A1").End(xlDown).Offset(1, 0) = path_vector(g)
Next g
Sheets("Arrebatador").Activate
End Sub
