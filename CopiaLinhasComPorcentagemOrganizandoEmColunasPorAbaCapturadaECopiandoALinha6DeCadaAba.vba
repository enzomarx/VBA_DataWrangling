Sub CopiarLinhasComPorcentagem()
    Dim ws As Worksheet
    Dim wsTemp As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim i As Long, j As Long
    Dim linhaJaCopiada As Collection
    Dim celula As Range
    Dim linhaIndice As Long
    Dim linhaNumero As Variant
    Dim wsOriginal As Worksheet
    Dim celulaLinha6 As Range
    Dim tempRow As Long
    
    ' Cria uma nova planilha temporária para armazenamento das linhas copiadas
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets("TempCopySheet")
    If wsTemp Is Nothing Then
        Set wsTemp = ThisWorkbook.Worksheets.Add
        wsTemp.Name = "TempCopySheet"
    Else
        wsTemp.Cells.Clear ' Limpa o conteúdo se a planilha já existir
    End If
    On Error GoTo 0
    
    ' Inicializa a coleção para armazenar os números das linhas a serem copiadas
    Set linhaJaCopiada = New Collection
    
    ' Adiciona cabeçalhos para a planilha temporária
    wsTemp.Cells(1, 1).Value = "Aba Original"
    wsTemp.Cells(1, 2).Value = "Linha"
    wsTemp.Cells(1, 3).Value = "Dados"
    
    ' Percorre todas as planilhas
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsTemp.Name Then ' Não processa a planilha temporária
            ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' Verifica cada célula no intervalo usado
            For i = 1 To ultimaLinha
                For j = 1 To ultimaColuna
                    Set celula = ws.Cells(i, j)
                    If ContainsPercentSymbol(celula) Then
                        ' Adiciona a linha à coleção se ainda não foi copiada
                        On Error Resume Next
                        linhaJaCopiada.Add Item:=i & "|" & ws.Name, Key:=CStr(i) & "|" & ws.Name
                        On Error GoTo 0
                        Exit For ' Se já encontrou uma célula com '%' na linha, não precisa verificar o resto da linha
                    End If
                Next j
            Next i
            
            ' Copia a linha 6 da aba atual
            tempRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row + 1
            ws.Rows(6).Copy wsTemp.Rows(tempRow)
            wsTemp.Cells(tempRow, 1).Value = ws.Name
        End If
    Next ws
    
    ' Copia as linhas identificadas para a nova planilha
    If linhaJaCopiada.Count > 0 Then
        linhaIndice = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row + 1
        For Each linhaNumero In linhaJaCopiada
            Dim partes() As String
            partes = Split(linhaNumero, "|")
            Set wsOriginal = ThisWorkbook.Worksheets(partes(1))
            If partes(0) <= wsOriginal.Cells(wsOriginal.Rows.Count, 1).End(xlUp).Row Then
                wsOriginal.Rows(CInt(partes(0))).Copy wsTemp.Rows(linhaIndice)
                wsTemp.Cells(linhaIndice, 1).Value = partes(1) ' Indica a aba original
                linhaIndice = linhaIndice + 1
            End If
        Next linhaNumero
        
        MsgBox "Linhas com valores em porcentagem e a linha 6 de todas as abas foram copiadas para a nova planilha.", vbInformation
    Else
        MsgBox "Nenhuma linha com valores em porcentagem foi encontrada.", vbExclamation
    End If
End Sub

' Função para verificar se a célula contém o caractere '%' em valores gerais
Function ContainsPercentSymbol(celula As Range) As Boolean
    Dim valor As String
    On Error Resume Next
    valor = CStr(celula.Value) ' Converte o valor da célula para texto
    ContainsPercentSymbol = InStr(1, valor, "%", vbTextCompare) > 0
    On Error GoTo 0
End Function
