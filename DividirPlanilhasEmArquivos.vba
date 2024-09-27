Sub DividirPlanilhasEmArquivos()
    Dim ws As Worksheet
    Dim numArquivos As Integer
    Dim numAbasPorArquivo As Integer
    Dim i As Integer, j As Integer
    Dim novoArquivo As Workbook
    Dim totalAbas As Integer
    Dim abasRestantes As Integer
    Dim arquivoPath As String

    ' Define o número de arquivos para dividir
    numArquivos = InputBox("Digite o número de arquivos desejados (1 a 20):", "Número de Arquivos")
    
    ' Valida o número de arquivos
    If numArquivos < 1 Or numArquivos > 20 Then
        MsgBox "Número de arquivos inválido. Deve ser entre 1 e 20.", vbExclamation
        Exit Sub
    End If

    ' Conta o total de abas na planilha original
    totalAbas = ThisWorkbook.Worksheets.Count
    
    ' Calcula o número de abas por arquivo
    numAbasPorArquivo = Application.WorksheetFunction.Ceiling(totalAbas / numArquivos, 1)
    
    ' Caminho para salvar os novos arquivos
    arquivoPath = ThisWorkbook.Path & "\PlanilhasParte"
    
    ' Cria novos arquivos e copia as abas
    For i = 1 To numArquivos
        Set novoArquivo = Workbooks.Add
        ' Copia as abas para o novo arquivo
        For j = 1 To numAbasPorArquivo
            If (i - 1) * numAbasPorArquivo + j <= totalAbas Then
                ThisWorkbook.Worksheets((i - 1) * numAbasPorArquivo + j).Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
            End If
        Next j
        
        ' Remove a planilha padrão criada ao adicionar um novo arquivo
        Application.DisplayAlerts = False
        novoArquivo.Sheets(1).Delete
        Application.DisplayAlerts = True
        
        ' Salva o novo arquivo
        novoArquivo.SaveAs Filename:=arquivoPath & i & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        novoArquivo.Close SaveChanges:=False
    Next i
    
    MsgBox "Planilhas divididas em " & numArquivos & " arquivos com sucesso.", vbInformation
End Sub
