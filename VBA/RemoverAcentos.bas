Attribute VB_Name = "Módulo1"
Sub Botao_Processar_Acentos()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wbNovo As Workbook
    Dim rngOrigem As Range
    Dim linha As Long, col As Long
    Dim dataHoje As String
    Dim caminhoBase As String
    Dim caminho As String
    
    ' Define a planilha ativa como origem
    Set wsOrigem = ActiveSheet
    
    ' Define o intervalo de A a F até a última linha preenchida
    linha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    Set rngOrigem = wsOrigem.Range("A1:F" & linha)
    
    ' Cria um novo workbook
    Set wbNovo = Workbooks.Add
    Set wsDestino = wbNovo.Sheets(1)
    
    ' Copia os dados convertendo os acentos
    For linha = 1 To rngOrigem.Rows.Count
        For col = 1 To rngOrigem.Columns.Count
            wsDestino.Cells(linha, col).Value = Acento(CStr(rngOrigem.Cells(linha, col).Value))
        Next col
    Next linha
    
    ' Pega o caminho base da célula M1 da planilha ativa
    caminhoBase = wsOrigem.Range("M1").Value

    ' Garante que o caminho termine com "\"
    If Right(caminhoBase, 1) <> "\" Then caminhoBase = caminhoBase & "\"
    
    ' Cria a pasta se não existir
    If Dir(caminhoBase, vbDirectory) = "" Then
        MkDir caminhoBase
    End If

    ' Define o nome do arquivo com data de hoje
    dataHoje = Format(Date, "yyyy-mm-dd")
    
    ' Junta o caminho da célula com o nome do arquivo
    caminho = caminhoBase & "areas_bloqueio_" & dataHoje & ".csv"

    ' Salva o novo workbook em CSV UTF-8
    Application.DisplayAlerts = False ' Evita alertas de sobrescrita
    wbNovo.SaveAs Filename:=caminho, FileFormat:=xlCSVUTF8
    Application.DisplayAlerts = True
    
    MsgBox "Arquivo salvo com sucesso em: " & caminho
End Sub

Function Acento(Caract As String) As String
    Dim A As String
    Dim B As String
    Dim i As Integer

    Const AccChars = "ŠšŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖÙÚÛÜİàáâãäåçèéêëìíîïğñòóôõöùúûüıÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    For i = 1 To Len(AccChars)
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        Caract = Replace(Caract, A, B)
    Next

    Acento = Caract
End Function

