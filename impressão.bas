Attribute VB_Name = "Módulo2"
Sub PrintPages()
    Dim ws As Worksheet
    Dim listaFuncionarios As String
    Dim escolha As Variant
    Dim encontrou As Boolean

    ' Monta lista de planilhas (exclui as que não devem aparecer)
    listaFuncionarios = ""
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Gerar Pontos" And _
           ws.Name <> "Molde" And _
           ws.Name <> "Referências" And _
           ws.Name <> "Feriados" And _
           LCase(ws.Name) <> "estag ianara" Then
            listaFuncionarios = listaFuncionarios & ws.Name & vbCrLf
        End If
    Next ws

    ' Mostra janela para escolher
    escolha = InputBox( _
        "Digite o nome de um funcionário para imprimir" & vbCrLf & _
        "OU digite TODOS para imprimir tudo:" & vbCrLf & vbCrLf & _
        listaFuncionarios, _
        "Escolher Funcionário")

    ' Cancelar ou vazio
    If escolha = "" Then
        MsgBox "Impressão cancelada.", vbExclamation
        Exit Sub
    End If

    ' ===== OPÇÃO IMPRIMIR TODOS =====
    If UCase(escolha) = "TODOS" Then
        For Each ws In ThisWorkbook.Sheets
            If ws.Name <> "Gerar Pontos" And _
               ws.Name <> "Molde" And _
               ws.Name <> "Referências" And _
               ws.Name <> "Feriados" And _
               LCase(ws.Name) <> "estag ianara" Then
               
                ws.PrintOut Copies:=1, Collate:=True
            End If
        Next ws

        MsgBox "Impressão concluída para TODOS os funcionários.", vbInformation
        Exit Sub
    End If

    ' ===== IMPRIMIR APENAS UM =====
    encontrou = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = escolha Then
            encontrou = True

            ' Segurança extra
            If ws.Name = "Molde" Or ws.Name = "Gerar Pontos" Or _
               ws.Name = "Feriados" Or ws.Name = "Referências" Or _
               LCase(ws.Name) = "estag ianara" Then
               
                MsgBox "Esta planilha não pode ser impressa.", vbExclamation
                Exit Sub
            End If

            ws.PrintOut Copies:=1, Collate:=True
            MsgBox "Impressão concluída para: " & ws.Name, vbInformation
            Exit Sub
        End If
    Next ws

    If Not encontrou Then
        MsgBox "Funcionário não encontrado. Verifique o nome e tente novamente.", vbCritical
    End If
End Sub

