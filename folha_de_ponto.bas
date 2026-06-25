Attribute VB_Name = "Módulo1"
Option Explicit

' =====================================================
' Converte nome do mês em número
' =====================================================
Public Function getMonth(selectedMonth As String) As Integer
    Dim dt As Date
    On Error GoTo ErrHandler
    dt = DateValue("1 " & selectedMonth & " 2000")
    getMonth = month(dt)
    Exit Function
ErrHandler:
    getMonth = -1
End Function

' =====================================================
' Lê os feriados da planilha "Feriados"
' =====================================================
Private Function getHolidays(sheetName As String) As Collection
    Dim ws As Worksheet
    Dim holidays As New Collection
    Dim lastRow As Long, i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "A planilha '" & sheetName & "' não foi encontrada. Ela será criada vazia.", vbExclamation
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
        ws.Range("A1").Value = "Datas de Feriados"
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 1).Value) Then
            holidays.Add CDate(ws.Cells(i, 1).Value)
        End If
    Next i

    Set getHolidays = holidays
End Function

' =====================================================
' Atualiza todas as folhas ou uma específica
' =====================================================
Sub atualizarTodasAsFolhas()
    Dim ws As Worksheet
    Dim selectedYear As Integer
    Dim selectedMonth As Integer
    Dim holidays As Collection
    Dim listaFuncionarios As String
    Dim escolha As Variant

    ' Monta lista de planilhas (exclui as que não devem aparecer)
listaFuncionarios = ""
    For Each ws In ThisWorkbook.Sheets
       If ws.Name <> "Menu" And ws.Name <> "Feriados" And ws.Name <> "Referencias" _
       And ws.Name <> "Molde" And LCase(ws.Name) <> "estag ianara" Then
        listaFuncionarios = listaFuncionarios & ws.Name & vbCrLf
    End If
Next ws

    ' Escolha do funcionário ou TODOS
    escolha = InputBox("Digite o nome de um funcionário ou digite 'TODOS' para gerar todas:" & vbCrLf & vbCrLf & listaFuncionarios, "Escolher Funcionário")

    If escolha = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    ' Lê mês e ano do Menu
    selectedYear = CInt(Sheets("Menu").Range("B3").Value)
    selectedMonth = getMonth(Sheets("Menu").Range("C3").Value)

    Set holidays = getHolidays("Feriados")

    Application.ScreenUpdating = False

    ' === Se for TODOS, atualiza todas as planilhas ===
    If UCase(escolha) = "TODOS" Then
        For Each ws In ThisWorkbook.Sheets
            If ws.Name <> "Menu" And ws.Name <> "Feriados" And ws.Name <> "Referencias" And ws.Name <> "Molde" And ws.Name <> "estag IANARA" Then
                ws.Range("A16:J47").ClearContents
                ws.Range("A16:J47").Interior.ColorIndex = xlNone
                fillMonthSheet ws, selectedYear, selectedMonth, holidays
                ws.Range("I10").Value = Format(DateSerial(selectedYear, selectedMonth, 1), "mm/yyyy")
            End If
        Next ws
        MsgBox "Todas as folhas foram geradas para " & Format(DateSerial(selectedYear, selectedMonth, 1), "mm/yyyy"), vbInformation

    ' === Caso contrário, apenas o funcionário escolhido ===
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(escolha)
        On Error GoTo 0

        If ws Is Nothing Then
            MsgBox "Funcionário não encontrado. Verifique o nome e tente novamente.", vbCritical
            Exit Sub
        End If

        ws.Range("A16:J47").ClearContents
        ws.Range("A16:J47").Interior.ColorIndex = xlNone

        fillMonthSheet ws, selectedYear, selectedMonth, holidays
        ws.Range("I10").Value = Format(DateSerial(selectedYear, selectedMonth, 1), "mm/yyyy")

        MsgBox "Folha de ponto atualizada para: " & ws.Name, vbInformation
    End If

    Application.ScreenUpdating = True
End Sub

' =====================================================
' Preenche o calendário do mês (prioriza fins de semana sobre feriados)
' =====================================================
Private Sub fillMonthSheet(ws As Worksheet, year As Integer, month As Integer, holidays As Collection)
    Dim rowStart As Long: rowStart = 16
    Dim dayCounter As Integer
    Dim currentDate As Date
    Dim isWeekend As Boolean, isHoliday As Boolean
    Dim colorRange As Range, h As Variant
    Dim dayLabel As String

    For dayCounter = 1 To Day(DateSerial(year, month + 1, 0))
        currentDate = DateSerial(year, month, dayCounter)
        ws.Cells(rowStart + dayCounter - 1, 1).Value = dayCounter

        ' Verifica fim de semana
        isWeekend = (Weekday(currentDate, vbMonday) >= 6)
        
        ' Verifica se é feriado
        isHoliday = False
        For Each h In holidays
            If h = currentDate Then
                isHoliday = True
                Exit For
            End If
        Next h

        ' === PRIORIDADE: fim de semana > feriado ===
        If isWeekend Then
            If Weekday(currentDate, vbMonday) = 6 Then
                dayLabel = "SÁBADO"
            Else
                dayLabel = "DOMINGO"
            End If
        ElseIf isHoliday Then
            dayLabel = "FERIADO"
        Else
            dayLabel = ""
        End If

        ' === Se for fim de semana ou feriado, pinta ===
        If isWeekend Or isHoliday Then
            Set colorRange = ws.Range(ws.Cells(rowStart + dayCounter - 1, 1), ws.Cells(rowStart + dayCounter - 1, 10))
            colorRange.Interior.Color = RGB(200, 200, 200)
            ws.Cells(rowStart + dayCounter - 1, 2).Value = dayLabel
        End If
    Next dayCounter
End Sub

