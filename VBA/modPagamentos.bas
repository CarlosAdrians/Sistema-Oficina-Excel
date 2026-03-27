Attribute VB_Name = "modPagamentos"
Option Explicit

Function BuscarPagamentoPorID(idPagamento As Long) As ListRow
    Dim row As ListRow
    For Each row In shPagamentos.ListObjects("tbPagamentos").ListRows
        If row.Range.Cells(1, 1).Value = idPagamento Then
            Set BuscarPagamentoPorID = row
            Exit Function
        End If
    Next row
    Set BuscarPagamentoPorID = Nothing
End Function


Function BuscarPagamentoPorOS(idOS As Long) As ListRow
    Dim row As ListRow
    For Each row In shPagamentos.ListObjects("tbPagamentos").ListRows
        If row.Range.Cells(1, 2).Value = idOS Then
            Set BuscarPagamentoPorOS = row
            Exit Function
        End If
    Next row
    Set BuscarPagamentoPorOS = Nothing
End Function


Function NovoPagamento(idOS As Long, adiantamento As Double, forma As String, parcelas As Integer) As Boolean
    Dim rowOS As ListRow
    Set rowOS = BuscarOSPorID(idOS)

    If rowOS Is Nothing Then
        MsgBox "OS não encontrada.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    ' Bloqueia pagamento em OS cancelada
    If rowOS.Range.Cells(1, 8).Value = "Cancelada" Then
        MsgBox "Não é possível registrar pagamento em OS cancelada.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    ' Verifica se ja existe pagamento para essa OS
    If Not BuscarPagamentoPorOS(idOS) Is Nothing Then
        MsgBox "Já existe pagamento registrado para essa OS.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    Dim total As Double
    total = rowOS.Range.Cells(1, 7).Value

    If total = 0 Then
        MsgBox "OS não possui itens registrados.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    If Not ValidarAdiantamento(adiantamento, total) Then
        MsgBox "Adiantamento inválido.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    If Not ValidarForma(forma) Then
        MsgBox "Forma de pagamento inválida.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    If Not ValidarParcelas(parcelas, forma) Then
        MsgBox "Número de parcelas inválido.", vbExclamation, "Atenção"
        NovoPagamento = False
        Exit Function
    End If

    Dim tb As ListObject
    Set tb = shPagamentos.ListObjects("tbPagamentos")

    Dim novaLinha As ListRow
    Set novaLinha = tb.ListRows.Add

    With novaLinha.Range
        .Cells(1, 1).Value = GerarID(tb, "ID_PAGAMENTO")
        .Cells(1, 2).Value = idOS
        .Cells(1, 3).Value = total
        .Cells(1, 4).Value = adiantamento
        .Cells(1, 5).Value = CalcularRestante(total, adiantamento)
        .Cells(1, 6).Value = Format(Now, "DD/MM/YYYY")
        .Cells(1, 7).Value = Trim(forma)
        .Cells(1, 8).Value = IIf(forma = "Crédito", parcelas, "")
        .Cells(1, 9).Value = CalcularStatusPagamento(adiantamento, total)
    End With

    NovoPagamento = True
End Function


Function EditarPagamento(idPagamento As Long, adiantamento As Double, forma As String, parcelas As Integer) As Boolean
    Dim row As ListRow
    Set row = BuscarPagamentoPorID(idPagamento)

    If row Is Nothing Then
        MsgBox "Pagamento não encontrado.", vbExclamation, "Atenção"
        EditarPagamento = False
        Exit Function
    End If

    ' Bloqueia edicao de pagamento quitado
    If row.Range.Cells(1, 9).Value = "Quitado" Then
        MsgBox "Pagamento quitado não pode ser editado.", vbExclamation, "Atenção"
        EditarPagamento = False
        Exit Function
    End If

    Dim total As Double
    total = row.Range.Cells(1, 3).Value

    If Not ValidarAdiantamento(adiantamento, total) Then
        MsgBox "Adiantamento inválido.", vbExclamation, "Atenção"
        EditarPagamento = False
        Exit Function
    End If

    If Not ValidarForma(forma) Then
        MsgBox "Forma de pagamento inválida.", vbExclamation, "Atenção"
        EditarPagamento = False
        Exit Function
    End If

    If Not ValidarParcelas(parcelas, forma) Then
        MsgBox "Número de parcelas inválido.", vbExclamation, "Atenção"
        EditarPagamento = False
        Exit Function
    End If

    With row.Range
        .Cells(1, 4).Value = adiantamento
        .Cells(1, 5).Value = CalcularRestante(total, adiantamento)
        .Cells(1, 7).Value = Trim(forma)
        .Cells(1, 8).Value = IIf(forma = "Crédito", parcelas, "")
        .Cells(1, 9).Value = CalcularStatusPagamento(adiantamento, total)
    End With

    EditarPagamento = True
End Function


Function ExcluirPagamento(idPagamento As Long) As Boolean
    Dim row As ListRow
    Set row = BuscarPagamentoPorID(idPagamento)

    If row Is Nothing Then
        MsgBox "Pagamento não encontrado.", vbExclamation, "Atenção"
        ExcluirPagamento = False
        Exit Function
    End If

    ' Bloqueia exclusao de pagamento quitado
    If row.Range.Cells(1, 9).Value = "Quitado" Then
        MsgBox "Pagamento quitado não pode ser excluído.", vbExclamation, "Atenção"
        ExcluirPagamento = False
        Exit Function
    End If

    row.Delete
    ExcluirPagamento = True
End Function
