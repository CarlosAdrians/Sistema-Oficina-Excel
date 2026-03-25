Attribute VB_Name = "modOS"
Option Explicit

Function BuscarOSPorID(idOS As Long) As ListRow
    Dim row As ListRow
    For Each row In shOs.ListObjects("tbOs").ListRows
        If row.Range.Cells(1, 1).Value = idOS Then
            Set BuscarOSPorID = row
            Exit Function
        End If
    Next row
    Set BuscarOSPorID = Nothing
End Function


Function NovaOS(idCliente As Long, idVeiculo As Long, observacoes As String) As Long
    If BuscarClientePorID(idCliente) Is Nothing Then
        MsgBox "Cliente não encontrado.", vbExclamation, "Atenção"
        NovaOS = 0
        Exit Function
    End If

    If BuscarVeiculoPorID(idVeiculo) Is Nothing Then
        MsgBox "Veículo não encontrado.", vbExclamation, "Atenção"
        NovaOS = 0
        Exit Function
    End If

    ' Verifica se veiculo pertence ao cliente
    Dim rowVeiculo As ListRow
    Set rowVeiculo = BuscarVeiculoPorID(idVeiculo)
    If rowVeiculo.Range.Cells(1, 2).Value <> idCliente Then
        MsgBox "Veículo não pertence ao cliente.", vbExclamation, "Atenção"
        NovaOS = 0
        Exit Function
    End If

    Dim tb As ListObject
    Set tb = shOs.ListObjects("tbOs")

    Dim novaLinha As ListRow
    Set novaLinha = tb.ListRows.Add

    Dim novoID As Long
    novoID = GerarID(tb, "ID_OS")

    With novaLinha.Range
        .Cells(1, 1).Value = novoID
        .Cells(1, 2).Value = Format(Now, "DD/MM/YYYY")
        .Cells(1, 3).Value = ""
        .Cells(1, 4).Value = idCliente
        .Cells(1, 5).Value = idVeiculo
        .Cells(1, 6).Value = Trim(observacoes)
        .Cells(1, 7).Value = 0
        .Cells(1, 8).Value = "Aberta"
    End With

    NovaOS = novoID
End Function


Function EditarOS(idOS As Long, observacoes As String, status As String) As Boolean
    Dim row As ListRow
    Set row = BuscarOSPorID(idOS)

    If row Is Nothing Then
        MsgBox "OS não encontrada.", vbExclamation, "Atenção"
        EditarOS = False
        Exit Function
    End If

    If Not ValidarStatusOS(status) Then
        MsgBox "Status inválido.", vbExclamation, "Atenção"
        EditarOS = False
        Exit Function
    End If

    ' Bloqueia edicao de OS concluida ou cancelada
    Dim statusAtual As String
    statusAtual = row.Range.Cells(1, 8).Value
    If statusAtual = "Concluída" Or statusAtual = "Cancelada" Then
        MsgBox "OS concluída ou cancelada não pode ser editada.", vbExclamation, "Atenção"
        EditarOS = False
        Exit Function
    End If

    row.Range.Cells(1, 6).Value = Trim(observacoes)
    row.Range.Cells(1, 8).Value = status

    ' Preenche data fechamento automaticamente
    If status = "Concluída" Or status = "Cancelada" Then
        row.Range.Cells(1, 3).Value = Format(Now, "DD/MM/YYYY")
    End If

    EditarOS = True
End Function


Function AtualizarValorTotalOS(idOS As Long) As Boolean
    Dim rowOS As ListRow
    Set rowOS = BuscarOSPorID(idOS)

    If rowOS Is Nothing Then
        AtualizarValorTotalOS = False
        Exit Function
    End If

    Dim total As Double
    total = 0

    Dim row As ListRow
    For Each row In shItensOS.ListObjects("tbItensOs").ListRows
        If row.Range.Cells(1, 2).Value = idOS Then
            total = total + row.Range.Cells(1, 7).Value
        End If
    Next row

    rowOS.Range.Cells(1, 7).Value = total
    AtualizarValorTotalOS = True
End Function


Function ExcluirOS(idOS As Long) As Boolean
    Dim row As ListRow
    Set row = BuscarOSPorID(idOS)

    If row Is Nothing Then
        MsgBox "OS não encontrada.", vbExclamation, "Atenção"
        ExcluirOS = False
        Exit Function
    End If

    ' Bloqueia exclusao de OS concluida
    If row.Range.Cells(1, 8).Value = "Concluída" Then
        MsgBox "OS concluída não pode ser excluída.", vbExclamation, "Atenção"
        ExcluirOS = False
        Exit Function
    End If

    ' Verifica se tem pagamento vinculado
    Dim rowPag As ListRow
    For Each rowPag In shPagamentos.ListObjects("tbPagamentos").ListRows
        If rowPag.Range.Cells(1, 2).Value = idOS Then
            MsgBox "OS possui pagamento vinculado e não pode ser excluída.", vbExclamation, "Atenção"
            ExcluirOS = False
            Exit Function
        End If
    Next rowPag

    ' Exclui itens vinculados antes de excluir a OS
    Dim rowItem As ListRow
    Dim i As Integer
    For i = shItensOS.ListObjects("tbItensOs").ListRows.Count To 1 Step -1
        Set rowItem = shItensOS.ListObjects("tbItensOs").ListRows(i)
        If rowItem.Range.Cells(1, 2).Value = idOS Then
            rowItem.Delete
        End If
    Next i

    row.Delete
    ExcluirOS = True
End Function
