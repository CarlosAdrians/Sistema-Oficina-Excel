Attribute VB_Name = "modVeiculos"
Option Explicit

Function BuscarVeiculoPorID(idVeiculo As Long) As ListRow
    Dim row As ListRow
    For Each row In shVeiculos.ListObjects("tbVeiculos").ListRows
        If row.Range.Cells(1, 1).Value = idVeiculo Then
            Set BuscarVeiculoPorID = row
            Exit Function
        End If
    Next row
    Set BuscarVeiculoPorID = Nothing
End Function


Function BuscarVeiculoPorPlaca(placa As String) As ListRow
    Dim row As ListRow
    placa = UCase(Replace(placa, "-", ""))

    For Each row In shVeiculos.ListObjects("tbVeiculos").ListRows
        If UCase(Replace(CStr(row.Range.Cells(1, 7).Value), "-", "")) = placa Then
            Set BuscarVeiculoPorPlaca = row
            Exit Function
        End If
    Next row
    Set BuscarVeiculoPorPlaca = Nothing
End Function


Function PlacaJaExiste(placa As String, idIgnorar As Long) As Boolean
    Dim row As ListRow
    placa = UCase(Replace(placa, "-", ""))

    For Each row In shVeiculos.ListObjects("tbVeiculos").ListRows
        If row.Range.Cells(1, 1).Value <> idIgnorar Then
            If UCase(Replace(CStr(row.Range.Cells(1, 7).Value), "-", "")) = placa Then
                PlacaJaExiste = True
                Exit Function
            End If
        End If
    Next row
    PlacaJaExiste = False
End Function


Function NovoVeiculo(idCliente As Long, marca As String, modelo As String, _
                     ano As Integer, cor As String, placa As String) As Boolean

    If BuscarClientePorID(idCliente) Is Nothing Then
        MsgBox "Cliente não encontrado.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If Not ValidarMarca(marca) Then
        MsgBox "Marca inválida.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If Not ValidarModelo(modelo) Then
        MsgBox "Modelo inválido.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If Not ValidarAnoVeiculo(ano) Then
        MsgBox "Ano inválido.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If Not ValidarCor(cor) Then
        MsgBox "Cor inválida.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If Not ValidarPlaca(placa) Then
        MsgBox "Placa inválida.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    If PlacaJaExiste(placa, 0) Then
        MsgBox "Placa já cadastrada.", vbExclamation, "Atenção"
        NovoVeiculo = False
        Exit Function
    End If

    Dim tb As ListObject
    Set tb = shVeiculos.ListObjects("tbVeiculos")

    Dim novaLinha As ListRow
    Set novaLinha = tb.ListRows.Add

    With novaLinha.Range
        .Cells(1, 1).Value = GerarID(tb, "ID_VEICULO")
        .Cells(1, 2).Value = idCliente
        .Cells(1, 3).Value = Trim(marca)
        .Cells(1, 4).Value = Trim(modelo)
        .Cells(1, 5).Value = ano
        .Cells(1, 6).Value = Trim(cor)
        .Cells(1, 7).Value = FormatarPlaca(placa)
    End With

    NovoVeiculo = True
End Function


Function EditarVeiculo(placa As String, marca As String, modelo As String, _
                       ano As Integer, cor As String) As Boolean

    Dim row As ListRow
    Set row = BuscarVeiculoPorPlaca(placa)

    If row Is Nothing Then
        MsgBox "Veículo não encontrado.", vbExclamation, "Atenção"
        EditarVeiculo = False
        Exit Function
    End If

    If Not ValidarMarca(marca) Then
        MsgBox "Marca inválida.", vbExclamation, "Atenção"
        EditarVeiculo = False
        Exit Function
    End If

    If Not ValidarModelo(modelo) Then
        MsgBox "Modelo inválido.", vbExclamation, "Atenção"
        EditarVeiculo = False
        Exit Function
    End If

    If Not ValidarAnoVeiculo(ano) Then
        MsgBox "Ano inválido.", vbExclamation, "Atenção"
        EditarVeiculo = False
        Exit Function
    End If

    If Not ValidarCor(cor) Then
        MsgBox "Cor inválida.", vbExclamation, "Atenção"
        EditarVeiculo = False
        Exit Function
    End If

    With row.Range
        .Cells(1, 3).Value = Trim(marca)
        .Cells(1, 4).Value = Trim(modelo)
        .Cells(1, 5).Value = ano
        .Cells(1, 6).Value = Trim(cor)
    End With

    EditarVeiculo = True
End Function


Function ExcluirVeiculo(placa As String) As Boolean
    Dim row As ListRow
    Set row = BuscarVeiculoPorPlaca(placa)

    If row Is Nothing Then
        MsgBox "Veículo não encontrado.", vbExclamation, "Atenção"
        ExcluirVeiculo = False
        Exit Function
    End If

    Dim idVeiculo As Long
    idVeiculo = row.Range.Cells(1, 1).Value

    Dim rowOS As ListRow
    For Each rowOS In shOs.ListObjects("tbOs").ListRows
        If rowOS.Range.Cells(1, 4).Value = idVeiculo Then
            MsgBox "Veículo possui OS vinculada e não pode ser excluído.", vbExclamation, "Atenção"
            ExcluirVeiculo = False
            Exit Function
        End If
    Next rowOS

    row.Delete
    ExcluirVeiculo = True
End Function
