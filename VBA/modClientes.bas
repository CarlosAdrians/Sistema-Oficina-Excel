Attribute VB_Name = "modClientes"
Option Explicit

Function BuscarClientePorID(idCliente As Long) As ListRow
    Dim row As ListRow
    For Each row In shClientes.ListObjects("tbClientes").ListRows
        If row.Range.Cells(1, 1).Value = idCliente Then
            Set BuscarClientePorID = row
            Exit Function
        End If
    Next row
    Set BuscarClientePorID = Nothing
End Function


Function BuscarClientePorDocumento(documento As String) As ListRow
    Dim row As ListRow
    documento = ApenasNumeros(documento)

    For Each row In shClientes.ListObjects("tbClientes").ListRows
        If ApenasNumeros(CStr(row.Range.Cells(1, 4).Value)) = documento Or _
           ApenasNumeros(CStr(row.Range.Cells(1, 5).Value)) = documento Then
            Set BuscarClientePorDocumento = row
            Exit Function
        End If
    Next row
    Set BuscarClientePorDocumento = Nothing
End Function


Function DocumentoJaExiste(documento As String, idIgnorar As Long) As Boolean
    Dim row As ListRow
    documento = ApenasNumeros(documento)

    For Each row In shClientes.ListObjects("tbClientes").ListRows
        If row.Range.Cells(1, 1).Value <> idIgnorar Then
            If ApenasNumeros(CStr(row.Range.Cells(1, 4).Value)) = documento Or _
               ApenasNumeros(CStr(row.Range.Cells(1, 5).Value)) = documento Then
                DocumentoJaExiste = True
                Exit Function
            End If
        End If
    Next row
    DocumentoJaExiste = False
End Function


Function NovoCliente(nome As String, tipoPessoa As String, cpf As String, _
                     cnpj As String, telefone As String, endereco As String) As Boolean

    If Not ValidarNome(nome) Then
        MsgBox "Nome inválido.", vbExclamation, "Atenção"
        NovoCliente = False
        Exit Function
    End If

    If tipoPessoa = "PF" Then
        If Not ValidarCPF(cpf) Then
            MsgBox "CPF inválido.", vbExclamation, "Atenção"
            NovoCliente = False
            Exit Function
        End If
        If DocumentoJaExiste(cpf, 0) Then
            MsgBox "CPF já cadastrado.", vbExclamation, "Atenção"
            NovoCliente = False
            Exit Function
        End If
        cnpj = ""
    ElseIf tipoPessoa = "PJ" Then
        If Not ValidarCNPJ(cnpj) Then
            MsgBox "CNPJ inválido.", vbExclamation, "Atenção"
            NovoCliente = False
            Exit Function
        End If
        If DocumentoJaExiste(cnpj, 0) Then
            MsgBox "CNPJ já cadastrado.", vbExclamation, "Atenção"
            NovoCliente = False
            Exit Function
        End If
        cpf = ""
    Else
        MsgBox "Tipo de pessoa inválido.", vbExclamation, "Atenção"
        NovoCliente = False
        Exit Function
    End If

    If Not ValidarEndereco(endereco) Then
        MsgBox "Endereço inválido.", vbExclamation, "Atenção"
        NovoCliente = False
        Exit Function
    End If

    Dim tb As ListObject
    Set tb = shClientes.ListObjects("tbClientes")

    Dim novaLinha As ListRow
    Set novaLinha = tb.ListRows.Add

    With novaLinha.Range
        .Cells(1, 1).Value = GerarID(tb, "ID_CLIENTE")
        .Cells(1, 2).Value = Trim(nome)
        .Cells(1, 3).Value = tipoPessoa
        .Cells(1, 4).Value = IIf(tipoPessoa = "PF", FormatarCPF(cpf), "")
        .Cells(1, 5).Value = IIf(tipoPessoa = "PJ", FormatarCNPJ(cnpj), "")
        .Cells(1, 6).Value = FormatarTelefone(telefone)
        .Cells(1, 7).Value = Trim(endereco)
        .Cells(1, 8).Value = DataCadastroAtual()
    End With

    NovoCliente = True
End Function


Function EditarCliente(documento As String, nome As String, _
                       telefone As String, endereco As String) As Boolean

    Dim row As ListRow
    Set row = BuscarClientePorDocumento(documento)

    If row Is Nothing Then
        MsgBox "Cliente não encontrado.", vbExclamation, "Atenção"
        EditarCliente = False
        Exit Function
    End If

    If Not ValidarNome(nome) Then
        MsgBox "Nome inválido.", vbExclamation, "Atenção"
        EditarCliente = False
        Exit Function
    End If

    If Not ValidarEndereco(endereco) Then
        MsgBox "Endereço inválido.", vbExclamation, "Atenção"
        EditarCliente = False
        Exit Function
    End If

    With row.Range
        .Cells(1, 2).Value = Trim(nome)
        .Cells(1, 6).Value = FormatarTelefone(telefone)
        .Cells(1, 7).Value = Trim(endereco)
    End With

    EditarCliente = True
End Function


Function ExcluirCliente(documento As String) As Boolean
    Dim row As ListRow
    Set row = BuscarClientePorDocumento(documento)

    If row Is Nothing Then
        MsgBox "Cliente não encontrado.", vbExclamation, "Atenção"
        ExcluirCliente = False
        Exit Function
    End If

    Dim idCliente As Long
    idCliente = row.Range.Cells(1, 1).Value

    Dim rowOS As ListRow
    For Each rowOS In shOs.ListObjects("tbOs").ListRows
        If rowOS.Range.Cells(1, 3).Value = idCliente Then
            MsgBox "Cliente possui OS vinculada e não pode ser excluído.", vbExclamation, "Atenção"
            ExcluirCliente = False
            Exit Function
        End If
    Next rowOS

    Dim rowVeiculo As ListRow
    For Each rowVeiculo In shVeiculos.ListObjects("tbVeiculos").ListRows
        If rowVeiculo.Range.Cells(1, 2).Value = idCliente Then
            MsgBox "Cliente possui veículo vinculado e não pode ser excluído.", vbExclamation, "Atenção"
            ExcluirCliente = False
            Exit Function
        End If
    Next rowVeiculo

    row.Delete
    ExcluirCliente = True
End Function

