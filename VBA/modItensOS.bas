Attribute VB_Name = "modItensOS"
Option Explicit

Function BuscarItemPorID(idItem As Long) As ListRow
    Dim row As ListRow
    For Each row In shItensOS.ListObjects("tbItensOs").ListRows
        If row.Range.Cells(1, 1).Value = idItem Then
            Set BuscarItemPorID = row
            Exit Function
        End If
    Next row
    Set BuscarItemPorID = Nothing
End Function


Function NovoItem(idOS As Long, tipo As String, qtd As Double, _
                  descricao As String, valorUnit As Double) As Boolean

    If BuscarOSPorID(idOS) Is Nothing Then
        MsgBox "OS não encontrada.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    ' Bloqueia adicao de item em OS concluida ou cancelada
    Dim rowOS As ListRow
    Set rowOS = BuscarOSPorID(idOS)
    If rowOS.Range.Cells(1, 8).Value = "Concluída" Or _
       rowOS.Range.Cells(1, 8).Value = "Cancelada" Then
        MsgBox "Não é possível adicionar itens em OS concluída ou cancelada.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    If Not ValidarTipoItem(tipo) Then
        MsgBox "Tipo inválido.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    If Not ValidarQtd(qtd) Then
        MsgBox "Quantidade inválida.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    If Not ValidarDescricaoItem(descricao) Then
        MsgBox "Descrição inválida.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    If Not ValidarValorUnit(valorUnit) Then
        MsgBox "Valor unitário inválido.", vbExclamation, "Atenção"
        NovoItem = False
        Exit Function
    End If

    Dim tb As ListObject
    Set tb = shItensOS.ListObjects("tbItensOs")

    Dim novaLinha As ListRow
    Set novaLinha = tb.ListRows.Add

    With novaLinha.Range
        .Cells(1, 1).Value = GerarID(tb, "ID_ITEM")
        .Cells(1, 2).Value = idOS
        .Cells(1, 3).Value = Trim(tipo)
        .Cells(1, 4).Value = qtd
        .Cells(1, 5).Value = Trim(descricao)
        .Cells(1, 6).Value = valorUnit
        .Cells(1, 7).Value = CalcularTotalItem(qtd, valorUnit)
    End With

    AtualizarValorTotalOS idOS
    NovoItem = True
End Function


Function EditarItem(idItem As Long, tipo As String, qtd As Double, _
                    descricao As String, valorUnit As Double) As Boolean

    Dim row As ListRow
    Set row = BuscarItemPorID(idItem)

    If row Is Nothing Then
        MsgBox "Item não encontrado.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    ' Bloqueia edicao de item em OS concluida ou cancelada
    Dim idOS As Long
    idOS = row.Range.Cells(1, 2).Value
    Dim rowOS As ListRow
    Set rowOS = BuscarOSPorID(idOS)
    If rowOS.Range.Cells(1, 8).Value = "Concluída" Or _
       rowOS.Range.Cells(1, 8).Value = "Cancelada" Then
        MsgBox "Não é possível editar itens em OS concluída ou cancelada.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    If Not ValidarTipoItem(tipo) Then
        MsgBox "Tipo inválido.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    If Not ValidarQtd(qtd) Then
        MsgBox "Quantidade inválida.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    If Not ValidarDescricaoItem(descricao) Then
        MsgBox "Descrição inválida.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    If Not ValidarValorUnit(valorUnit) Then
        MsgBox "Valor unitário inválido.", vbExclamation, "Atenção"
        EditarItem = False
        Exit Function
    End If

    With row.Range
        .Cells(1, 3).Value = Trim(tipo)
        .Cells(1, 4).Value = qtd
        .Cells(1, 5).Value = Trim(descricao)
        .Cells(1, 6).Value = valorUnit
        .Cells(1, 7).Value = CalcularTotalItem(qtd, valorUnit)
    End With

    AtualizarValorTotalOS idOS
    EditarItem = True
End Function


Function ExcluirItem(idItem As Long) As Boolean
    Dim row As ListRow
    Set row = BuscarItemPorID(idItem)

    If row Is Nothing Then
        MsgBox "Item não encontrado.", vbExclamation, "Atenção"
        ExcluirItem = False
        Exit Function
    End If

    ' Bloqueia exclusao de item em OS concluida ou cancelada
    Dim idOS As Long
    idOS = row.Range.Cells(1, 2).Value
    Dim rowOS As ListRow
    Set rowOS = BuscarOSPorID(idOS)
    If rowOS.Range.Cells(1, 8).Value = "Concluída" Or _
       rowOS.Range.Cells(1, 8).Value = "Cancelada" Then
        MsgBox "Não é possível excluir itens em OS concluída ou cancelada.", vbExclamation, "Atenção"
        ExcluirItem = False
        Exit Function
    End If

    row.Delete
    AtualizarValorTotalOS idOS
    ExcluirItem = True
End Function

