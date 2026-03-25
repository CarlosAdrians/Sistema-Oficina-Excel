Attribute VB_Name = "modUtilitarios"
Option Explicit

'Para todos

Function GerarID(tabela As ListObject, nomeColuna As String) As Long
    If tabela.ListRows.Count = 0 Then
        GerarID = 1
        Exit Function
    End If
    
    Dim colIndex As Long
    colIndex = tabela.ListColumns(nomeColuna).Index
    GerarID = Application.WorksheetFunction.Max(tabela.ListColumns(nomeColuna).DataBodyRange) + 1
End Function

'Clientes

Function ValidarCPF(cpf As String) As Boolean
    cpf = ApenasNumeros(cpf)

    If Len(cpf) <> 11 Then
        ValidarCPF = False
        Exit Function
    End If

    If cpf = String(11, Left(cpf, 1)) Then
        ValidarCPF = False
        Exit Function
    End If

    Dim i As Integer, soma As Long, resto As Long

    soma = 0
    For i = 1 To 9
        soma = soma + CInt(Mid(cpf, i, 1)) * (11 - i)
    Next i
    resto = (soma * 10) Mod 11
    If resto = 10 Or resto = 11 Then resto = 0
    If resto <> CInt(Mid(cpf, 10, 1)) Then
        ValidarCPF = False
        Exit Function
    End If

    soma = 0
    For i = 1 To 10
        soma = soma + CInt(Mid(cpf, i, 1)) * (12 - i)
    Next i
    resto = (soma * 10) Mod 11
    If resto = 10 Or resto = 11 Then resto = 0
    If resto <> CInt(Mid(cpf, 11, 1)) Then
        ValidarCPF = False
        Exit Function
    End If

    ValidarCPF = True
End Function


Function ValidarCNPJ(cnpj As String) As Boolean
    cnpj = ApenasNumeros(cnpj)

    If Len(cnpj) <> 14 Then
        ValidarCNPJ = False
        Exit Function
    End If

    If cnpj = String(14, Left(cnpj, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If

    Dim pesos1 As Variant
    Dim pesos2 As Variant
    pesos1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    pesos2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)

    Dim i As Integer, soma As Long, resto As Long

    soma = 0
    For i = 0 To 11
        soma = soma + CInt(Mid(cnpj, i + 1, 1)) * pesos1(i)
    Next i
    resto = soma Mod 11
    If resto < 2 Then resto = 0 Else resto = 11 - resto
    If resto <> CInt(Mid(cnpj, 13, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If

    soma = 0
    For i = 0 To 12
        soma = soma + CInt(Mid(cnpj, i + 1, 1)) * pesos2(i)
    Next i
    resto = soma Mod 11
    If resto < 2 Then resto = 0 Else resto = 11 - resto
    If resto <> CInt(Mid(cnpj, 14, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If

    ValidarCNPJ = True
End Function

Function ValidarNome(nome As String) As Boolean
    ValidarNome = (Len(Trim(nome)) >= 3)
End Function

Function ValidarEndereco(endereco As String) As Boolean
    ValidarEndereco = (Len(Trim(endereco)) >= 5)
End Function

Function DataCadastroAtual() As String
    DataCadastroAtual = Format(Now, "DD/MM/YYYY")
End Function

Function ApenasNumeros(texto As String) As String
    Dim resultado As String
    Dim i As Integer
    resultado = ""
    For i = 1 To Len(texto)
        If Mid(texto, i, 1) >= "0" And Mid(texto, i, 1) <= "9" Then
            resultado = resultado & Mid(texto, i, 1)
        End If
    Next i
    ApenasNumeros = resultado
End Function

Function FormatarCPF(cpf As String) As String
    cpf = ApenasNumeros(cpf)
    If Len(cpf) <> 11 Then
        FormatarCPF = cpf
        Exit Function
    End If
    FormatarCPF = Left(cpf, 3) & "." & Mid(cpf, 4, 3) & "." & Mid(cpf, 7, 3) & "-" & Right(cpf, 2)
End Function

Function FormatarCNPJ(cnpj As String) As String
    cnpj = ApenasNumeros(cnpj)
    If Len(cnpj) <> 14 Then
        FormatarCNPJ = cnpj
        Exit Function
    End If
    FormatarCNPJ = Left(cnpj, 2) & "." & Mid(cnpj, 3, 3) & "." & Mid(cnpj, 6, 3) & "/" & Mid(cnpj, 9, 4) & "-" & Right(cnpj, 2)
End Function

Function FormatarTelefone(tel As String) As String
    tel = ApenasNumeros(tel)
    Select Case Len(tel)
        Case 11
            FormatarTelefone = "(" & Left(tel, 2) & ") " & Mid(tel, 3, 5) & "-" & Right(tel, 4)
        Case 10
            FormatarTelefone = "(" & Left(tel, 2) & ") " & Mid(tel, 3, 4) & "-" & Right(tel, 4)
        Case Else
            FormatarTelefone = tel
    End Select
End Function


'Veiculos

Function ValidarPlaca(placa As String) As Boolean
    placa = UCase(Replace(placa, "-", ""))

    If Len(placa) <> 7 Then
        ValidarPlaca = False
        Exit Function
    End If

    Dim i As Integer
    Dim c As String

    For i = 1 To 3
        c = Mid(placa, i, 1)
        If c < "A" Or c > "Z" Then
            ValidarPlaca = False
            Exit Function
        End If
    Next i

    c = Mid(placa, 4, 1)
    If c < "0" Or c > "9" Then
        ValidarPlaca = False
        Exit Function
    End If

    c = Mid(placa, 5, 1)
    Dim isMercosul As Boolean
    isMercosul = (c >= "A" And c <= "Z")
    If Not isMercosul Then
        If c < "0" Or c > "9" Then
            ValidarPlaca = False
            Exit Function
        End If
    End If

    For i = 6 To 7
        c = Mid(placa, i, 1)
        If c < "0" Or c > "9" Then
            ValidarPlaca = False
            Exit Function
        End If
    Next i

    ValidarPlaca = True
End Function


Function FormatarPlaca(placa As String) As String
    placa = UCase(Replace(placa, "-", ""))
    If Len(placa) <> 7 Then
        FormatarPlaca = placa
        Exit Function
    End If
    FormatarPlaca = Left(placa, 3) & "-" & Right(placa, 4)
End Function


Function ValidarAnoVeiculo(ano As Integer) As Boolean
    ValidarAnoVeiculo = (ano >= 1950 And ano <= Year(Now) + 1)
End Function


Function ValidarMarca(marca As String) As Boolean
    ValidarMarca = (Len(Trim(marca)) >= 2)
End Function


Function ValidarModelo(modelo As String) As Boolean
    ValidarModelo = (Len(Trim(modelo)) >= 2)
End Function


Function ValidarCor(cor As String) As Boolean
    ValidarCor = (Len(Trim(cor)) >= 2)
End Function
