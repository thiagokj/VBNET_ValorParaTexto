Public Class ValorPorExtensoV2
    'Conversor habilitado para escrever até Trilhões
    Public Shared Function Converte(valor As Decimal) As String
        If valor <= 0 Then Return "ZERO"
        If valor >= 1000000000000000 Then Return "Valor não suportado pelo sistema."

        'Converte valor recebido para uma string de 17 (0 a 16) posições
        Dim strValor = valor.ToString("000000000000000.00")
        Return ConverteValorRecebido(strValor)
    End Function

    Private Shared Function ConverteValorRecebido(strValor As String)
        Dim valorPorExtenso As String = String.Empty

        'Compara a cada 3 digitos a parte inteira do valor
        For digito As Integer = 0 To 15 Step 3
            valorPorExtenso += ComparaDigito(ConverteDigitoPorPosicao(
                "DECIMAL", strValor, digito, 3))

            If digito = 0 And Not Equals(valorPorExtenso, String.Empty) Then
                valorPorExtenso = AdicionaSufixoPorDigito(valorPorExtenso, strValor,
                    0, "TRILHÃO", "TRILHÕES")
            End If

            If digito = 3 And Not Equals(valorPorExtenso, String.Empty) Then
                valorPorExtenso = AdicionaSufixoPorDigito(valorPorExtenso, strValor,
                    3, "BILHÃO", "BILHÕES")
            End If

            If digito = 6 And Not Equals(valorPorExtenso, String.Empty) Then
                valorPorExtenso = AdicionaSufixoPorDigito(valorPorExtenso, strValor,
                    6, "MILHÃO", "MILHÕES")
            End If

            If digito = 9 And Not Equals(valorPorExtenso, String.Empty) Then
                valorPorExtenso = AdicionaSufixoPorDigito(valorPorExtenso, strValor,
                    9, "MIL", "MIL")
            End If

            If digito = 12 Then
                valorPorExtenso = AdicionarDe(valorPorExtenso, 6, 6)
                valorPorExtenso = AdicionarDe(valorPorExtenso, 7, 7)
                valorPorExtenso = AdicionarDe(valorPorExtenso, 8, 8)
                valorPorExtenso = AdicionarDe(valorPorExtenso, 8, 7)
                valorPorExtenso = AdicionarReal(valorPorExtenso, strValor,
                    ConverteDigitoPorPosicao("INT64", strValor, 0, 15))
            End If

            ' Se o digito for 15, avalia as casas após a vírgula
            If digito = 15 Then
                valorPorExtenso = AdicionarCentavos(valorPorExtenso,
                    ConverteDigitoPorPosicao("INT64", strValor, 16, 2))
            End If
        Next

        Return valorPorExtenso
    End Function

    Private Shared Function AdicionarCentavos(valorPorExtenso As String, valor As Long) As String
        If valor = 1 Then
            valorPorExtenso += " CENTAVO"
        ElseIf valor > 1 Then
            valorPorExtenso += " CENTAVOS"
        End If

        Return valorPorExtenso
    End Function

    Private Shared Function AdicionarDe(valorPorExtenso As String, posInicial As Integer, posFinal As Integer) As String
        If valorPorExtenso.Length > posFinal AndAlso valorPorExtenso.Substring(posInicial, posFinal - posInicial + 1) = "MILHÃO" Then
            valorPorExtenso += " DE"
        ElseIf valorPorExtenso.Length > posFinal AndAlso valorPorExtenso.Substring(posInicial, posFinal - posInicial + 1) = "BILHÃO" Then
            valorPorExtenso += " DE"
        ElseIf valorPorExtenso.Length > posFinal + 1 AndAlso valorPorExtenso.Substring(posInicial, posFinal - posInicial + 2) = "MILHÕES" Then
            valorPorExtenso += " DE"
        ElseIf valorPorExtenso.Length > posFinal + 1 AndAlso valorPorExtenso.Substring(posInicial, posFinal - posInicial + 2) = "BILHÕES" Then
            valorPorExtenso += " DE"
        ElseIf valorPorExtenso.Length > posFinal + 1 AndAlso valorPorExtenso.Substring(posInicial, posFinal - posInicial + 2) = "TRILHÕES" Then
            valorPorExtenso += " DE"
        End If

        Return valorPorExtenso
    End Function

    Private Shared Function AdicionarReal(valorPorExtenso As String, strValor As String, valor As Long) As String
        If valor = 1 Then
            valorPorExtenso += " REAL"
        ElseIf valor > 1 Then
            valorPorExtenso += " REAIS"
        End If

        If ConverteDigitoPorPosicao("INT32", strValor, 16, 2) > 0 _
            AndAlso Not String.IsNullOrEmpty(valorPorExtenso) Then
            valorPorExtenso += " E "
        End If

        Return valorPorExtenso
    End Function

    Private Shared Function AdicionaSufixoPorDigito(valorPorExtenso As String, strValor As String, digito As Integer, sufixoSingular As String, sufixoPlural As String) As String
        Dim valor As Integer = ConverteDigitoPorPosicao("INT32", strValor, digito, 3)
        Dim plural = valor > 1
        If valor > 0 Then
            valorPorExtenso += $" {IIf(plural, sufixoPlural, sufixoSingular)}" +
            If(digito = 9, " ", If(ConverteDigitoPorPosicao("DECIMAL", strValor, digito + 3, 3) > 0,
                " E ", String.Empty))
        End If
        Return valorPorExtenso
    End Function

    Private Shared Function ConverteDigitoPorPosicao(tipoValor As String, strValor As String,
                                                         posicaoIni As Integer, posicaoFim As Integer)
        Select Case tipoValor
            Case "DECIMAL"
                Return Convert.ToDecimal(strValor.Substring(posicaoIni, posicaoFim))
            Case "INT32"
                Return Convert.ToInt32(strValor.Substring(posicaoIni, posicaoFim))
            Case "INT64"
                Return Convert.ToInt64(strValor.Substring(posicaoIni, posicaoFim))
            Case Else
                Return MsgBox("Valor não pode ser comparado")
        End Select
    End Function

    Private Shared Function ComparaDigito(valor As Decimal) As String
        If valor <= 0 Then
            Return String.Empty
        Else
            Dim valorExtenso As String = String.Empty

            'Se o valor for informado como (0,XX), multiplica para comparar 3 posições
            If valor > 0 And valor < 1 Then
                valor *= 100
            End If

            Dim strValor As String = valor.ToString("000")

            Dim posicao0 As Integer = ConverteDigitoPorPosicao("INT32", strValor, 0, 1)
            Dim posicao1 As Integer = ConverteDigitoPorPosicao("INT32", strValor, 1, 1)
            Dim posicao2 As Integer = ConverteDigitoPorPosicao("INT32", strValor, 2, 1)

            'Compara CENTENAS
            Select Case posicao0
                Case = 1
                    valorExtenso += If(posicao1 + posicao2 = 0, "CEM", "CENTO")
                Case = 2
                    valorExtenso += "DUZENTOS"
                Case = 3
                    valorExtenso += "TREZENTOS"
                Case = 4
                    valorExtenso += "QUATROCENTOS"
                Case = 5
                    valorExtenso += "QUINHENTOS"
                Case = 6
                    valorExtenso += "SEISCENTOS"
                Case = 7
                    valorExtenso += "SETECENTOS"
                Case = 8
                    valorExtenso += "OITOCENTOS"
                Case = 9
                    valorExtenso += "NOVECENTOS"
            End Select

            'Compara DEZENAS
            Select Case posicao1
                Case = 1
                    'Avalia numeros de 10 a 19
                    Select Case posicao2
                        Case = 0
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZ"
                        Case = 1
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "ONZE"
                        Case = 2
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DOZE"
                        Case = 3
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "TREZE"
                        Case = 4
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUATORZE"
                        Case = 5
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUINZE"
                        Case = 6
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZESSEIS"
                        Case = 7
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZESSETE"
                        Case = 8
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZOITO"
                        Case = 9
                            valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZENOVE"
                    End Select

                Case = 2
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "VINTE"
                Case = 3
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "TRINTA"
                Case = 4
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUARENTA"
                Case = 5
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "CINQUENTA"
                Case = 6
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "SESSENTA"
                Case = 7
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "SETENTA"
                Case = 8
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "OITENTA"
                Case = 9
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "NOVENTA"
            End Select

            'Adiciona preposição 'E' entre as Dezenas e Unidades
            If strValor.Substring(1, 1) <> "1" _
                    And posicao2 <> 0 _
                    And valorExtenso <> String.Empty Then
                valorExtenso += " E "
            End If

            'Compara UNIDADES
            If strValor.Substring(1, 1) <> "1" Then
                Select Case posicao2
                    Case = 1
                        valorExtenso += "UM"
                    Case = 2
                        valorExtenso += "DOIS"
                    Case = 3
                        valorExtenso += "TRÊS"
                    Case = 4
                        valorExtenso += "QUATRO"
                    Case = 5
                        valorExtenso += "CINCO"
                    Case = 6
                        valorExtenso += "SEIS"
                    Case = 7
                        valorExtenso += "SETE"
                    Case = 8
                        valorExtenso += "OITO"
                    Case = 9
                        valorExtenso += "NOVE"
                End Select
            End If

            Return valorExtenso
        End If
    End Function
End Class


