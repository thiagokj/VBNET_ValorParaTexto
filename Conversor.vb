Public Class Conversor
    'Conversor habilitado para escrever até Trilhões
    Public Shared Function EscreverExtenso(valor As Decimal) As String
        If valor < 0 Or valor >= 1000000000000000 Then
            Return "Valor não suportado pelo sistema."
        ElseIf valor = 0 Then
            Return "ZERO"
        Else
            'Converte valor recebido para uma string de 17 (0 a 16) posições
            Dim strValor As String = valor.ToString("000000000000000.00")
            Dim valorPorExtenso As String = String.Empty

            'Compara a cada 3 digitos a parte inteira do valor
            For digito As Integer = 0 To 15 Step 3
                valorPorExtenso += ComparaDigito(ConverteDigitoPorPosicao("DECIMAL", strValor, digito, 3))

                'Se o digito for ZERO e o valor por extenso estiver preenchido, compara TRILHÕES
                If digito = 0 And Not Equals(valorPorExtenso, String.Empty) Then

                    If ConverteDigitoPorPosicao("INT32", strValor, 0, 3) = 1 Then
                        valorPorExtenso += " TRILHÃO" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 3, 12) > 0, " E ", String.Empty)

                    ElseIf ConverteDigitoPorPosicao("INT32", strValor, 0, 3) > 1 Then
                        valorPorExtenso += " TRILHÕES" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 3, 12) > 0, " E ", String.Empty)
                    End If

                    'Se o digito for TRÊS e o valor por extenso estiver preenchido,, compara BILHÕES
                ElseIf digito = 3 And Not Equals(valorPorExtenso, String.Empty) Then

                    If ConverteDigitoPorPosicao("INT32", strValor, 3, 3) = 1 Then
                        valorPorExtenso += " BILHÃO" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 6, 9) > 0, " E ", String.Empty)

                    ElseIf ConverteDigitoPorPosicao("INT32", strValor, 3, 3) > 1 Then
                        valorPorExtenso += " BILHÕES" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 6, 9) > 0, " E ", String.Empty)
                    End If

                    'Se o digito for SEIS e o valor por extenso estiver preenchido, compara MILHÕES
                ElseIf digito = 6 And Not Equals(valorPorExtenso, String.Empty) Then

                    If ConverteDigitoPorPosicao("INT32", strValor, 6, 3) = 1 Then
                        valorPorExtenso += " MILHÃO" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 9, 6) > 0, " E ", String.Empty)

                    ElseIf ConverteDigitoPorPosicao("INT32", strValor, 6, 3) > 1 Then
                        valorPorExtenso += " MILHÕES" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 9, 6) > 0, " E ", String.Empty)
                    End If

                    'Se o digito for NOVE e o valor por extenso estiver preenchido, compara MILHARES
                ElseIf digito = 9 And Not Equals(valorPorExtenso, String.Empty) Then

                    If ConverteDigitoPorPosicao("INT32", strValor, 9, 3) > 0 Then
                        valorPorExtenso += " MIL" +
                            If(ConverteDigitoPorPosicao("DECIMAL", strValor, 12, 3) > 0, " E ", String.Empty)
                    End If
                End If

                'Se o digito for DOZE, avalia o tamanho do valor por extenso e acrescenta preposição e sufixo
                If digito = 12 Then

                    If valorPorExtenso.Length > 8 Then
                        Select Case valorSubstring(valorPorExtenso, 6, 6)
                            Case "MILHÃO", "BILHÃO"
                                valorPorExtenso += " DE"
                        End Select

                        Select Case valorSubstring(valorPorExtenso, 7, 7)
                            Case "MILHÕES", "BILHÕES"
                                valorPorExtenso += " DE"
                        End Select

                        Select Case valorSubstring(valorPorExtenso, 8, 7)
                            Case "TRILHÕES"
                                valorPorExtenso += " DE"
                        End Select

                        Select Case valorSubstring(valorPorExtenso, 8, 8)
                            Case "TRILHÕES"
                                valorPorExtenso += " DE"
                        End Select
                    End If

                    Select Case ConverteDigitoPorPosicao("INT64", strValor, 0, 15)
                        Case = 1
                            valorPorExtenso += " REAL"
                        Case > 1
                            valorPorExtenso += " REAIS"
                    End Select

                    If ConverteDigitoPorPosicao("INT32", strValor, 16, 2) > 0 _
                        AndAlso Not Equals(valorPorExtenso, String.Empty) Then
                        valorPorExtenso += " E "
                    End If
                End If

                'Se o digito for 15, avalia as casas após a virgula
                If digito = 15 Then
                    Select Case ConverteDigitoPorPosicao("INT32", strValor, 16, 2)
                        Case = 1
                            valorPorExtenso += " CENTAVO"
                        Case > 1
                            valorPorExtenso += " CENTAVOS"
                    End Select
                End If
            Next

            Return valorPorExtenso
        End If
    End Function

    Private Shared Function valorSubstring(valor As String, ini As Integer, fim As Integer)
        Return valor.Substring(valor.Length - ini, fim)
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
