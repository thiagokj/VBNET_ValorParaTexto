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

                'Se o digito for DOZE, avalia o tamanho do valor por extenso e acrescenta sufixo singular ou plural
                If digito = 12 Then

                    If valorPorExtenso.Length > 8 Then

                        If Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 6, 6), "BILHÃO") _
                            Or Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 6, 6), "MILHÃO") Then
                            valorPorExtenso += " DE"

                        ElseIf Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 7, 7), "BILHÕES") _
                            Or Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 7, 7), "MILHÕES") _
                            Or Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 8, 7), "TRILHÕES") Then
                            valorPorExtenso += " DE"

                        ElseIf Equals(valorPorExtenso.Substring(valorPorExtenso.Length - 8, 8), "TRILHÕES") Then
                            valorPorExtenso += " DE"
                        End If

                    End If

                    If ConverteDigitoPorPosicao("INT64", strValor, 0, 15) = 1 Then
                        valorPorExtenso += " REAL"
                    ElseIf ConverteDigitoPorPosicao("INT64", strValor, 0, 15) > 1 Then
                        valorPorExtenso += " REAIS"
                    End If

                    If ConverteDigitoPorPosicao("INT32", strValor, 16, 2) > 0 _
                        AndAlso Not Equals(valorPorExtenso, String.Empty) Then
                        valorPorExtenso += " E "
                    End If
                End If

                'Se o digito for 15, avalia as casas após a virgula
                If digito = 15 Then

                    If ConverteDigitoPorPosicao("INT32", strValor, 16, 2) = 1 Then
                        valorPorExtenso += " CENTAVO"

                    ElseIf ConverteDigitoPorPosicao("INT32", strValor, 16, 2) > 1 Then
                        valorPorExtenso += " CENTAVOS"
                    End If

                End If
            Next

            Return valorPorExtenso
        End If
    End Function

    Private Shared Function ConverteDigitoPorPosicao(tipoValor As String, strValor As String,
                                                    posicaoIni As Integer, posicaoFim As Integer)
        If tipoValor = "DECIMAL" Then
            Return Convert.ToDecimal(strValor.Substring(posicaoIni, posicaoFim))

        ElseIf tipoValor = "INT32" Then
            Return Convert.ToInt32(strValor.Substring(posicaoIni, posicaoFim))

        ElseIf tipoValor = "INT64" Then
            Return Convert.ToInt64(strValor.Substring(posicaoIni, posicaoFim))

        Else
            Return MsgBox("Valor não pode ser comparado")
        End If
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
            If posicao0 = 1 Then
                valorExtenso += If(posicao1 + posicao2 = 0, "CEM", "CENTO")
            ElseIf posicao0 = 2 Then
                valorExtenso += "DUZENTOS"
            ElseIf posicao0 = 3 Then
                valorExtenso += "TREZENTOS"
            ElseIf posicao0 = 4 Then
                valorExtenso += "QUATROCENTOS"
            ElseIf posicao0 = 5 Then
                valorExtenso += "QUINHENTOS"
            ElseIf posicao0 = 6 Then
                valorExtenso += "SEISCENTOS"
            ElseIf posicao0 = 7 Then
                valorExtenso += "SETECENTOS"
            ElseIf posicao0 = 8 Then
                valorExtenso += "OITOCENTOS"
            ElseIf posicao0 = 9 Then
                valorExtenso += "NOVECENTOS"
            End If

            'Compara DEZENAS
            If posicao1 = 1 Then

                If posicao2 = 0 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZ"
                ElseIf posicao2 = 1 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "ONZE"
                ElseIf posicao2 = 2 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DOZE"
                ElseIf posicao2 = 3 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "TREZE"
                ElseIf posicao2 = 4 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUATORZE"
                ElseIf posicao2 = 5 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUINZE"
                ElseIf posicao2 = 6 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZESSEIS"
                ElseIf posicao2 = 7 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZESSETE"
                ElseIf posicao2 = 8 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZOITO"
                ElseIf posicao2 = 9 Then
                    valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "DEZENOVE"
                End If

            ElseIf posicao1 = 2 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "VINTE"
            ElseIf posicao1 = 3 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "TRINTA"
            ElseIf posicao1 = 4 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "QUARENTA"
            ElseIf posicao1 = 5 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "CINQUENTA"
            ElseIf posicao1 = 6 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "SESSENTA"
            ElseIf posicao1 = 7 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "SETENTA"
            ElseIf posicao1 = 8 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "OITENTA"
            ElseIf posicao1 = 9 Then
                valorExtenso += If(posicao0 > 0, " E ", String.Empty) + "NOVENTA"
            End If

            If Not Equals(strValor.Substring(1, 1), "1") And posicao2 <> 0 _
                And Not Equals(valorExtenso, String.Empty) Then
                valorExtenso += " E "
            End If

            'Compara UNIDADES
            If Not Equals(strValor.Substring(1, 1), "1") Then

                If posicao2 = 1 Then
                    valorExtenso += "UM"
                ElseIf posicao2 = 2 Then
                    valorExtenso += "DOIS"
                ElseIf posicao2 = 3 Then
                    valorExtenso += "TRÊS"
                ElseIf posicao2 = 4 Then
                    valorExtenso += "QUATRO"
                ElseIf posicao2 = 5 Then
                    valorExtenso += "CINCO"
                ElseIf posicao2 = 6 Then
                    valorExtenso += "SEIS"
                ElseIf posicao2 = 7 Then
                    valorExtenso += "SETE"
                ElseIf posicao2 = 8 Then
                    valorExtenso += "OITO"
                ElseIf posicao2 = 9 Then
                    valorExtenso += "NOVE"
                End If
            End If

            Return valorExtenso
        End If
    End Function
End Class
