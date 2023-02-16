Public Class ValorPorExtensoV3
    'Todo
    Private Shared Function ConverteNumeroPorExtenso(ByVal numero As Long) As String
        Dim unidades() As String = {"", "UM", "DOIS", "TRÊS", "QUATRO", "CINCO", "SEIS", "SETE", "OITO", "NOVE", "DEZ",
                                "ONZE", "DOZE", "TREZE", "QUATORZE", "QUINZE", "DEZESSEIS", "DEZESSETE", "DEZOITO", "DEZENOVE"}
        Dim dezenas() As String = {"", "", "VINTE", "TRINTA", "QUARENTA", "CINQUENTA", "SESSENTA", "SETENTA", "OITENTA", "NOVENTA"}
        Dim centenas() As String = {"", "CEM", "DUZENTOS", "TREZENTOS", "QUATROCENTOS", "QUINHENTOS", "SEISCENTOS", "SETECENTOS", "OITOCENTOS", "NOVECENTOS"}

        If numero = 0 Then
            Return "ZERO"
        ElseIf numero < 0 Then
            Return "MENOS " & ConverteNumeroPorExtenso(-numero)
        End If

        Dim partes() As Long = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim contador As Integer = 0
        Dim resto As Long = 0

        While numero > 0
            resto = numero Mod 1000
            partes(contador) = resto
            contador += 1
            numero \= 1000
        End While

        Dim extenso As String = ""

        For i As Integer = contador - 1 To 0 Step -1
            If partes(i) > 0 Then
                If i > 0 Then
                    If partes(i) = 1 Then
                        extenso &= "MIL "
                    Else
                        extenso &= ConverteNumeroPorExtenso(partes(i)) & " MIL "
                    End If
                ElseIf partes(i) = 1 Then
                    extenso &= "UM "
                ElseIf partes(i) > 1 And partes(i) < 20 Then
                    extenso &= unidades(partes(i)) & " "
                ElseIf partes(i) >= 20 Then
                    extenso &= dezenas(partes(i) \ 10) & " "
                    If (partes(i) Mod 10) > 0 Then
                        extenso &= "E " & unidades(partes(i) Mod 10) & " "
                    End If
                End If
            End If
        Next

        Return extenso.TrimEnd()
    End Function
End Class
