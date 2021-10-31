
Public Class FormPrincipal
    Private Sub txtNumero_TextChanged(sender As Object, e As EventArgs) Handles txtNumero.TextChanged
        If txtNumero.TextLength > 0 Then
            Dim numDecimal As Decimal = txtNumero.Text
            'Converte valor para cultura corrente (R$)
            txtValor.Text = numDecimal.ToString("C")
            txtValorDescri.Text = ValorPorExtenso.Converter(txtValor.Text)

        ElseIf txtNumero.TextLength = 0 Then
            txtValor.Clear()
            txtValorDescri.Clear()
        End If
    End Sub

End Class
