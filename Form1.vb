
Public Class Form1
    Private Sub txtNumero_TextChanged(sender As Object, e As EventArgs) Handles txtNumero.TextChanged
        If txtNumero.TextLength > 0 Then
            Dim numDecimal As Decimal = txtNumero.Text
            txtValor.Text = numDecimal.ToString("C")
            txtValorDescri.Text = Conversor.EscreverExtenso(txtValor.Text)

        ElseIf txtNumero.TextLength = 0 Then
            txtValor.Clear()
            txtValorDescri.Clear()
        End If
    End Sub

End Class
