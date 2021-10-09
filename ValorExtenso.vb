Public Class ValorExtenso
    Shared ones As String() = New String() {"", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"}
    Shared teens As String() = New String() {"Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"}
    Shared tens As String() = New String() {"Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"}
    Shared thousandsGroups As String() = {"", " Thousand", " Million", " Billion"}

    Private Function FriendlyInteger(ByVal n As Integer, ByVal leftDigits As String, ByVal thousands As Integer) As String
        If n = 0 Then
            Return leftDigits
        End If

        Dim friendlyInt As String = leftDigits

        If friendlyInt.Length > 0 Then
            friendlyInt += " "
        End If

        If n < 10 Then
            friendlyInt += ones(n)
        ElseIf n < 20 Then
            friendlyInt += teens(n - 10)
        ElseIf n < 100 Then
            friendlyInt += FriendlyInteger(n Mod 10, tens(n / 10 - 2), 0)
        ElseIf n < 1000 Then
            friendlyInt += FriendlyInteger(n Mod 100, (ones(n / 100) & " Hundred"), 0)
        Else
            friendlyInt += FriendlyInteger(n Mod 1000, FriendlyInteger(n / 1000, "", thousands + 1), 0)

            If n Mod 1000 = 0 Then
                Return friendlyInt
            End If
        End If

        Return friendlyInt & thousandsGroups(thousands)
    End Function

    Function IntegerToWritten(ByVal n As Integer) As String
        If n = 0 Then
            Return "Zero"
        ElseIf n < 0 Then
            Return "Negative " & IntegerToWritten(-n)
        End If

        Return FriendlyInteger(n, "", 0)
    End Function
End Class
