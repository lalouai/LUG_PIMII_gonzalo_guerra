Public Class Division : Inherits Operacion
    Public Overloads Function resultado()
        If (operando2 = 0) Then
            Throw New Excepciones("Me resulta imposible dividir por cero... :(")
        End If

        Return operando1 / operando2
    End Function
End Class
