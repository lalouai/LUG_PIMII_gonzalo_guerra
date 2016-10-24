Public Class Division : Inherits Operacion
    Public Overloads Function resultado()
        Try
            If (operando2 = 0) Then Throw New Excepciones("Me resulta imposible dividir por cero... :(")
            Return operando1 / operando2
        Catch ex As Excepciones
            MsgBox("Excepcion" & vbCrLf & ex.Message)
        End Try
        Return Nothing
    End Function
End Class
