Public Class Operacion
    Dim _operando1 As Double
    Dim _operando2 As Double

    Sub Operacion()

    End Sub

    Public Property operando1 As Double
        Get
            Return _operando1
        End Get
        Set(value As Double)
            _operando1 = value
        End Set
    End Property

    Public Property operando2 As Double
        Get
            Return _operando2
        End Get
        Set(value As Double)
            _operando2 = value
        End Set
    End Property

    Function resultado() As Double
        Return 0
    End Function


End Class
