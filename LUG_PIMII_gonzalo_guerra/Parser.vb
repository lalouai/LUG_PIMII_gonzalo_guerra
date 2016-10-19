Public Class Parser

    Private e As Integer
    Private exp As String
    Private rta As Double
    Private simbolo As String

    Sub New()
        e = 0
        exp = ""
        simbolo = ""
        RaiseEvent informacion("Parser instanciado")
    End Sub

    Public Event informacion(ByVal mensaje As String)

    Public Function parsear(ByVal expresion As String) As Double
        RaiseEvent informacion("Iniciando parseo...")
        exp = expresion
        obtenerElemento()
        RaiseEvent informacion("Símbolo detectado ->" & simbolo)
        rta = sumaResta()
        Return rta
        RaiseEvent informacion("Parseo terminado")
    End Function

    Private Sub obtenerElemento()
        simbolo = ""
        Dim aux As Char = Nothing
        aux = exp(e)
        RaiseEvent informacion("AUX -> " & aux)
        While (aux.Equals(" ") Or aux.Equals("\t"))
            e += 1
            aux = exp(e)
        End While

        If (Not aux.Equals(vbCrLf)) Then

            ''' Mientras sea número
            If (IsNumeric(aux)) Then
                While (IsNumeric(aux))
                    simbolo += aux
                    e += 1
                    aux = exp(e)
                End While
            End If
            Return

        End If
        ''' multiplicacion
        If (aux.Equals("*")) Then
            simbolo = aux
            e += 1
            Return
        End If
        ''' division
        If (aux.Equals("/")) Then
            simbolo = aux
            e += 1
            Return
        End If
        ''' suma
        If (aux.Equals("+")) Then
            simbolo = aux
            e += 1
            Return
        End If
        ''' resta
        If (aux.Equals("-")) Then
            simbolo = aux
            e += 1
            Return
        End If
        simbolo = ""
        Return
    End Sub

    Private Function sumaResta() As Double
        Dim salida As Double
        Dim suma As Suma
        Dim resta As Resta
        salida = multiDiv()
        While (simbolo.Equals("+") Or simbolo.Equals("-"))
            Dim operador As String = simbolo
            obtenerElemento()
            RaiseEvent informacion("Símbolo detectado ->" & simbolo)
            If (operador.Equals("+")) Then
                RaiseEvent informacion("SUMA")
                suma = New Suma
                suma.operando1 = salida
                suma.operando2 = multiDiv()
                salida = suma.resultado()
            ElseIf (operador.Equals("-")) Then
                RaiseEvent informacion("RESTA")
                resta = New Resta
                resta.operando1 = salida
                resta.operando2 = multiDiv()
                salida = resta.resultado()
            End If
        End While
        Return salida
    End Function

    Private Function multiDiv() As Double
        Dim salida As Double
        Dim multi As Multiplicacion
        Dim divi As Division
        salida = parentesis()
        While (simbolo.Equals("*") Or simbolo.Equals("/"))
            Dim operador As String = simbolo
            obtenerElemento()
            RaiseEvent informacion("Símbolo detectado ->" & simbolo)
            If (operador.Equals("*")) Then
                RaiseEvent informacion("MULTIPLICACIÓN")
                multi = New Multiplicacion
                multi.operando1 = salida
                multi.operando2 = parentesis()
                salida = multi.resultado()
            ElseIf (operador.Equals("/")) Then
                RaiseEvent informacion("DIVISIÓN")
                divi = New Division
                divi.operando1 = salida
                divi.operando2 = parentesis()
                salida = divi.resultado()
            End If
        End While
        Return salida
    End Function



    Private Function parentesis() As Double
        RaiseEvent informacion("PARÉNTESIS")
        If (simbolo.Equals("(")) Then
            obtenerElemento()
            RaiseEvent informacion("Símbolo detectado ->" & simbolo)
            Dim salida As Double = sumaResta()
            If (simbolo.Equals(")")) Then
                Throw New Excepciones("Debe cerrar los paréntesis abiertos, expresión mal formada")
            End If
            obtenerElemento()
            Return salida
        End If
        Return numeros()
    End Function

    Private Function numeros()
        RaiseEvent informacion("NUMEROS")
        Dim salida As Double = 0
        salida = Integer.Parse(simbolo)
        obtenerElemento()
        RaiseEvent informacion("Símbolo detectado ->" & simbolo)
        Return salida
    End Function

End Class
