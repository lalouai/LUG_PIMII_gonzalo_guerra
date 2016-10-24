Public Class Parser

    Private e As Integer
    Private exp As String
    Private rta As Double
    Private simbolo As String
    Private longitud As Integer


    Sub New()
        e = 0
        exp = ""
        simbolo = ""
        longitud = 0
    End Sub

    Public Event informacion(ByVal mensaje As String)

    Public Function parsear(ByVal expresion As String) As Double
        RaiseEvent informacion("Iniciando parseo...")
        exp = expresion + "=" + vbCrLf
        obtenerElemento()
        rta = sumaResta()
        longitud = expresion.Length
        RaiseEvent informacion("Parseo terminado")
        Return rta
    End Function

    Private Sub obtenerElemento()

        simbolo = ""
        Dim aux As Char = Nothing
        aux = exp(e)

        RaiseEvent informacion("exp(e) -> " & exp(e))
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

                RaiseEvent informacion("obtenerElemento() --> número " & simbolo)
                Return
            End If
            ''' multiplicacion
            If (aux.ToString = "*") Then
                RaiseEvent informacion("obtenerElemento() --> multiplicación *")
                simbolo = aux
                e += 1
                Return
            End If
            ''' division
            If (aux.ToString = "/") Then
                RaiseEvent informacion("obtenerElemento() --> división /")
                simbolo = aux
                e += 1
                Return
            End If
            ''' suma
            If (aux.ToString = "+") Then
                RaiseEvent informacion("obtenerElemento() --> suma +")
                simbolo = aux
                e += 1
                Return
            End If
            ''' resta
            If (aux.ToString = "-") Then
                RaiseEvent informacion("obtenerElemento() --> resta -")
                simbolo = aux
                e += 1
                Return
            End If

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
        salida = Double.Parse(simbolo)
        obtenerElemento()
        Return salida
    End Function

End Class
