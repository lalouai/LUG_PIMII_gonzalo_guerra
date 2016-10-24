Public Class Parser

    Private e As Integer
    Private exp As String
    Private rta As Double
    Private simbolo As String
    Private semaforo As Integer

    Sub New()
        e = 0
        exp = ""
        simbolo = ""
        semaforo = 0
    End Sub

    Public Event informacion(ByVal mensaje As String)

    Public Function parsear(ByVal expresion As String) As Double
        RaiseEvent informacion("Iniciando parseo...")
        exp = expresion + vbCrLf
        obtenerElemento()
        Try
            rta = sumaResta()
            If semaforo > 0 Then
                Throw New Excepciones("Expresion mal formada, nunca se cierra el parntesis")
            Else
                RaiseEvent informacion("Parseo terminado")
                Return rta
            End If
        Catch ex As Exception
            MsgBox("Excepcion" & vbCrLf & ex.Message)
        End Try

        Return Nothing
    End Function

    Private Sub obtenerElemento()
        simbolo = ""
        Dim aux As Char = Nothing
        aux = exp(e)
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
                RaiseEvent informacion("Número " & simbolo & " detectado")
                Return
            End If
            ''' multiplicacion
            If (aux.ToString = "*") Then
                RaiseEvent informacion("Multiplicación (*) detectada")
                simbolo = aux
                e += 1
                Return
            End If
            ''' division
            If (aux.ToString = "/") Then
                RaiseEvent informacion("División (/) detectada")
                simbolo = aux
                e += 1
                Return
            End If
            ''' suma
            If (aux.ToString = "+") Then
                RaiseEvent informacion("Suma (+) detectada")
                simbolo = aux
                e += 1
                Return
            End If
            ''' resta
            If (aux.ToString = "-") Then
                RaiseEvent informacion("Resta (-) detectada")
                simbolo = aux
                e += 1
                Return
            End If
            If (aux.ToString = "(") Then
                RaiseEvent informacion("Paréntesis ( ( ) detectado")
                simbolo = aux
                e += 1
                semaforo += 1
                Return
            End If
            If (aux.ToString = ")") Then
                RaiseEvent informacion("Parentesis ( ) ) detectado")
                simbolo = aux
                e += 1
                semaforo -= 1
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
                RaiseEvent informacion("Instanciando Suma")
                suma = New Suma
                suma.operando1 = salida
                suma.operando2 = multiDiv()
                salida = suma.resultado()
                RaiseEvent informacion("Resultado de la Suma -> " & salida)
            ElseIf (operador.Equals("-")) Then
                RaiseEvent informacion("Instanciando Resta")
                resta = New Resta
                resta.operando1 = salida
                resta.operando2 = multiDiv()
                salida = resta.resultado()
                RaiseEvent informacion("Resultado de la Resta -> " & salida)
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
                RaiseEvent informacion("Instanciando Multiplicación")
                multi = New Multiplicacion
                multi.operando1 = salida
                multi.operando2 = parentesis()
                salida = multi.resultado()
                RaiseEvent informacion("Resultado de la Multiplicación -> " & salida)
            ElseIf (operador.Equals("/")) Then
                RaiseEvent informacion("Instanciando División")
                divi = New Division
                divi.operando1 = salida
                divi.operando2 = parentesis()
                salida = divi.resultado()
                RaiseEvent informacion("Resultado de la División -> " & salida)
            End If
        End While
        Return salida
    End Function

    Private Function parentesis() As Double
        If (simbolo.Equals("(")) Then
            RaiseEvent informacion("Paréntesis")
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
        RaiseEvent informacion("Número")
        Dim salida As Double = 0
        salida = Double.Parse(simbolo)
        obtenerElemento()
        Return salida
    End Function

    Public Sub limpiar()
        e = 0
        exp = ""
        simbolo = ""
        semaforo = 0
    End Sub
End Class
