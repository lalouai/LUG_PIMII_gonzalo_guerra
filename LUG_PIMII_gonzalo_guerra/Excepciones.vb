Imports System

Public Class Excepciones
    Inherits Exception

    Public Sub New()
    End Sub

    Public Sub New(mensaje As String)
        MyBase.New(mensaje)
    End Sub
    Public Sub New(mensaje As String, inner As Exception)
        MyBase.New(mensaje, inner)
    End Sub

End Class
