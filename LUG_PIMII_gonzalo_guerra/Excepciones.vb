Imports System

Public Class Excepciones
    Inherits Exception

    Private _mensaje As String

    Public Sub New()
    End Sub

    Public Sub New(mensaje As String)
        _mensaje = mensaje
    End Sub

    Public Overrides ReadOnly Property Message As String
        Get
            Return _mensaje
        End Get
    End Property

End Class
