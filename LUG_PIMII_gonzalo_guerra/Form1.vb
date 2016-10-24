Public Class Form1
    Dim WithEvents parseador As Parser

    Sub New()
        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        parseador = New Parser
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If tb_expresion.Text.Trim.Length > 0 Then
            lbl_resultado.Text = parseador.parsear(tb_expresion.Text)
        Else
            MsgBox("Lo siento, la expresin no puede ser vacia," & vbCrLf & " por favor intentelo nuevamente con un expresion adecuada")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        tb_expresion.Text = ""
        lbl_resultado.Text = ""
        lb_informacion.Items.Clear()
        parseador.limpiar()
        tb_expresion.Focus()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.Height > 300 Then
            Me.Height = 146
        Else
            Me.Height = 463
        End If
    End Sub

    Private Sub informacion(ByVal mensaje As String) Handles parseador.informacion
        lb_informacion.Items.Add(mensaje)
    End Sub

End Class
