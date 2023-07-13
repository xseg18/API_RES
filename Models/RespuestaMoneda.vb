Public Class RespuestaMoneda

    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstMonedas As New List(Of Monedas)

    Public Class Monedas
        Public StrCodigoMoneda As String = String.Empty
        Public StrNombreMoneda As String = String.Empty
    End Class

End Class
