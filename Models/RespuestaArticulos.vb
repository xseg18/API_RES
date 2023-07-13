Public Class RespuestaArticulos

    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstArticulos As New List(Of Articulos)

    Public Class Articulos
        Public ItemCode As String = String.Empty
        Public ItemName As String = String.Empty
        Public FrgnName As String = String.Empty
    End Class

End Class
