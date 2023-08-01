Public Class RespuestaAlmacenes
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstAlmacenes As New List(Of Almacen)

    Public Class Almacen
        Public WhsCode As String = String.Empty
        Public WhsName As String = String.Empty
    End Class
End Class