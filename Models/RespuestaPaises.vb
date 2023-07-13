Public Class RespuestaPaises
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstPaises As New List(Of Paises)

    Public Class Paises
        Public Code As String = String.Empty
        Public Name As String = String.Empty
    End Class
End Class
