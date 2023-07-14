Public Class RespuestaCentrosCosto
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstCentroCosto As New List(Of CentroCosto)

    Public Class CentroCosto
        Public PrcCode As String = String.Empty
        Public PrcName As String = String.Empty
    End Class
End Class
