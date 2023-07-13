Public Class RespuestaTasaCambio
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstTasaCambio As New List(Of TasaCambio)

    Public Class TasaCambio
        Public RateDate As String = String.Empty
        Public currency As String = String.Empty
        Public Rate As String = String.Empty
    End Class
End Class
