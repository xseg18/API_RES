
Public Class RespuestaClientes

    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstClientes As New List(Of Cliente)

    Public Class Cliente
        Public ID As String = String.Empty
        Public Name As String = String.Empty
        Public CommercialName As String = String.Empty
        Public Address As String = String.Empty
        Public NIT As String = String.Empty
        Public Phone As String = String.Empty
    End Class
End Class
