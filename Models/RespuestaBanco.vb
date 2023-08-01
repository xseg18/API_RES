Public Class RespuestaBanco
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstBancos As New List(Of Banco)

    Public Class Banco
        Public BankCode As String = String.Empty
        Public BankName As String = String.Empty
    End Class
End Class