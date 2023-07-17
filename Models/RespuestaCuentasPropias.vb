

Public Class RespuestaCuentasPropias
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstCuentasPropias As New List(Of CuentaPropia)

    Public Class CuentaPropia
        Public BankCode As String
        Public Account As String
    End Class
End Class
