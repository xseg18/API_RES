Public Class RespuestaCuentasContables
    Public StrMensaje As String
    Public BlnEstado As Boolean
    Public LstCuentaContable As New List(Of CuentaContable)

    Public Class CuentaContable
        Public AcctCode As String = String.Empty
        Public AcctName As String = String.Empty
        Public CurrTotal As String = String.Empty
    End Class
End Class
