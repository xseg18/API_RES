Public Class Pagos


    Public StrDocDate As String = String.Empty
    Public StrTaxDate As String = String.Empty
    Public StrDocDueDate As String = String.Empty
    Public StrTotalPago As String = String.Empty
    Public DocCurrency As String = String.Empty
    Public StrSerie As String = String.Empty
    Public JournalRemarks As String = String.Empty

    Public AccountCode As String = String.Empty
    Public Decription As String = String.Empty
    Public SumPaid As Double = 0
    Public ProfitCenter As String = String.Empty

    Public CashSum As Double = 0
    Public CashAccount As String = String.Empty

    Public DblMontoTransferencia As Double = 0
    Public StrCuentaTransferencia As String = String.Empty
    Public StrFechaTransferencia As String = String.Empty
    Public StrReferenciaTransferencia As String = String.Empty

    Public LstTC As New List(Of TarjetaCredito)
    Public LstCK As New List(Of Cheques)
    Public Class TarjetaCredito
        Public StrCodigoTarjetaSAP As String = String.Empty
        Public StrCodigoFormaPagoSAP As String = String.Empty
        Public StrNumeroTarjeta As String = String.Empty
        Public StrFechaVencimiento As String = String.Empty
        Public DblMontoTarjeta As Double
        Public StrCuentaContableTarjeta As String = String.Empty
        Public StrNoVoucher As String = String.Empty
    End Class

    Public Class Cheques
        Public DueDate As String = String.Empty
        Public CheckNumber As Integer = 0
        Public BankCode As String = String.Empty
        Public CheckSum As Double = 0
        Public CheckAccount As String = String.Empty
        Public CountryCode As String = String.Empty
    End Class

End Class
