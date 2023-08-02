Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM

Namespace Controllers
    Public Class PagoController
        Inherits ApiController

        ' GET: Pago
        Public Function PostValue(ByVal NombreBaseDatos As String, <FromBody()> ByVal pPagoSAP As Pagos) As RespuestaPagos

            Dim oPago As SAPbobsCOM.Payments
            Dim IntRespuesta As Integer = 0
            Dim IntCodigoError As Integer = 0
            Dim StrErrorSAP As String = ""
            Dim oRespuesta As New RespuestaPagos
            Dim oConexion As SAPbobsCOM.Company
            Dim lErrCode As Long
            Dim sErrMsg As String = ""
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty

            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then

                    oPago = Nothing
                    oPago = oConexion.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    oPago.DocType = BoRcptTypes.rAccount
                    oPago.DocDate = pPagoSAP.StrDocDate
                    oPago.DueDate = pPagoSAP.StrDocDueDate
                    oPago.DocCurrency = pPagoSAP.DocCurrency
                    oPago.TaxDate = pPagoSAP.StrTaxDate
                    oPago.Series = pPagoSAP.StrSerie
                    oPago.JournalRemarks = pPagoSAP.JournalRemarks

                    oPago.AccountPayments.AccountCode = pPagoSAP.AccountCode
                    oPago.AccountPayments.Decription = pPagoSAP.Decription
                    oPago.AccountPayments.SumPaid = pPagoSAP.SumPaid
                    oPago.AccountPayments.ProfitCenter = pPagoSAP.ProfitCenter

                    'Efectivo
                    If pPagoSAP.CashSum <> 0 Then
                        oPago.CashSum = pPagoSAP.CashSum
                        oPago.CashAccount = pPagoSAP.CashAccount
                    End If

                    'TRANSFERENCIA
                    If pPagoSAP.DblMontoTransferencia <> 0 Then
                        oPago.TransferSum = CDbl(pPagoSAP.DblMontoTransferencia) 'Monto de la transferencia
                        oPago.TransferAccount = pPagoSAP.StrCuentaTransferencia 'cuenta contable (enviar el SYS)
                        oPago.TransferDate = pPagoSAP.StrFechaTransferencia 'Fecha de la transferencia
                        oPago.TransferReference = pPagoSAP.StrReferenciaTransferencia 'Comentarios
                    End If

                    'TARJETA DE CREDITO
                    For i As Integer = 0 To pPagoSAP.LstTC.Count - 1
                        oPago.CreditCards.CreditCard = pPagoSAP.LstTC(i).StrCodigoTarjetaSAP 'Indicar el tipo de tarjeta Visa o mastercard (ver el codigo en el catalogo de sap)
                        oPago.CreditCards.CreditCardNumber = pPagoSAP.LstTC(i).StrNumeroTarjeta 'numero de tarjeta enviar los ultimos 4 digitos o colocar 1234
                        oPago.CreditCards.CardValidUntil = pPagoSAP.LstTC(i).StrFechaVencimiento 'fecha de vencimiento mandar formato yyyy-MM-dd
                        oPago.CreditCards.PaymentMethodCode = pPagoSAP.LstTC(i).StrCodigoFormaPagoSAP 'Forma de pago (ver el codigo en el catalogo de sap)
                        oPago.CreditCards.VoucherNum = pPagoSAP.LstTC(i).StrNoVoucher 'numero de voucher
                        oPago.CreditCards.CreditSum = CDbl(pPagoSAP.LstTC(i).DblMontoTarjeta) 'monto del pago
                        oPago.CreditCards.CreditAcct = pPagoSAP.LstTC(i).StrCuentaContableTarjeta 'cuenta contable (enviar el SYS)
                    Next

                    'CHEQUES
                    For i As Integer = 0 To pPagoSAP.LstCK.Count - 1
                        oPago.Checks.DueDate = pPagoSAP.LstCK(i).DueDate
                        oPago.Checks.CheckNumber = pPagoSAP.LstCK(i).CheckNumber
                        oPago.Checks.BankCode = pPagoSAP.LstCK(i).BankCode
                        oPago.Checks.CheckSum = pPagoSAP.LstCK(i).CheckSum
                        oPago.Checks.CheckAccount = pPagoSAP.LstCK(i).CheckAccount
                        oPago.Checks.CountryCode = pPagoSAP.LstCK(i).CountryCode
                    Next


                    IntRespuesta = oPago.Add

                    If IntRespuesta <> 0 Then
                        oConexion.GetLastError(IntCodigoError, StrErrorSAP)
                        oRespuesta.Mensaje = "Error Al Crear Pago: " & IntCodigoError & " - " & StrErrorSAP
                        oRespuesta.Estatus = False
                        oRespuesta.IDSAPDocumento = ""
                    Else
                        oRespuesta.Mensaje = "PAGO CREADO CORRECTAMENTE"
                        oRespuesta.Estatus = True
                        oRespuesta.IDSAPDocumento = oConexion.GetNewObjectKey().ToString()
                    End If

                Else
                    oRespuesta = New RespuestaPagos
                    oRespuesta.Mensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.Estatus = False
                    Return oRespuesta
                End If

                Return oRespuesta

            Catch ex As Exception
                oRespuesta = New RespuestaPagos
                oRespuesta.Mensaje = "Error De Servidor: " & ex.Message
                oRespuesta.Estatus = False
                Return oRespuesta
            End Try

        End Function

    End Class
End Namespace