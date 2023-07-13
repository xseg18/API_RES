Imports System.Security.Cryptography
Imports System.Web.Http
Imports System.Web.Mvc
Imports System.Xml
Imports SAPbobsCOM
Imports WS_REST.RespuestaArticulos

Namespace Controllers
    Public Class FacturasController
        Inherits ApiController

        ' POST: Facturas
        Public Function PostValue(ByVal NombreBaseDatos As String, <FromBody()> ByVal pFactura As FacturaSAP) As RespuestaFacturaSAP
            Dim oRespuesta As New RespuestaFacturaSAP
            Dim oConexion As SAPbobsCOM.Company
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim oDocumento As Documents
            Dim intRespuesta As Integer
            Dim lErrCode As Long
            Dim sErrMsg As String = ""

            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then
                    oDocumento = Nothing
                    oDocumento = oConexion.GetBusinessObject(BoObjectTypes.oInvoices)

                    oDocumento.CardCode = pFactura.CardCode
                    oDocumento.CardName = pFactura.CardName
                    oDocumento.TaxDate = CDate(pFactura.TaxDate)
                    oDocumento.DocDate = CDate(pFactura.DocDate)
                    oDocumento.DocDueDate = CDate(pFactura.DocDueDate)
                    oDocumento.Comments = pFactura.Comments
                    oDocumento.Address = pFactura.Address

                    For i As Integer = 0 To pFactura.LstItems.Count - 1

                        If i > 0 Then
                            oDocumento.Lines.Add()
                        End If

                        oDocumento.Lines.ItemCode = pFactura.LstItems.Item(i).ItemCode
                        oDocumento.Lines.ItemDescription = pFactura.LstItems.Item(i).ItemDescription
                        oDocumento.Lines.TaxCode = pFactura.LstItems.Item(i).TaxCode
                        oDocumento.Lines.Quantity = pFactura.LstItems.Item(i).Quantity
                        oDocumento.Lines.PriceAfterVAT = pFactura.LstItems.Item(i).PriceAfterVAT
                    Next

                    intRespuesta = oDocumento.Add()

                    If intRespuesta = 0 Then
                        oRespuesta.Mensaje = "Factura Creada Correctamente"
                        oRespuesta.Estado = True
                        oRespuesta.IDFacturaSAP = oConexion.GetNewObjectKey
                    Else
                        oConexion.GetLastError(lErrCode, sErrMsg)
                        oRespuesta.Mensaje = "Error Al Crear Factura: " & lErrCode & " - " & sErrMsg
                        oRespuesta.Estado = False
                    End If

                Else
                    oRespuesta = New RespuestaFacturaSAP
                    oRespuesta.Mensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.Estado = False
                    Return oRespuesta
                End If

                Return oRespuesta

            Catch ex As Exception
                oRespuesta = New RespuestaFacturaSAP
                oRespuesta.Mensaje = "Error De Servidor: " & ex.Message
                oRespuesta.Estado = False
                Return oRespuesta
            End Try

        End Function
    End Class
End Namespace