Imports System.Security.Cryptography
Imports System.Web.Http
Imports System.Web.Mvc
Imports System.Xml
Imports SAPbobsCOM
Imports WS_REST.RespuestaArticulos
Imports WS_REST.FacturaSAP
Namespace Controllers
    Public Class FacturasAsientoController
        Inherits ApiController

        ' POST: Facturas
        Public Function PostValue(ByVal NombreBaseDatos As String, <FromBody()> ByVal pFactura As FacturaSAP) As RespuestaFacturaSAP
            Dim oRespuesta As New RespuestaFacturaSAP
            Dim oConexion As SAPbobsCOM.Company
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim oDocumento As Documents
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Dim intRespuesta As Integer
            Dim lErrCode As Long
            Dim sErrMsg As String = ""
            Dim Art As SAPbobsCOM.Items
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then
                    oDocumento = Nothing
                    oDocumento = oConexion.GetBusinessObject(BoObjectTypes.oJournalEntries)

                    oDocumento.JournalMemo = pFactura.Comments
                    oDocumento.DocTotal = pFactura.DocTotal
                    oDocumento.TaxDate = CDate(pFactura.TaxDate)
                    oDocumento.DocDate = CDate(pFactura.DocDate)
                    oDocumento.DocDueDate = CDate(pFactura.DocDueDate)

                    oDocumento.Lines.ShortName = pFactura.CardCode
                    oDocumento.Lines.Debit = pFactura.DocTotal
                    oDocumento.Lines.Credit = 0
                    oDocumento.Lines.Add()

                    For Each itm In pFactura.LstItems
                        Dim Item As SAPbobsCOM.Items = oConexion.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        Item.GetByKey(itm.ItemCode)
                        Dim ItemGroup As SAPbobsCOM.ItemGroups = oConexion.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        ItemGroup.GetByKey(Item.ItemsGroupCode)
                        oDocumento.Lines.ShortName = ItemGroup.ExpensesAccount
                        If (itm.TaxCode = "IVA") Then
                            oDocumento.Lines.Credit = itm.PriceAfterVAT / 1.12
                            oDocumento.Lines.Debit = 0
                            oDocumento.Lines.Add()
                            oDocumento.Lines.ShortName = "_SYS00000000345"
                            oDocumento.Lines.Credit = (itm.PriceAfterVAT / 1.12) * 12%
                            oDocumento.Lines.Debit = 0
                            oDocumento.Lines.Add()
                        Else
                            oDocumento.Lines.Credit = itm.PriceAfterVAT
                            oDocumento.Lines.Debit = 0
                            oDocumento.Lines.Add()
                        End If
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