Imports System.Security.Cryptography
Imports System.Web.Http
Imports System.Web.Mvc
Imports System.Xml
Imports SAPbobsCOM
Imports WS_REST.RespuestaArticulos
Imports WS_REST.FacturaSAP
Namespace Controllers
    Public Class AsientoContableController
        Inherits ApiController

        ' POST: Facturas
        Public Function PostValue(ByVal NombreBaseDatos As String, <FromBody()> ByVal pAsientoContable As AsientoSAP) As RespuestaFacturaSAP
            Dim oRespuesta As New RespuestaFacturaSAP
            Dim oConexion As SAPbobsCOM.Company
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim oDocumento As JournalEntries
            Dim intRespuesta As Integer
            Dim lErrCode As Long
            Dim sErrMsg As String = ""

            Try

                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then
                    oDocumento = Nothing
                    oDocumento = oConexion.GetBusinessObject(BoObjectTypes.oJournalEntries)

                    oDocumento.ReferenceDate = pAsientoContable.ReferenceDate
                    oDocumento.DueDate = pAsientoContable.DueDate
                    oDocumento.TaxDate = pAsientoContable.TaxDate
                    oDocumento.Memo = pAsientoContable.Memo
                    oDocumento.ProjectCode = pAsientoContable.ProjectCode
                    oDocumento.Reference = pAsientoContable.Reference
                    oDocumento.Reference2 = pAsientoContable.Reference2
                    oDocumento.Reference3 = pAsientoContable.Reference3

                    Dim iCont As Integer = 0

                    For Each line In pAsientoContable.LstLinea

                        If iCont > 0 Then
                            oDocumento.Lines.Add()
                        End If

                        oDocumento.Lines.ShortName = Busca_Cuenta(oConexion, line.ShortName)
                        oDocumento.Lines.AccountCode = Busca_Cuenta(oConexion, line.AccountCode)
                        oDocumento.Lines.Debit = line.debit
                        oDocumento.Lines.Credit = line.Credit
                        oDocumento.Lines.LineMemo = line.LineMemo
                        oDocumento.Lines.Reference1 = pAsientoContable.Reference
                        oDocumento.Lines.Reference2 = pAsientoContable.Reference2
                        oDocumento.Lines.CostingCode = line.CostingCode

                        iCont += 1
                    Next

                    'oDocumento.Lines.ShortName = pAsientoContable.CardCode
                    'oDocumento.Lines.Debit = pAsientoContable.DocTotal
                    'oDocumento.Lines.Credit = 0
                    'oDocumento.Lines.Add()


                    'For Each itm In pFactura.LstItems
                    '    Dim Item As SAPbobsCOM.Items = oConexion.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    '    Item.GetByKey(itm.ItemCode)
                    '    Dim ItemGroup As SAPbobsCOM.ItemGroups = oConexion.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    '    ItemGroup.GetByKey(Item.ItemsGroupCode)
                    '    oDocumento.Lines.ShortName = ItemGroup.ExpensesAccount
                    '    If (itm.TaxCode = "IVA") Then
                    '        oDocumento.Lines.Credit = itm.PriceAfterVAT / 1.12
                    '        oDocumento.Lines.Debit = 0
                    '        oDocumento.Lines.Add()
                    '        oDocumento.Lines.ShortName = "_SYS00000000345"
                    '        oDocumento.Lines.Credit = (itm.PriceAfterVAT / 1.12) * 12%
                    '        oDocumento.Lines.Debit = 0
                    '        oDocumento.Lines.Add()
                    '    Else
                    '        oDocumento.Lines.Credit = itm.PriceAfterVAT
                    '        oDocumento.Lines.Debit = 0
                    '        oDocumento.Lines.Add()
                    '    End If
                    'Next

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

        Private Function Busca_Cuenta(ByVal pcnn As SAPbobsCOM.Company, ByVal sFormatCode As String) As String
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim sQuery As String

            Try
                sQuery = "Select AcctCode from OACT where FormatCode = '" & sFormatCode & "'"

                oRecordSet = Nothing
                oRecordSet = pcnn.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRecordSet.DoQuery(sQuery)

                oRecordSet.MoveFirst()

                If Not oRecordSet.EoF Then
                    Return oRecordSet.Fields.Item(0).Value
                Else
                    Return sFormatCode
                End If
            Catch ex As Exception
                Return 0
            Finally
                If System.Runtime.InteropServices.Marshal.IsComObject(oRecordSet) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
            End Try
        End Function

    End Class
End Namespace