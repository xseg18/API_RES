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
                    oDocumento.Series = pAsientoContable.Series
                    oDocumento.ProjectCode = pAsientoContable.ProjectCode
                    oDocumento.Reference = pAsientoContable.Reference
                    oDocumento.Reference2 = pAsientoContable.Reference2
                    oDocumento.Reference3 = pAsientoContable.Reference3

                    Dim iCont As Integer = 0

                    For Each line In pAsientoContable.LstLinea

                        If iCont > 0 Then
                            oDocumento.Lines.Add()
                        End If

                        oDocumento.Lines.ShortName = line.AccountCode
                        oDocumento.Lines.AccountCode = line.AccountCode
                        oDocumento.Lines.Debit = line.debit
                        oDocumento.Lines.Credit = line.Credit
                        oDocumento.Lines.LineMemo = line.LineMemo
                        oDocumento.Lines.Reference1 = pAsientoContable.Reference
                        oDocumento.Lines.Reference2 = pAsientoContable.Reference2
                        oDocumento.Lines.CostingCode = line.CostingCode

                        iCont += 1
                    Next



                    intRespuesta = oDocumento.Add()

                    If intRespuesta = 0 Then
                        oRespuesta.Mensaje = "Asiento Creado Correctamente"
                        oRespuesta.Estado = True
                        oRespuesta.IDAsientoSAP = oConexion.GetNewObjectKey
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