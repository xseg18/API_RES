Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaTasaCambio

Namespace Controllers
    Public Class TasaCambioController
        Inherits ApiController

        ' GET: TasaCambio
        Public Function GetValue(ByVal NombreBaseDatos As String, ByVal FechaTasaCambio As String) As RespuestaTasaCambio
            Dim oRespuesta As New RespuestaTasaCambio
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New TasaCambio
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)
                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = "select ""RateDate"",""Currency"",""Rate"" from ORTT 
                                Where Cast(""RateDate"" as date) = '" & FechaTasaCambio & "'  "
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()
                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New TasaCambio
                            ObjRespuesta.RateDate = oRecordsetConsulta.Fields.Item("RateDate").Value
                            ObjRespuesta.currency = oRecordsetConsulta.Fields.Item("currency").Value
                            ObjRespuesta.Rate = oRecordsetConsulta.Fields.Item("Rate").Value
                            oRespuesta.LstTasaCambio.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaTasaCambio
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaTasaCambio
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function
    End Class
End Namespace