Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaCentrosCosto
Imports WS_REST.RespuestaMoneda

Namespace Controllers
    Public Class CentrosCostoController
        Inherits ApiController
        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaCentrosCosto
            Dim oRespuesta As New RespuestaCentrosCosto
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New CentroCosto
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)
                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = "select ""PrcCode"",""PrcName"" from OPRC "
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()
                    oRespuesta = New RespuestaCentrosCosto
                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New CentroCosto
                            ObjRespuesta.PrcCode = oRecordsetConsulta.Fields.Item("PrcCode").Value
                            ObjRespuesta.PrcName = oRecordsetConsulta.Fields.Item("PrcName").Value
                            oRespuesta.LstCentroCosto.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaCentrosCosto
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaCentrosCosto
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function
    End Class
End Namespace