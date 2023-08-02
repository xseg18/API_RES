Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaBanco

Namespace Controllers
    Public Class PeriodoContableController
        Inherits ApiController

        ' GET: Bancos
        Public Function GetValue(ByVal NombreBaseDatos As String, ByVal FechaPeriodo As String) As RespuestaPeriodoContable
            Dim oRespuesta As New RespuestaPeriodoContable
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset

            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = " select ""Code"",""Name"",case when ""PeriodStat"" = 'Y' then 'BLOQUEADO' else 'DESBLOQUEADO' END as Estado  from OFPR where ""Code"" = '" & CDate(FechaPeriodo).ToString("yyyy-MM") & "'"
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()

                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.Mensaje = "NO SE ENCUENTRAN REGISTROS"
                        oRespuesta.EstadoRespuesta = False
                    Else
                        oRespuesta.Mensaje = "OK"
                        oRespuesta.EstadoRespuesta = True

                        If oRecordsetConsulta.Fields.Item("Estado").Value = "BLOQUEADO" Then
                            oRespuesta.PeriodoEstado = False
                        Else
                            oRespuesta.PeriodoEstado = True
                        End If

                        oRespuesta.EstadoPeriodo = oRecordsetConsulta.Fields.Item("Estado").Value
                        oRespuesta.Code = oRecordsetConsulta.Fields.Item("Code").Value
                        oRespuesta.Name = oRecordsetConsulta.Fields.Item("Name").Value

                    End If
                        oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaPeriodoContable
                    oRespuesta.Mensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.EstadoRespuesta = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaPeriodoContable
                oRespuesta.Mensaje = "Error De Servidor: " & ex.Message
                oRespuesta.EstadoRespuesta = False
                Return oRespuesta
            End Try
        End Function

    End Class
End Namespace