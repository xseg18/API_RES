Imports System.Security.Cryptography
Imports System.Web.Http
Imports System.Web.Mvc
Imports System.Xml
Imports SAPbobsCOM
Imports WS_REST.RespuestaMoneda

Namespace Controllers
    Public Class MonedaController
        Inherits ApiController

        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaMoneda
            Dim oRespuesta As New RespuestaMoneda
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New Monedas
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)
                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = "select ""CurrCode"",""CurrName"" from OCRN "
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()
                    oRespuesta = New RespuestaMoneda
                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New Monedas
                            ObjRespuesta.StrCodigoMoneda = oRecordsetConsulta.Fields.Item("CurrCode").Value
                            ObjRespuesta.StrNombreMoneda = oRecordsetConsulta.Fields.Item("CurrName").Value
                            oRespuesta.LstMonedas.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaMoneda
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaMoneda
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function


    End Class
End Namespace