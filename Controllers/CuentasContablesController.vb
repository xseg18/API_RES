Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaCuentasContables
Imports WS_REST.RespuestaCuentasPropias

Namespace Controllers
    Public Class CuentasContablesController
        Inherits ApiController

        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaCuentasContables
            Dim oRespuesta As New RespuestaCuentasContables
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New CuentaContable
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)
                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = "select ""AcctCode"", ""AcctName"", ""CurrTotal"" from OACT"
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()
                    oRespuesta = New RespuestaCuentasContables
                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New CuentaContable
                            ObjRespuesta.AcctCode = oRecordsetConsulta.Fields.Item("AcctCode").Value
                            ObjRespuesta.AcctName = oRecordsetConsulta.Fields.Item("AcctName").Value
                            ObjRespuesta.CurrTotal = oRecordsetConsulta.Fields.Item("CurrTotal").Value
                            oRespuesta.LstCuentaContable.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaCuentasContables
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaCuentasContables
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function
    End Class
End Namespace