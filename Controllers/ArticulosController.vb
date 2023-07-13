Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaArticulos

Namespace Controllers
    Public Class ArticulosController
        Inherits ApiController

        ' GET: Articulos
        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaArticulos
            Dim oRespuesta As New RespuestaArticulos
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New Articulos
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset

            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = " select ""ItemCode"",""ItemName"",""FrgnName"" from OITM "
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()

                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New Articulos
                            ObjRespuesta.ItemCode = oRecordsetConsulta.Fields.Item("ItemCode").Value
                            ObjRespuesta.ItemName = oRecordsetConsulta.Fields.Item("ItemName").Value
                            ObjRespuesta.FrgnName = oRecordsetConsulta.Fields.Item("FrgnName").Value
                            oRespuesta.LstArticulos.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaArticulos
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaArticulos
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function
    End Class
End Namespace