Imports System.Web.Http
Imports System.Web.Mvc
Imports SAPbobsCOM
Imports WS_REST.RespuestaAlmacenes
Imports WS_REST.RespuestaArticulos

Namespace Controllers
    Public Class AlmacenesController
        Inherits ApiController

        ' GET: Almacenes
        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaAlmacenes
            Dim oRespuesta As New RespuestaAlmacenes
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New Almacen
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset

            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)

                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = " select ""WhsCode"",""WhsName"" from OWHS "
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()

                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New Almacen
                            ObjRespuesta.WhsCode = oRecordsetConsulta.Fields.Item("WhsCode").Value
                            ObjRespuesta.WhsName = oRecordsetConsulta.Fields.Item("ItemName").Value
                            oRespuesta.LstAlmacenes.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaAlmacenes
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaAlmacenes
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function


    End Class
End Namespace