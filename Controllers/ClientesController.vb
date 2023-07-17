Imports System.Web.Http
Imports SAPbobsCOM
Imports WS_REST.RespuestaClientes
Imports WS_REST.RespuestaMoneda

Namespace Controllers
    Public Class ClientesController
        Inherits ApiController

        Public Function GetValue(ByVal NombreBaseDatos As String) As RespuestaClientes
            Dim oRespuesta As New RespuestaClientes
            Dim oConexion As SAPbobsCOM.Company
            Dim StrQuery As String = String.Empty
            Dim lIntCodigoError As Integer = 0
            Dim lStrTextoError As String = String.Empty
            Dim ObjRespuesta As New Cliente
            Dim oRecordsetConsulta As SAPbobsCOM.Recordset
            Try
                oConexion = Funciones.FncSAPConexion(NombreBaseDatos, lIntCodigoError, lStrTextoError)
                If lStrTextoError.Length = 0 Then

                    oRecordsetConsulta = Nothing
                    oRecordsetConsulta = oConexion.GetBusinessObject(BoObjectTypes.BoRecordset)
                    StrQuery = "select A.""CardCode"", A.""CardName"", A.""LicTradNum"",B.""CardFName"", B.""Street"", B.""City"", C.""Name"" as Departamento, D.""Name"" as Pais from OCRD A 
inner join CRD1 B on A.""CardCode"" = B.""CardCode"" 
inner join OCST C on C.""Code"" = B.""State"" 
inner join OCRY D on D.""Code"" = B.""Country"""
                    oRecordsetConsulta.DoQuery(StrQuery)
                    oRecordsetConsulta.MoveFirst()
                    oRespuesta = New RespuestaClientes
                    If oRecordsetConsulta.RecordCount = 0 Then
                        oRespuesta.BlnEstado = False
                        oRespuesta.StrMensaje = "NO SE ENCUENTRAN REGISTROS"
                    Else
                        oRespuesta.BlnEstado = True
                        oRespuesta.StrMensaje = "OK"
                        For i As Integer = 0 To oRecordsetConsulta.RecordCount - 1
                            ObjRespuesta = New Cliente
                            ObjRespuesta.ID = oRecordsetConsulta.Fields.Item("CardCode").Value
                            ObjRespuesta.Name = oRecordsetConsulta.Fields.Item("CardName").Value
                            ObjRespuesta.CommercialName = oRecordsetConsulta.Fields.Item("CardFName").Value
                            ObjRespuesta.NIT = oRecordsetConsulta.Fields.Item("LicTradNum").Value
                            ObjRespuesta.Address = oRecordsetConsulta.Fields.Item("Street").Value + ", " + oRecordsetConsulta.Fields.Item("City").Value + ", " + oRecordsetConsulta.Fields.Item("Departamento").Value + ", " + oRecordsetConsulta.Fields.Item("Pais").Value
                            oRespuesta.LstClientes.Add(ObjRespuesta)
                            oRecordsetConsulta.MoveNext()
                        Next
                    End If
                    oConexion.Disconnect()
                    Return oRespuesta
                Else
                    oRespuesta = New RespuestaClientes
                    oRespuesta.StrMensaje = "Error Conexion: " & lIntCodigoError & " - " & lStrTextoError
                    oRespuesta.BlnEstado = False
                    Return oRespuesta
                End If
            Catch ex As Exception
                oRespuesta = New RespuestaClientes
                oRespuesta.StrMensaje = "Error De Servidor: " & ex.Message
                oRespuesta.BlnEstado = False
                Return oRespuesta
            End Try
        End Function

    End Class
End Namespace