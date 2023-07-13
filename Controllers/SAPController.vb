Imports System.IO
Imports System.Security.Cryptography
Imports System.Web.Http
Imports System.Web.Mvc
Imports System.Xml
Imports SAPbobsCOM

Namespace Controllers

    Public Class SAPController
        Inherits ApiController
        Dim strruta As String = "C:\WS_REST"
        Shared oConexion As Dictionary(Of Integer, Conexion) = New Dictionary(Of Integer, Conexion)()

        Public Function PostValue(<FromBody()> ByVal pConexion As Conexion) As RespuestaConexion
            Dim oRespuesta As New RespuestaConexion
            Try

                Dim IntRespuesta As Integer
                Dim IntCodigoError As Integer
                Dim StrMensaje As String = String.Empty
                Dim oCnn As CompanyClass
                oCnn = New CompanyClass
                Dim sXML As String = ""

                oCnn.Server = pConexion.StrServer
                oCnn.LicenseServer = pConexion.StrLicenseServer
                oCnn.DbUserName = pConexion.StrDbUserName
                oCnn.DbPassword = pConexion.StrDbPassword
                oCnn.CompanyDB = pConexion.StrCompanyDB
                oCnn.UserName = pConexion.StrUserName
                oCnn.Password = pConexion.StrPassword
                oCnn.SLDServer = pConexion.StrSLD

                If pConexion.StrDbServerType = "dst_HANADB" Then
                    oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                ElseIf pConexion.StrDbServerType = "MSSQL_2017" Then
                    oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017
                ElseIf pConexion.StrDbServerType = "MSSQL_2016" Then
                    oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                ElseIf pConexion.StrDbServerType = "MSSQL_2012" Then
                    oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                End If

                IntRespuesta = oCnn.Connect

                If IntRespuesta <> 0 Then
                    oCnn.GetLastError(IntCodigoError, StrMensaje)
                    oRespuesta.Estatus = False
                    oRespuesta.Mensaje = StrMensaje
                Else
                    Dim CreateXML As FileStream

                    oRespuesta.Estatus = True
                    oRespuesta.Mensaje = "CONECTADO CORRECTAMENTE"

                    If Not Directory.Exists(strruta) Then My.Computer.FileSystem.CreateDirectory(strruta)

                    If Not File.Exists(strruta & "\oConexion.xml") Then
                        CreateXML = File.Create(strruta & "\oConexion.xml")
                        sXML = ("<?xml version='1.0' encoding='utf-8'?>")
                        sXML &= ("<ConexionesSAP>")
                        sXML &= ("<SAPBD>")
                        sXML &= ("<CompaniaDB>" & FncEncriptar(pConexion.StrCompanyDB) & "</CompaniaDB>")
                        sXML &= ("<Server>" & FncEncriptar(pConexion.StrServer) & "</Server>")
                        sXML &= ("<Usuario>" & FncEncriptar(pConexion.StrUserName) & "</Usuario>")
                        sXML &= ("<Contrasena>" & FncEncriptar(pConexion.StrPassword) & "</Contrasena>")
                        sXML &= ("<UsuarioDB>" & FncEncriptar(pConexion.StrDbUserName) & "</UsuarioDB>")
                        sXML &= ("<ContrasenaDB>" & FncEncriptar(pConexion.StrDbPassword) & "</ContrasenaDB>")
                        sXML &= ("<LicenciaServer>" & FncEncriptar(pConexion.StrLicenseServer) & "</LicenciaServer>")
                        sXML &= ("<TypeServer>" & pConexion.StrDbServerType & "</TypeServer>")
                        sXML &= ("<SLD>" & pConexion.StrSLD & "</SLD>")
                        sXML &= ("</SAPBD>")
                        sXML &= ("</ConexionesSAP>")

                        Dim info As Byte() = New UTF8Encoding(True).GetBytes(sXML)
                        CreateXML.Write(info, 0, info.Length)
                        CreateXML.Close()
                    Else
                        Dim documento As XmlDocument = New XmlDocument()
                        documento.Load(strruta & "\oConexion.xml")
                        Dim listaConexion As XmlNodeList = documento.GetElementsByTagName("ConexionesSAP")
                        Dim conexionSAP As XmlNodeList = (CType(listaConexion(0), XmlElement)).GetElementsByTagName("SAPBD")
                        Dim ObjConexion As New ConexionSAP

                        For Each conexion As XmlElement In conexionSAP
                            Dim ObjCredenciales As New ConexionSAP.Credenciales
                            Dim sServer As XmlNodeList = conexion.GetElementsByTagName("Server")
                            Dim cCompaniaDB As XmlNodeList = conexion.GetElementsByTagName("CompaniaDB")
                            Dim uUsuario As XmlNodeList = conexion.GetElementsByTagName("Usuario")
                            Dim cContrasena As XmlNodeList = conexion.GetElementsByTagName("Contrasena")
                            Dim uUsuarioDB As XmlNodeList = conexion.GetElementsByTagName("UsuarioDB")
                            Dim cContraenaDB As XmlNodeList = conexion.GetElementsByTagName("ContrasenaDB")
                            Dim lLicenciaServer As XmlNodeList = conexion.GetElementsByTagName("LicenciaServer")
                            Dim sServerType As XmlNodeList = conexion.GetElementsByTagName("TypeServer")
                            Dim sSLD As XmlNodeList = conexion.GetElementsByTagName("SLD")
                            ObjCredenciales.StrServer = sServer(0).InnerText
                            ObjCredenciales.StrLicenseServer = lLicenciaServer(0).InnerText
                            ObjCredenciales.StrDbUserName = uUsuarioDB(0).InnerText
                            ObjCredenciales.StrDbPassword = cContraenaDB(0).InnerText
                            ObjCredenciales.StrCompanyDB = cCompaniaDB(0).InnerText
                            ObjCredenciales.StrUserName = uUsuario(0).InnerText
                            ObjCredenciales.StrPassword = cContrasena(0).InnerText
                            ObjCredenciales.StrDbServerType = sServerType(0).InnerText
                            ObjCredenciales.StrSLD = sSLD(0).InnerText
                            ObjConexion.LstConexion.Add(ObjCredenciales)
                        Next
                        File.Delete(strruta & "\oConexion.xml")
                        CreateXML = File.Create(strruta & "\oConexion.xml")
                        sXML = ("<?xml version='1.0' encoding='utf-8'?>")
                        sXML &= ("<ConexionesSAP>")
                        sXML &= ("<SAPBD>")
                        sXML &= ("<CompaniaDB>" & FncEncriptar(pConexion.StrCompanyDB) & "</CompaniaDB>")
                        sXML &= ("<Server>" & FncEncriptar(pConexion.StrServer) & "</Server>")
                        sXML &= ("<Usuario>" & FncEncriptar(pConexion.StrUserName) & "</Usuario>")
                        sXML &= ("<Contrasena>" & FncEncriptar(pConexion.StrPassword) & "</Contrasena>")
                        sXML &= ("<UsuarioDB>" & FncEncriptar(pConexion.StrDbUserName) & "</UsuarioDB>")
                        sXML &= ("<ContrasenaDB>" & FncEncriptar(pConexion.StrDbPassword) & "</ContrasenaDB>")
                        sXML &= ("<LicenciaServer>" & FncEncriptar(pConexion.StrLicenseServer) & "</LicenciaServer>")
                        sXML &= ("<TypeServer>" & pConexion.StrDbServerType & "</TypeServer>")
                        sXML &= ("<SLD>" & FncEncriptar(pConexion.StrSLD) & "</SLD>")
                        sXML &= ("</SAPBD>")
                        For i As Integer = 0 To ObjConexion.LstConexion.Count - 1
                            sXML &= ("<SAPBD>")
                            sXML &= ("<CompaniaDB>" & FncEncriptar(ObjConexion.LstConexion(i).StrCompanyDB) & "</CompaniaDB>")
                            sXML &= ("<Server>" & FncEncriptar(ObjConexion.LstConexion(i).StrServer) & "</Server>")
                            sXML &= ("<Usuario>" & FncEncriptar(ObjConexion.LstConexion(i).StrUserName) & "</Usuario>")
                            sXML &= ("<Contrasena>" & FncEncriptar(ObjConexion.LstConexion(i).StrPassword) & "</Contrasena>")
                            sXML &= ("<UsuarioDB>" & FncEncriptar(ObjConexion.LstConexion(i).StrDbUserName) & "</UsuarioDB>")
                            sXML &= ("<ContrasenaDB>" & FncEncriptar(ObjConexion.LstConexion(i).StrDbPassword) & "</ContrasenaDB>")
                            sXML &= ("<LicenciaServer>" & FncEncriptar(ObjConexion.LstConexion(i).StrLicenseServer) & "</LicenciaServer>")
                            sXML &= ("<TypeServer>" & (ObjConexion.LstConexion(i).StrDbServerType) & "</TypeServer>")
                            sXML &= ("<SLD>" & FncEncriptar(ObjConexion.LstConexion(i).StrSLD) & "</SLD>")
                            sXML &= ("</SAPBD>")
                        Next
                        sXML &= ("</ConexionesSAP>")

                        Dim info As Byte() = New UTF8Encoding(True).GetBytes(sXML)
                        CreateXML.Write(info, 0, info.Length)
                        CreateXML.Close()
                    End If

                End If
                oCnn.Disconnect()
                Return oRespuesta

            Catch ex As Exception
                oRespuesta.Estatus = False
                oRespuesta.Mensaje = ex.Message
                Return oRespuesta
            End Try

        End Function

        Private Function FncEncriptar(ByVal pStrInformacionAEncriptar As String) As String
            Dim IV() As Byte = ASCIIEncoding.ASCII.GetBytes("1Inf1N$t")
            Dim EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5")
            Dim buffer() As Byte = Encoding.UTF8.GetBytes(pStrInformacionAEncriptar)
            Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
            des.Key = EncryptionKey
            des.IV = IV
            Return Convert.ToBase64String(des.CreateEncryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
        End Function
        Private Function FncDesencriptar(ByVal pStrInformacionADesencriptar As String) As String
            Dim IV() As Byte = ASCIIEncoding.ASCII.GetBytes("1Inf1N$t") 'La clave debe ser de 8 caracteres
            Dim EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
            Dim buffer() As Byte = Convert.FromBase64String(pStrInformacionADesencriptar)
            Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
            des.Key = EncryptionKey
            des.IV = IV
            Return Encoding.UTF8.GetString(des.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
        End Function

    End Class
End Namespace