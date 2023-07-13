Imports System.Security.Cryptography
Imports System.Xml

Module Funciones
    Dim strruta As String = "C:\WS_REST"
    Public Function FncDesencriptar(ByVal pStrInformacionADesencriptar As String) As String
        Dim IV() As Byte = ASCIIEncoding.ASCII.GetBytes("1Inf1N$t") 'La clave debe ser de 8 caracteres
        Dim EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
        Dim buffer() As Byte = Convert.FromBase64String(pStrInformacionADesencriptar)
        Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        des.Key = EncryptionKey
        des.IV = IV
        Return Encoding.UTF8.GetString(des.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
    End Function
    Public Function FncSAPConexion(ByVal BaseDatosSAP As String, ByRef IntCodigoError As Integer, ByRef StrMensaje As String) As SAPbobsCOM.Company
        Try
            Dim oCnn As SAPbobsCOM.Company
            Dim documento As XmlDocument = New XmlDocument()
            Dim IntRespuesta As Integer
            oCnn = New SAPbobsCOM.Company

            documento.Load(strruta & "\oConexion.xml")
            Dim listaConexion As XmlNodeList = documento.GetElementsByTagName("ConexionesSAP")
            Dim conexionSAP As XmlNodeList = (CType(listaConexion(0), XmlElement)).GetElementsByTagName("SAPBD")

            For Each conexion As XmlElement In conexionSAP
                Dim sServer As XmlNodeList = conexion.GetElementsByTagName("Server")
                Dim cCompaniaDB As XmlNodeList = conexion.GetElementsByTagName("CompaniaDB")
                Dim uUsuario As XmlNodeList = conexion.GetElementsByTagName("Usuario")
                Dim cContrasena As XmlNodeList = conexion.GetElementsByTagName("Contrasena")
                Dim uUsuarioDB As XmlNodeList = conexion.GetElementsByTagName("UsuarioDB")
                Dim cContraenaDB As XmlNodeList = conexion.GetElementsByTagName("ContrasenaDB")
                Dim lLicenciaServer As XmlNodeList = conexion.GetElementsByTagName("LicenciaServer")
                Dim sServerType As XmlNodeList = conexion.GetElementsByTagName("TypeServer")
                Dim sSLD As XmlNodeList = conexion.GetElementsByTagName("SLD")

                If Not BaseDatosSAP = FncDesencriptar(cCompaniaDB(0).InnerText) Then Continue For

                oCnn.Server = FncDesencriptar(sServer(0).InnerText)
                oCnn.LicenseServer = FncDesencriptar(lLicenciaServer(0).InnerText)
                oCnn.DbUserName = FncDesencriptar(uUsuarioDB(0).InnerText)
                oCnn.DbPassword = FncDesencriptar(cContraenaDB(0).InnerText)
                oCnn.CompanyDB = FncDesencriptar(cCompaniaDB(0).InnerText)
                oCnn.UserName = FncDesencriptar(uUsuario(0).InnerText)
                oCnn.Password = FncDesencriptar(cContrasena(0).InnerText)
                oCnn.SLDServer = (sSLD(0).InnerText)
                Select Case sServerType(0).InnerText
                    Case "dst_HANADB"
                        oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                    Case "MSSQL_2017"
                        oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017
                    Case "MSSQL_2016"
                        oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                    Case "MSSQL_2014"
                        oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                    Case "MSSQL_2012"
                        oCnn.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                End Select

            Next

            IntRespuesta = oCnn.Connect

            If IntRespuesta <> 0 Then
                oCnn.GetLastError(IntCodigoError, StrMensaje)
            End If

            Return oCnn

        Catch ex As Exception
            StrMensaje = ex.Message
        End Try

    End Function
End Module
