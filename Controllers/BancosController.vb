Imports System.Web.Mvc

Namespace Controllers
    Public Class BancosController
        Inherits Controller

        ' GET: Bancos
        Function Index() As ActionResult
            Return View()
        End Function
    End Class
End Namespace