Public Class FacturaSAP

    Public Property CardCode As String
    Public Property CardName As String
    Public Property TaxDate As String
    Public Property DocDate As String
    Public Property DocDueDate As String
    Public Property Comments As String
    Public Property Address As String
    Public Property DocTotal As String

    Public LstItems As New List(Of items)

    Public Class items
        Public Property ItemCode As String
        Public Property ItemDescription As String
        Public Property TaxCode As String
        Public Property Quantity As Double

        Public Property PriceAfterVAT As Double

    End Class

End Class
