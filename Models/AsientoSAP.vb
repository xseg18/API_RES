Public Class AsientoSAP

    Public Property ReferenceDate
    Public Property DueDate As String
    Public Property TaxDate As String
    Public Property Memo As String
    Public Property ProjectCode As String
    Public Property Reference As String
    Public Property Reference2 As String
    Public Property Reference3 As String

    Public LstLinea As New List(Of DetalleContable)

    Public Class DetalleContable
        Public Property ShortName As String
        Public Property Credit As Double
        Public Property debit As Double
        Public Property AccountCode As String

        Public Property LineMemo As String

        Public Property CostingCode As String

    End Class

End Class
