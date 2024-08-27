Public Class ClsPurchaseOrderResponse
    Public Property Value As List(Of ClsPurchaseOrder_JSON)

End Class

Public Class ClsPurchaseOrder_JSON
    Public Property DocEntry As Integer
    Public Property DocNum As Integer
    Public Property DocType As String
    Public Property DocDate As String
    Public Property DocDueDate As Date?
    Public Property CardCode As String
    Public Property CardName As String
    Public Property Address As String
    Public Property DocumentStatus As String
    Public Property Cancelled As String
    Public Property DocumentLines As List(Of ClsPurchaseOrder_DocumentLine_JSON)
End Class
Public Class ClsPurchaseOrder_DocumentLine_JSON
    Public Property LineNum As Integer
    Public Property ItemCode As String
    Public Property ItemDescription As String
    Public Property Quantity As Decimal
End Class
