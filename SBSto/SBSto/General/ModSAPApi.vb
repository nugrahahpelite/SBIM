Imports System.Net
Imports System.Net.Security
Imports RestSharp
Imports Newtonsoft.Json
Imports System.Security.Cryptography.X509Certificates
Imports System.Data.SqlClient

Module ModSAPApi
    Public Structure stuSAPApi_Modul
        Const Login = "Login"
        Const PurchaseOrder = "Purchase Order"
    End Structure

    Public Structure stuSAPApi_Field
        Const Method = "SAPApi_Method"
        Const EndPoint = "SAPApi_EndPoint"
        Const Link = "SAPApi_Link"
    End Structure

    Dim vsB1SESSION As RestResponseCookie
    Dim vsROUTEID As RestResponseCookie
    Private Function ValidateServerCertificate(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
        ' Bypass SSL/TLS verification (always return true)
        Return True
    End Function

    Public Sub pbuSAP_GetApiLink(vriTextStream As Scripting.TextStream, vriDtbApi As DataTable, vriModul As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<------------------------pbuSAP_GetApiLink")
        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("vriModul = " & vriModul)

        Dim vnQuery As String
        vnQuery = "Select * From Sys_SAPApi_MA with(nolock) Where SAPApi_Modul='" & vriModul & "'"
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("vnQuery = " & vnQuery)
        pbuFillDtbSQLTrans(vriDtbApi, vnQuery, vriSQLConn, vriSQLTrans)

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("<<========================pbuSAP_GetApiLink")
        vriTextStream.WriteLine("")
    End Sub

    Public Function fbuSAP_Login(vriTextStream As Scripting.TextStream, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction, ByRef vroMessage As String) As Boolean
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<------------------------fbuSAP_Login")
        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        Try
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)

            Dim vnDtbApi As New DataTable
            pbuSAP_GetApiLink(vriTextStream, vnDtbApi, stuSAPApi_Modul.Login, vriSQLConn, vriSQLTrans)

            Dim vnClient_Login As RestClient, vnAPIParams_Login As String, vnResponse_Login As IRestResponse
            'vnClient_Login = New RestClient("https ://sap.sumberberkat.com:51100/b1s/v1/Login")
            vnClient_Login = New RestClient(vnDtbApi.Rows(0).Item(stuSAPApi_Field.Link))
            vnClient_Login.Timeout = -1

            Dim vnRequest_Login = New RestRequest(Method.POST)
            vnAPIParams_Login = "{"
            vnAPIParams_Login += " ""CompanyDB"":""BAPJTESTING"","
            vnAPIParams_Login += " ""UserName"":""WMS01"","
            vnAPIParams_Login += " ""Password"":""B@pJ#$%@2023"""
            vnAPIParams_Login += "}"
            vnRequest_Login.AddHeader("Content-Type", "application/json")
            vnRequest_Login.AddParameter("application/json", vnAPIParams_Login, ParameterType.RequestBody)
            vnResponse_Login = vnClient_Login.Execute(vnRequest_Login)

            vsB1SESSION = vnResponse_Login.Cookies.FirstOrDefault(Function(c) c.Name = "B1SESSION")
            vsROUTEID = vnResponse_Login.Cookies.FirstOrDefault(Function(c) c.Name = "ROUTEID")

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnResponse_Login.Content")
            vriTextStream.WriteLine(vnResponse_Login.Content)
            vriTextStream.WriteLine("")

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Return True")
            vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vriTextStream.WriteLine("<<========================fbuSAP_Login")
            vriTextStream.WriteLine("")
            vroMessage = ""

            Return True
        Catch ex As Exception
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("ERROR")
            vriTextStream.WriteLine(ex.Message)
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Return False")
            vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vriTextStream.WriteLine("<<========================fbuSAP_Login")
            vriTextStream.WriteLine("")
            vroMessage = ex.Message

            Return False
        End Try
    End Function

    Public Function fbuSAPApi_GetPurchaseOrder(vriTextStream As Scripting.TextStream, vriTextStream_Data As Scripting.TextStream, vriUserOID As String, vriHOID As String, vriCompanyCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction, ByRef vroMessage As String) As Boolean
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<------------------------fbuSAPApi_GetPurchaseOrder")
        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))

        Dim vnMessage As String = ""

        If fbuSAP_Login(vriTextStream, vriSQLConn, vriSQLTrans, vnMessage) Then
            Dim vnDtbApi As New DataTable
            pbuSAP_GetApiLink(vriTextStream, vnDtbApi, stuSAPApi_Modul.PurchaseOrder, vriSQLConn, vriSQLTrans)

            Dim vnGetDate As String = fbuGetDateNowSQLTrans(vriSQLConn, vriSQLTrans)

            For vn = 0 To 1000
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("************************************************************************************")
                vriTextStream.WriteLine("skip - vn = " & vn)

                Dim vnClient_PO As RestClient, vnResponse_PO As IRestResponse
                'vnClient_PO = New RestClient("https ://sap.sumberberkat.com:51100/b1s/v1/PurchaseOrders")
                vnClient_PO = New RestClient(vnDtbApi.Rows(0).Item(stuSAPApi_Field.Link) & "?$skip=" & vn * 20)
                vnClient_PO.Timeout = -1

                Dim vnRequest_PO = New RestRequest(Method.GET)
                vnRequest_PO.AddHeader("Content-Type", "application/json")

                vnRequest_PO.AddCookie("B1SESSION", vsB1SESSION.Value)
                vnRequest_PO.AddCookie("ROUTEID", vsROUTEID.Value)

                vnResponse_PO = vnClient_PO.Execute(vnRequest_PO)

                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnResponse_PO.Content")
                vriTextStream.WriteLine(vnResponse_PO.Content)
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("")

                vriTextStream_Data.WriteLine("")
                vriTextStream_Data.WriteLine(vnResponse_PO.Content)

                Dim vnClsPO As ClsPurchaseOrderResponse = JsonConvert.DeserializeObject(Of ClsPurchaseOrderResponse)(vnResponse_PO.Content)

                Dim vnQuery As String
                Dim vnDataCount As Integer = 0

                For Each vnH As ClsPurchaseOrder_JSON In vnClsPO.Value
                    vnDataCount = vnDataCount + 1
                    Dim vnHDocEntry As String = vnH.DocEntry
                    Dim vnHDocNum As String = vnH.DocNum
                    Dim vnHDocType As String = vnH.DocType
                    Dim vnHDocDate As String = vnH.DocDate
                    Dim vnHDocDueDate As String = vnH.DocDueDate
                    Dim vnHCardCode As String = vnH.CardCode
                    Dim vnHCardName As String = vnH.CardName
                    Dim vnHAddress As String = vnH.Address
                    Dim vnHDocumentStatus As String = vnH.DocumentStatus
                    Dim vnHCancelled As String = vnH.Cancelled
                    Dim vnHGDGCode As String = ""
                    vriTextStream.WriteLine("")
                    vriTextStream.WriteLine("=======")

                    vriTextStream.WriteLine("vnHDocEntry=" & vnHDocEntry)
                    vriTextStream.WriteLine("vnHDocNum=" & vnHDocNum)
                    vriTextStream.WriteLine("vnHDocType=" & vnHDocType)
                    vriTextStream.WriteLine("vnHDocDate=" & vnHDocDate)
                    vriTextStream.WriteLine("vnHDocDueDate=" & vnHDocDueDate)
                    vriTextStream.WriteLine("vnHCardCode=" & vnHCardCode)
                    vriTextStream.WriteLine("vnHCardName=" & vnHCardName)
                    vriTextStream.WriteLine("vnHAddress=" & vnHAddress)
                    vriTextStream.WriteLine("vnHDocumentStatus=" & vnHDocumentStatus)
                    vriTextStream.WriteLine("vnHCancelled=" & vnHCancelled)

                    vriTextStream.WriteLine("-------")
                    For Each vnD As ClsPurchaseOrder_DocumentLine_JSON In vnH.DocumentLines
                        Dim vnDLineNum As String = vnD.LineNum
                        Dim vnDItemCode As String = vnD.ItemCode
                        Dim vnDItemDescription As String = vnD.ItemDescription
                        Dim vnDQuantity As String = vnD.Quantity

                        vriTextStream.WriteLine("vnDLineNum=" & vnDLineNum)
                        vriTextStream.WriteLine("vnDItemCode=" & vnDItemCode)
                        vriTextStream.WriteLine("vnDItemDescription=" & vnDItemDescription)
                        vriTextStream.WriteLine("vnDQuantity=" & vnDQuantity)

                        '<---23 Oct 2023 Sampai sini
                        vnQuery = "Insert into #Sys_SsoPO_Temp"
                        vnQuery += vbCrLf & "(CompanyCode,PO_NO,JURNAL,PO_DATE,SUB,BRG_ORIG,BRG,GDGCODE,"
                        vnQuery += vbCrLf & "QTY,NAMA_SUPPLIER,NAMA_BARANG,JOBNAME,"
                        vnQuery += vbCrLf & "SAP_DocEntry,SAP_DocType,SAP_DocumentStatus,SAP_Cancelled,SAP_LineNum,"
                        vnQuery += vbCrLf & "POFileXlsOID,UploadSourceOID,UploadDatetime"
                        vnQuery += vbCrLf & ")"

                        vnQuery += vbCrLf & "Select '" & vriCompanyCode & "'CompanyCode,'" & vnHDocNum & "'PO_NO,''JURNAL,'" & vnHDocDate & "'PO_DATE,'" & vnHCardCode & "'SUB,'" & vnDItemCode & "'BRG_ORIG,'" & vnDItemCode & "'BRG,'" & vnHGDGCode & "'GDGCODE,"
                        vnQuery += vbCrLf & "'" & vnDQuantity & "'QTY,'" & vnHCardName & "'NAMA_SUPPLIER,''NAMA_BARANG,''JOBNAME,"
                        vnQuery += vbCrLf & vnHDocEntry & " SAP_DocEntry,'" & vnHDocType & "'SAP_DocType,'" & vnHDocumentStatus & "'SAP_DocumentStatus,'" & vnHCancelled & "'SAP_Cancelled," & vnDLineNum & " SAP_LineNum,"
                        vnQuery += vbCrLf & vriHOID & "," & enuUploadSource.SAP_Api & ",'" & vnGetDate & "'"
                        vriTextStream.WriteLine("vnQuery")
                        vriTextStream.WriteLine(vnQuery)
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                        '<<==
                    Next
                Next

                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("************************************************************************************")
                vriTextStream.WriteLine("************************************************************************************")

                If vnDataCount = 0 Then
                    fbuSAP_Logout(vriTextStream)

                    vriTextStream.WriteLine("")
                    vriTextStream.WriteLine("Return True")
                    vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                    vriTextStream.WriteLine("<<===========================fbuSAPApi_GetPurchaseOrder")
                    vriTextStream.WriteLine("")
                    Return True
                End If
            Next
        Else
            vroMessage = vnMessage

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Return False")
            vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vriTextStream.WriteLine("<<===========================fbuSAPApi_GetPurchaseOrder")
            vriTextStream.WriteLine("")
            Return False
        End If
    End Function

    Public Function fbuSAPApi_GetPurchaseOrder_20231023_Orig_Bef_With_Skip(vriTextStream As Scripting.TextStream, vriTextStream_Data As Scripting.TextStream, vriUserOID As String, vriHOID As String, vriCompanyCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction, ByRef vroMessage As String) As Boolean
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<------------------------fbuSAPApi_GetPurchaseOrder")
        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))

        Dim vnMessage As String = ""

        If fbuSAP_Login(vriTextStream, vriSQLConn, vriSQLTrans, vnMessage) Then
            Dim vnDtbApi As New DataTable
            pbuSAP_GetApiLink(vriTextStream, vnDtbApi, stuSAPApi_Modul.PurchaseOrder, vriSQLConn, vriSQLTrans)

            Dim vnClient_PO As RestClient, vnResponse_PO As IRestResponse
            'vnClient_PO = New RestClient("https ://sap.sumberberkat.com:51100/b1s/v1/PurchaseOrders")
            vnClient_PO = New RestClient(vnDtbApi.Rows(0).Item(stuSAPApi_Field.Link))
            vnClient_PO.Timeout = -1

            Dim vnRequest_PO = New RestRequest(Method.GET)
            vnRequest_PO.AddHeader("Content-Type", "application/json")

            vnRequest_PO.AddCookie("B1SESSION", vsB1SESSION.Value)
            vnRequest_PO.AddCookie("ROUTEID", vsROUTEID.Value)

            vnResponse_PO = vnClient_PO.Execute(vnRequest_PO)

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnResponse_PO.Content")
            vriTextStream.WriteLine(vnResponse_PO.Content)
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("")

            vriTextStream_Data.WriteLine("")
            vriTextStream_Data.WriteLine(vnResponse_PO.Content)

            Dim vnClsPO As ClsPurchaseOrderResponse = JsonConvert.DeserializeObject(Of ClsPurchaseOrderResponse)(vnResponse_PO.Content)

            Dim vnGetDate As String = fbuGetDateNowSQLTrans(vriSQLConn, vriSQLTrans)

            Dim vnQuery As String
            For Each vnH As ClsPurchaseOrder_JSON In vnClsPO.Value
                Dim vnHDocEntry As String = vnH.DocEntry
                Dim vnHDocNum As String = vnH.DocNum
                Dim vnHDocType As String = vnH.DocType
                Dim vnHDocDate As String = vnH.DocDate
                Dim vnHDocDueDate As String = vnH.DocDueDate
                Dim vnHCardCode As String = vnH.CardCode
                Dim vnHCardName As String = vnH.CardName
                Dim vnHAddress As String = vnH.Address
                Dim vnHDocumentStatus As String = vnH.DocumentStatus
                Dim vnHCancelled As String = vnH.Cancelled
                Dim vnHGDGCode As String = ""
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("=======")

                vriTextStream.WriteLine("vnHDocEntry=" & vnHDocEntry)
                vriTextStream.WriteLine("vnHDocNum=" & vnHDocNum)
                vriTextStream.WriteLine("vnHDocType=" & vnHDocType)
                vriTextStream.WriteLine("vnHDocDate=" & vnHDocDate)
                vriTextStream.WriteLine("vnHDocDueDate=" & vnHDocDueDate)
                vriTextStream.WriteLine("vnHCardCode=" & vnHCardCode)
                vriTextStream.WriteLine("vnHCardName=" & vnHCardName)
                vriTextStream.WriteLine("vnHAddress=" & vnHAddress)
                vriTextStream.WriteLine("vnHDocumentStatus=" & vnHDocumentStatus)
                vriTextStream.WriteLine("vnHCancelled=" & vnHCancelled)

                vriTextStream.WriteLine("-------")
                For Each vnD As ClsPurchaseOrder_DocumentLine_JSON In vnH.DocumentLines
                    Dim vnDLineNum As String = vnD.LineNum
                    Dim vnDItemCode As String = vnD.ItemCode
                    Dim vnDItemDescription As String = vnD.ItemDescription
                    Dim vnDQuantity As String = vnD.Quantity

                    vriTextStream.WriteLine("vnDLineNum=" & vnDLineNum)
                    vriTextStream.WriteLine("vnDItemCode=" & vnDItemCode)
                    vriTextStream.WriteLine("vnDItemDescription=" & vnDItemDescription)
                    vriTextStream.WriteLine("vnDQuantity=" & vnDQuantity)

                    '<---23 Oct 2023 Sampai sini
                    vnQuery = "Insert into #Sys_SsoPO_Temp"
                    vnQuery += vbCrLf & "(CompanyCode,PO_NO,JURNAL,PO_DATE,SUB,BRG,GDGCODE,"
                    vnQuery += vbCrLf & "QTY,NAMA_SUPPLIER,NAMA_BARANG,JOBNAME,"
                    vnQuery += vbCrLf & "DocEntry,DocType,DocumentStatus,Cancelled,"
                    vnQuery += vbCrLf & "POFileXlsOID,UploadSourceOID,UploadDatetime"
                    vnQuery += vbCrLf & ")"

                    vnQuery += vbCrLf & "Select '" & vriCompanyCode & "'CompanyCode,'" & vnHDocNum & "'PO_NO,''JURNAL,'" & vnHDocDate & "'PO_DATE,'" & vnHCardCode & "'SUB,'" & vnDItemCode & "'BRG,'" & vnHGDGCode & "'GDGCODE,"
                    vnQuery += vbCrLf & "'" & vnDQuantity & "'QTY,'" & vnHCardName & "'NAMA_SUPPLIER,''NAMA_BARANG,''JOBNAME,"
                    vnQuery += vbCrLf & vnHDocEntry & ",'" & vnHDocType & "','" & vnHDocumentStatus & "','" & vnHCancelled & "',"
                    vnQuery += vbCrLf & vriHOID & "," & enuUploadSource.SAP_Api & ",'" & vnGetDate & "'"
                    vriTextStream.WriteLine("vnQuery")
                    vriTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                    '<<==
                Next
            Next

            fbuSAP_Logout(vriTextStream)

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Return True")
            vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vriTextStream.WriteLine("<<===========================fbuSAPApi_GetPurchaseOrder")
            vriTextStream.WriteLine("")
            Return True
        Else
            vroMessage = vnMessage

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Return False")
            vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vriTextStream.WriteLine("<<===========================fbuSAPApi_GetPurchaseOrder")
            vriTextStream.WriteLine("")
            Return False
        End If
    End Function

    Public Function fbuSAP_Logout(vriTextStream As Scripting.TextStream) As Boolean
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<------------------------fbuSAP_Logout")

        Dim vnClient_Logout As RestClient, vnResponse_Logout As IRestResponse
        vnClient_Logout = New RestClient("https://sap.sumberberkat.com:51100/b1s/v1/Logout")
        vnClient_Logout.Timeout = -1

        Dim vnRequest_Logout = New RestRequest(Method.GET)
        vnRequest_Logout.AddHeader("Content-Type", "application/json")

        vnRequest_Logout.AddCookie("B1SESSION", vsB1SESSION.Value)
        vnRequest_Logout.AddCookie("ROUTEID", vsROUTEID.Value)

        vnResponse_Logout = vnClient_Logout.Execute(vnRequest_Logout)

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("vnResponse_Logout.Content")
        vriTextStream.WriteLine(vnResponse_Logout.Content)
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("")

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("Return True")
        vriTextStream.WriteLine("<<=========================fbuSAP_Logout")
        vriTextStream.WriteLine("")

        Return True
    End Function
End Module
