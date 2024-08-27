Imports System.Data.SqlClient
Public Class WbfSsoTransStatus
    Inherits System.Web.UI.Page
    Const csTNoPrefix = "TS"

    Enum ensColList
        TransStatusDescr = 0
        vTransStatusBy = 1
        vTransStatusInfo = 2
        vTransStatusDatetime = 3
    End Enum
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If vbuPreviewOnClose <> "1" Then
            If Not IsPostBack Then
                TxtTransNo.Text = Request.QueryString("vqTrNo")
                psFillGrvList()
            End If
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnTbName As String = ""
        Dim vnWhere As String = ""

        Select Case Request.QueryString("vqTrCode")
            Case stuTransCode.SsoSSOH
                vnTbName = "Sys_SsoSOStatus_TR tr"
                vnWhere = "tr.SOHOID="
            Case stuTransCode.SsoReceiving
                vnTbName = "Sys_SsoRcvStatus_TR tr"
                vnWhere = "tr.RcvHOID="
            Case stuTransCode.SsoSSOC
                vnTbName = "Sys_SsoSOCompareStatus_TR tr"
                vnWhere = "tr.SOCHOID="
            Case stuTransCode.SsoPOPackingList
                vnTbName = "Sys_SsoPLStatus_TR tr"
                vnWhere = "tr.PLHOID="
            Case stuTransCode.SsoSummaryTRB
                vnTbName = "Sys_SsoSmTRBStatus_TR tr"
                vnWhere = "tr.SmTRBHOID="
            Case stuTransCode.SsoPenerimaanPembelian
                vnTbName = "Sys_SsoRcvPOStatus_TR tr"
                vnWhere = "tr.RcvPOHOID="
            Case stuTransCode.SsoPutaway
                vnTbName = "Sys_SsoPWStatus_TR tr"
                vnWhere = "tr.PWHOID="
            Case stuTransCode.SsoPutaway_Antar_Wh
                vnTbName = "Sys_SsoPYStatus_TR tr"
                vnWhere = "tr.PYHOID="
            Case stuTransCode.SsoPickList
                vnTbName = "Sys_SsoPCLStatus_TR tr"
                vnWhere = "tr.PCLHOID="
            Case stuTransCode.SsoPicking
                vnTbName = "Sys_SsoPCKStatus_TR tr"
                vnWhere = "tr.PCKHOID="
            Case stuTransCode.SsoDispatch
                vnTbName = "Sys_SsoDSPStatus_TR tr"
                vnWhere = "tr.DSPHOID="
            Case stuTransCode.SsoDispatchReceive
                vnTbName = "Sys_SsoDSRStatus_TR tr"
                vnWhere = "tr.DSRHOID="
            Case stuTransCode.SsoPenerimaanLain2
                vnTbName = "Sys_SsoRcvMscStatus_TR tr"
                vnWhere = "tr.RcvMscHOID="
            Case stuTransCode.SsoPenerimaanKarantina
                vnTbName = "Sys_SsoRcvKRStatus_TR tr"
                vnWhere = "tr.RcvKRHOID="
            Case stuTransCode.SsoVoidSO
                vnTbName = "Sys_SsoSOrderVoidStatus_TR tr"
                vnWhere = "tr.SOrderVoidHOID="
            Case stuTransCode.SsoChangeSKU
                vnTbName = "Sys_SsoCSKUStatus_TR tr"
                vnWhere = "tr.CSKUHOID="
        End Select

        vnWhere = vnWhere & Request.QueryString("vqTrOID")

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        If Request.QueryString("vqTrCode") = stuTransCode.SsoDispatch Then
            GrvList.Columns(ensColList.vTransStatusInfo).Visible = True

            Dim vnDBDcm As String = fbuGetDBDcm()
            vnQuery = "Select ma.TransStatusDescr,um.UserName vTransStatusBy,"
            vnQuery += vbCrLf & "       dm.DcmDriverName+' '+vm.VehicleNo+'<br />'+StatusNote vTransStatusInfo,"
            vnQuery += vbCrLf & "       Convert(varchar(11),tr.TransStatusDatetime,106) + ' '+ Convert(varchar(11),tr.TransStatusDatetime,108)vTransStatusDatetime"
            vnQuery += vbCrLf & "       From " & vnTbName
            vnQuery += vbCrLf & "	         inner join Sys_SsoTransStatus_MA ma on ma.TransCode=tr.TransCode and ma.TransStatus=tr.TransStatus"
            vnQuery += vbCrLf & "	         inner join " & vnDBDcm & "Sys_DcmDriver_MA dm with(nolock) on dm.OID=tr.DcmSchDriverOID"
            vnQuery += vbCrLf & "	         inner join " & vnDBDcm & "Sys_DcmVehicle_MA vm with(nolock) on vm.OID=tr.DcmVehicleOID"
            vnQuery += vbCrLf & "			 inner join Sys_SsoUser_MA um on um.OID=tr.TransStatusUserOID"
            vnQuery += vbCrLf & "	   Where " & vnWhere
            vnQuery += vbCrLf & "Order by tr.OID"

        Else
            GrvList.Columns(ensColList.vTransStatusInfo).Visible = False

            vnQuery = "Select ma.TransStatusDescr,um.UserName vTransStatusBy,''vTransStatusInfo,"
            vnQuery += vbCrLf & "       Convert(varchar(11),tr.TransStatusDatetime,106) + ' '+ Convert(varchar(11),tr.TransStatusDatetime,108)vTransStatusDatetime"
            vnQuery += vbCrLf & "       From " & vnTbName
            vnQuery += vbCrLf & "	         inner join Sys_SsoTransStatus_MA ma on ma.TransCode=tr.TransCode and ma.TransStatus=tr.TransStatus"
            vnQuery += vbCrLf & "			 inner join Sys_SsoUser_MA um on um.OID=tr.TransStatusUserOID"
            vnQuery += vbCrLf & "	   Where " & vnWhere
            vnQuery += vbCrLf & "Order by tr.OID"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

End Class