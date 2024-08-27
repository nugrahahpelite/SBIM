Imports System.Data.SqlClient
Public Class WbfDmInvSummary
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfDmInvSummary"

    Enum ensColCust
        CompanyCode = 0
        CustCode = 1
        CustName = 2
        CustAddress = 3
    End Enum
    Enum ensColSM1
        NotaHOID = 0
        NotaNo = 1
        vNotaDate = 2
        KodeBarang = 3
        NamaBarang = 4
        vTotalQtyInvoice = 5
        vTotalQtyPKDOT = 6
    End Enum

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If

        Session("CurrentFolder") = "Reporting"
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPickList, vnSQLConn)
            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, True, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, True, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If

    End Sub
    Private Sub psDefaultDisplay()

    End Sub

    Private Sub psFillGrvInv(vriSQLConn As SqlConnection)
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select CompanyCode,NotaNo,KodeBarang,"
        vnQuery += vbCrLf & "	vTotalQty_Invoice,vTotalQty_Invoice_OnPickList,vTotalQty_Picklist,"
        vnQuery += vbCrLf & "	(isnull(vTotalQty_Invoice,0) - isnull(vTotalQty_Invoice_OnPickList,0))vSisa,"
        vnQuery += vbCrLf & "	(isnull(vTotalQty_Invoice,0) - isnull(vTotalQty_Picklist,0))vSelisihAC,"
        vnQuery += vbCrLf & "	(isnull(vTotalQty_Invoice_OnPickList,0) - isnull(vTotalQty_Picklist,0))vSelisihBC"
        vnQuery += vbCrLf & "From"
        vnQuery += vbCrLf & "	("
        vnQuery += vbCrLf & "    Select rtrim(nh.CompanyCode)CompanyCode,nh.NotaNo,nd.KodeBarang,"
        vnQuery += vbCrLf & "	        cast(sum(nd.TotalQty + nd.TotalQtyBonus) as int)vTotalQty_Invoice,"
        vnQuery += vbCrLf & "	 	    cast(sum(nd.TotalQtyOnPickList) as int)vTotalQty_Invoice_OnPickList"
        vnQuery += vbCrLf & "		   From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"
        vnQuery += vbCrLf & "		        inner join " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR nd with(nolock) on nd.NotaHOID=nh.OID"
        vnQuery += vbCrLf & "		        inner join Sys_SsoCutOfDate_CNF cd with(nolock) on cd.CompanyCode=rtrim(nh.CompanyCode)"

        If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "               inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=nh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "		  Where nh.UploadDatetime>=cd.CutOfDate"
        vnQuery += vbCrLf & "		  Group by rtrim(nh.CompanyCode),nh.NotaNo,nd.KodeBarang"
        vnQuery += vbCrLf & "   )mj_sum"
        vnQuery += vbCrLf & "   inner join"
        vnQuery += vbCrLf & "   (Select pch.PCLCompanyCode,pch.PCLRefHNo,pcr.BRGCODE,sum(pcr.ReservedQty)vTotalQty_Picklist"
        vnQuery += vbCrLf & "		   From Sys_SsoPCLReserve_TR pcr with(nolock)"
        vnQuery += vbCrLf & "				inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pcr.PCLHOID"
        vnQuery += vbCrLf & "		  Where pch.SchDTypeOID=" & enuSchDType.Invoice & " and pch.TransStatus > " & enuTCPICK.Baru
        vnQuery += vbCrLf & "		  Group by pch.PCLCompanyCode,pch.PCLRefHNo,pcr.BRGCODE"
        vnQuery += vbCrLf & "	)pc_sum on pc_sum.PCLCompanyCode=mj_sum.CompanyCode and pc_sum.BRGCODE=mj_sum.KodeBarang and pc_sum.PCLRefHNo=mj_sum.NotaNo"

        vnQuery += vbCrLf & " Where 1=1"
        If DstCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "            and CompanyCode='" & DstCompany.SelectedValue & "'"
        End If
        If ChkVarianOnly.Checked Then
            vnQuery += vbCrLf & "            and ("
            vnQuery += vbCrLf & "                  (isnull(vTotalQty_Invoice,0) - isnull(vTotalQty_Invoice_OnPickList,0)) < 0"
            vnQuery += vbCrLf & "                  or"
            vnQuery += vbCrLf & "                  (isnull(vTotalQty_Invoice_OnPickList,0) - isnull(vTotalQty_Picklist,0))<>0"
            vnQuery += vbCrLf & "                )"
        End If
        vnQuery += vbCrLf & "Order by CompanyCode,NotaNo,KodeBarang"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvInv.DataSource = vnDtb
        GrvInv.DataBind()

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub Inv_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvInv.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvInv.Rows(vnIdx)
    End Sub

    Private Sub BtnNotaFind_Click(sender As Object, e As EventArgs) Handles BtnNotaFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvInv(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvCust_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvInv.SelectedIndexChanged

    End Sub
End Class