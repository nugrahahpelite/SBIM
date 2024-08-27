Imports System.Data.SqlClient
Public Class WbfDmDOTSummary
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfDOTSummary"

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

            psFillGrvCust(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If

    End Sub
    Private Sub psDefaultDisplay()

    End Sub

    Private Sub psFillGrvCust(vriSQLConn As SqlConnection)
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select distinct nh.CompanyCode,nh.CustCode,nh.CustName,nh.CustAddress"
        vnQuery += vbCrLf & "      From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"

        If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=mj.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & " Where nh.NotaDOT=1"
        If DstCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "            and nh.CompanyCode='" & DstCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & "Order by nh.CustCode"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvCust.DataSource = vnDtb
        GrvCust.DataBind()
    End Sub

    Private Sub GrvCust_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvCust.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvCust.Rows(vnIdx)
        If e.CommandName = "CustCode" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnCompCode As String = vnGRow.Cells(ensColCust.CompanyCode).Text
            Dim vnCustCode As String = DirectCast(vnGRow.Cells(ensColCust.CustCode).Controls(0), LinkButton).Text

            HdfCompCode.Value = vnCompCode
            HdfCustCode.Value = vnCustCode
            LblSM.Text = vnCompCode & " " & vnCustCode & " " & vnGRow.Cells(ensColCust.CustName).Text

            psFillGrvSM1(vnCompCode, vnCustCode, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvSM1(vriCompCode As String, vriCustCode As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String = ""

        If RdlSM.SelectedValue = "BRG" Then
            vnQuery = "Select 0 NotaHOID,''NotaNo,Null vNotaDate,nd.KodeBarang,nd.NamaBarang,"
            vnQuery += vbCrLf & "	    sum(nd.Qty + nd.QtyBonus)vTotalQtyInvoice,sum(nd.QtyOnPKDOT)vTotalQtyPKDOT,sum(nd.Qty + nd.QtyBonus - nd.QtyOnPKDOT)vTotalQtySisa" 'sum(nd.QtyOnPickList)vTotalQtyOnPickList"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"
            vnQuery += vbCrLf & "	    inner join " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) on nd.NotaHOID=nh.OID"
            vnQuery += vbCrLf & " Where nh.NotaDOT=1 and nh.CompanyCode='" & vriCompCode & "' and nh.CustCode='" & vriCustCode & "'"
            vnQuery += vbCrLf & " Group by nd.KodeBarang,nd.NamaBarang"

            vnQuery += vbCrLf & "Order by nd.KodeBarang"

            GrvSM1.Columns(ensColSM1.NotaNo).HeaderStyle.CssClass = ""
            GrvSM1.Columns(ensColSM1.NotaNo).ItemStyle.CssClass = ""

            GrvSM1.Columns(ensColSM1.vNotaDate).HeaderStyle.CssClass = ""
            GrvSM1.Columns(ensColSM1.vNotaDate).ItemStyle.CssClass = ""

        ElseIf RdlSM.SelectedValue = "INV_BRG" Then
            vnQuery = "Select nd.NotaHOID,nh.NotaNo,convert(varchar(11),nh.NotaDate,106)vNotaDate,nd.KodeBarang,nd.NamaBarang,"
            vnQuery += vbCrLf & "	    sum(nd.Qty + nd.QtyBonus)vTotalQtyInvoice,sum(nd.QtyOnPKDOT)vTotalQtyPKDOT,sum(nd.Qty + nd.QtyBonus - nd.QtyOnPKDOT)vTotalQtySisa" 'sum(nd.QtyOnPickList)vTotalQtyOnPickList"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"
            vnQuery += vbCrLf & "	    inner join " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) on nd.NotaHOID=nh.OID"
            vnQuery += vbCrLf & " Where nh.NotaDOT=1 and nh.CompanyCode='" & vriCompCode & "' and nh.CustCode='" & vriCustCode & "'"
            vnQuery += vbCrLf & " Group by nd.NotaHOID,nh.NotaNo,convert(varchar(11),nh.NotaDate,106),nd.KodeBarang,nd.NamaBarang"

            vnQuery += vbCrLf & "Order by nh.NotaNo,nd.KodeBarang"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSM1.DataSource = vnDtb
        GrvSM1.DataBind()

        psSetColSM1()
    End Sub

    Private Sub psSetColSM1()
        If RdlSM.SelectedValue = "BRG" Then
            GrvSM1.Columns(ensColSM1.NotaNo).HeaderStyle.CssClass = "myDisplayNone"
            GrvSM1.Columns(ensColSM1.NotaNo).ItemStyle.CssClass = "myDisplayNone"

            GrvSM1.Columns(ensColSM1.vNotaDate).HeaderStyle.CssClass = "myDisplayNone"
            GrvSM1.Columns(ensColSM1.vNotaDate).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvSM1.Columns(ensColSM1.NotaNo).HeaderStyle.CssClass = ""
            GrvSM1.Columns(ensColSM1.NotaNo).ItemStyle.CssClass = ""

            GrvSM1.Columns(ensColSM1.vNotaDate).HeaderStyle.CssClass = ""
            GrvSM1.Columns(ensColSM1.vNotaDate).ItemStyle.CssClass = ""
        End If
    End Sub
    Protected Sub GrvSM1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSM1.SelectedIndexChanged

    End Sub

    Private Sub RdlSM_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdlSM.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvSM1(HdfCompCode.Value, HdfCustCode.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvSM1_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSM1.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvSM1.Rows(vnIdx)
        If e.CommandName = "KodeBarang" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnBrgCode As String = DirectCast(vnGRow.Cells(ensColSM1.KodeBarang).Controls(0), LinkButton).Text
            HdfBrgCode.Value = vnBrgCode

            LblSM2.Text = "HISTORY PERINTAH KIRIM DO TITIP " & vnBrgCode & " " & vnGRow.Cells(ensColSM1.NamaBarang).Text

            psFillGrvSM2(HdfCompCode.Value, HdfCustCode.Value, vnBrgCode, Replace(vnGRow.Cells(ensColSM1.vTotalQtyPKDOT).Text, ",", ""), vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvSM2(vriCompCode As String, vriCustCode As String, vriBrgCode As String, vriTotalQtyPKDOT As Integer, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String = ""

        vnQuery = "Select * From("
        vnQuery += vbCrLf & "Select TransCode,PKDOTHOID,vTransNo,vTransDate,PKDOTNote,"
        vnQuery += vbCrLf & "      vShipToName,vTransQty,Null vTotalQtyPKDOT"
        vnQuery += vbCrLf & "  From " & vnDBDcm & "fnTbl_DcmDOTitip_History_ByComp_Cust_Brg('" & vriCompCode & "','" & vriCustCode & "','" & vriBrgCode & "')"

        vnQuery += vbCrLf & "UNION"

        vnQuery += vbCrLf & "Select 'TOTAL' TransCode,Null PKDOTHOID,''vTransNo,''vTransDate,''PKDOTNote,"
        vnQuery += vbCrLf & "      ''vShipToName,sum(vTransQty)vTransQty," & vriTotalQtyPKDOT & " vTotalQtyPKDOT"
        vnQuery += vbCrLf & "  From " & vnDBDcm & "fnTbl_DcmDOTitip_History_ByComp_Cust_Brg('" & vriCompCode & "','" & vriCustCode & "','" & vriBrgCode & "')"

        vnQuery += vbCrLf & ")tb Order by TransCode,vTransNo"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSM2.DataSource = vnDtb
        GrvSM2.DataBind()

        psSetColSM1()
    End Sub

    Private Sub BtnNotaFind_Click(sender As Object, e As EventArgs) Handles BtnNotaFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvCust(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvCust_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvCust.SelectedIndexChanged

    End Sub
End Class