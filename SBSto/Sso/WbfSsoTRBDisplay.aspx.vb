Imports System.Data.SqlClient
Public Class WbfSsoTRBDisplay
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoTRBDisplay"

    Enum ensColTRBH
        TRBHOID = 0
        CompanyCode = 1
        NoBukti = 2
        vTanggal = 3
        GudangAsal = 4
        GudangTujuan = 5
        GudangAsalOID = 6
        GudangTujuanOID = 7
    End Enum

    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Sso"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgTRBError.Text = pbMsgError
                LblMsgTRBError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPickList, vnSQLConn)
            'pbuFillDstDcmGudang(DstTRBWhAsal, True, vnSQLConn)
            'pbuFillDstDcmGudang(DstTRBWhTujuan, True, vnSQLConn)

            pbuFillDstCompanyByUser(Session("UserOID"), DstTRBCompany, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub
    Private Sub psClearMessage()
        LblMsgTRBError.Text = ""
        LblMsgTRBFindError.Text = ""
        LblMsgTRBCompany.Text = ""
    End Sub

    Protected Sub BtnTRBFind_Click(sender As Object, e As EventArgs) Handles BtnTRBFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgTRBFindError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgTRBFindError.Visible = True
            Exit Sub
        End If

        LblMsgTRBFindError.Text = ""
        LblMsgTRBFindError.Visible = False

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgTRBFindError.Text = pbMsgError
            LblMsgTRBFindError.Visible = True
            Exit Sub
        End If

        psFillGrvTRBH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvTRBH(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String

        vnCriteria = "      Where 1=1"

        If DstTRBCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and tbh.CompanyCode='" & DstTRBCompany.SelectedValue & "'"
        End If
        If Trim(TxtTRBNo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and tbh.NoBukti like '%" & fbuFormatString(Trim(TxtTRBNo.Text)) & "%'"
        End If
        If IsDate(TxtTRBStart.Text) Then
            vnCriteria += vbCrLf & "            and tbh.Tanggal >= '" & TxtTRBStart.Text & "'"
        End If
        If IsDate(TxtTRBEnd.Text) Then
            vnCriteria += vbCrLf & "            and tbh.Tanggal <= '" & TxtTRBEnd.Text & "'"
        End If

        vnQuery = "Select tbh.OID TRBHOID,tbh.CompanyCode,tbh.NoBukti,convert(varchar(11),tbh.Tanggal,106)vTanggal,tbh.GudangAsal,tbh.GudangTujuan,tbh.GudangAsalOID,tbh.GudangTujuanOID,"
        vnQuery += vbCrLf & "       case when abs(tbh.IsPickListClosed)=1 then 'Y' else 'N' end vIsPickListClosed,"
        vnQuery += vbCrLf & "       case when abs(tbh.TRBCancel)=1 then 'Y' else 'N' end vTRBCancel"
        vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmTRBHeader_TR tbh"

        If vnUserCompanyCode <> "" And DstTRBCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=poh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria

        If DstTRBCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and tbh.CompanyCode='" & DstTRBCompany.SelectedValue & "'"
        End If
        'If DstTRBWhAsal.SelectedIndex > 0 Then
        '    vnQuery += vbCrLf & "            and tbh.GudangAsalOID=" & DstTRBWhAsal.SelectedValue
        'End If
        'If DstTRBWhTujuan.SelectedIndex > 0 Then
        '    vnQuery += vbCrLf & "            and tbh.GudangTujuanOID=" & DstTRBWhTujuan.SelectedValue
        'End If

        vnQuery += vbCrLf & "Order by tbh.CompanyCode,tbh.NoBukti"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvTRBH.DataSource = vnDtb
        GrvTRBH.DataBind()

        PanTRBD.Visible = False

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub
    Private Sub psFillGrvTRBD(vriTRBHOID As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select row_number()over(order by KodeBrg)vSeqNo,KodeBrg,NamaBrg,Qty,QtyOnPickList,Satuan,Keterangan"
        vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmTRBDetail_TR tbrb"

        vnQuery += vbCrLf & "      Where TRBHOID=" & vriTRBHOID
        vnQuery += vbCrLf & "Order by KodeBrg"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvTRBD.DataSource = vnDtb
        GrvTRBD.DataBind()

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub GrvTRBD_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvTRBD.PageIndexChanging
        GrvTRBD.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgTRBError.Text = pbMsgError
            LblMsgTRBError.Visible = True
            Exit Sub
        End If

        psFillGrvTRBD(LblMsgTRBHOID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvTRBD_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvTRBD.SelectedIndexChanged

    End Sub

    Protected Sub GrvTRBH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvTRBH.SelectedIndexChanged

    End Sub

    Private Sub GrvTRBH_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvTRBH.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvTRBH.Rows(vnIdx)
        If vnIdx >= GrvTRBH.Rows.Count Then Exit Sub

        If e.CommandName = "NoBukti" Then
            Dim vnTRBHOID As String = vnGRow.Cells(ensColTRBH.TRBHOID).Text
            Dim vnTRBNo As String = DirectCast(vnGRow.Cells(ensColTRBH.NoBukti).Controls(0), LinkButton).Text
            LblMsgTRBHOID.Text = vnTRBHOID
            LblMsgTRBDNo.Text = vnTRBNo

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgTRBError.Text = pbMsgError
                LblMsgTRBError.Visible = True
                Exit Sub
            End If

            psFillGrvTRBD(vnTRBHOID, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanTRBD.Visible = True
        End If
    End Sub

    Private Sub GrvTRBH_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvTRBH.PageIndexChanging
        GrvTRBH.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgTRBError.Text = pbMsgError
            LblMsgTRBError.Visible = True
            Exit Sub
        End If

        psFillGrvTRBH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        PanTRBD.Visible = False
    End Sub

End Class