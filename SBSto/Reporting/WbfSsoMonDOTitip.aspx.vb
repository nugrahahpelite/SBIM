Imports System.Data.SqlClient
Public Class WbfSsoMonDOTitip
    Inherits System.Web.UI.Page
    Private Sub psDefaultDisplay()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Reporting"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoStorageStock, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        psClearMessage()

        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            Exit Sub
        End If

        psFillGrvList()
    End Sub

    Private Sub psFillGrvList()
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select sm.vCompanyCode,sm.vKodeBarang,mb.BRGNAME,sm.vTotal_QtySisaInvoice,sm.vTotal_QtySisaStock,sm.vTotal_QtySelisih"
        vnQuery += vbCrLf & "From ("
        vnQuery += vbCrLf & "Select case when isnull(tb1.CompanyCode,'')='' then tb2.CompanyCode else tb1.CompanyCode end vCompanyCode,"
        vnQuery += vbCrLf & "       case when isnull(tb1.KodeBarang,'')='' then tb2.BRGCODE else tb1.KodeBarang end vKodeBarang,"
        vnQuery += vbCrLf & "	    isnull(tb1.vTotal_QtySisaInvoice,0)vTotal_QtySisaInvoice,"
        vnQuery += vbCrLf & "	    isnull(tb2.vTotal_QtySisaStock,0) vTotal_QtySisaStock,"
        vnQuery += vbCrLf & "	    isnull(tb1.vTotal_QtySisaInvoice,0) - isnull(tb2.vTotal_QtySisaStock,0) vTotal_QtySelisih"
        vnQuery += vbCrLf & "  From " & vnDBDcm & "fnTbl_DcmDOTitip_SisaInvoice() tb1"
        vnQuery += vbCrLf & "       full join fnTbl_SsoDOTitip_SisaStorageStock()tb2 on tb2.CompanyCode=tb1.CompanyCode and tb2.BRGCODE=tb1.KodeBarang"
        vnQuery += vbCrLf & ")sm"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.vCompanyCode and mb.BRGCODE=sm.vKodeBarang"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=sm.vCompanyCode and uc.UserOID=" & vnUserOID
        End If

        vnQuery += vbCrLf & " Where 1=1"
        vnQuery += vbCrLf & "            and sm.vCompanyCode='" & DstListCompany.SelectedValue & "'"

        If ChkVarianOnly.Checked Then
            vnQuery += vbCrLf & "       and isnull(sm.vTotal_QtySisaInvoice,0) - isnull(sm.vTotal_QtySisaStock,0)<>0"
        End If

        vnQuery += vbCrLf & "Order by sm.vCompanyCode,sm.vKodeBarang"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Protected Sub BtnXLS_Click(sender As Object, e As EventArgs) Handles BtnXLS.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnFileName As String = ""
        pbuCreateXlsx_MonDoTitip(vnFileName, Session("UserOID"), DstListCompany, ChkVarianOnly, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class