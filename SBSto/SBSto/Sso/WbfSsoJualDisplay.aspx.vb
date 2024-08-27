Imports System.Data.SqlClient
Public Class WbfSsoJualDisplay
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoJualDisplay"

    Enum ensColList
        CompanyCode = 0
        no_nota = 1
        vtanggal = 2
        kode_cust = 3
        CUSTOMER = 4
    End Enum

    Enum ensColSumNota
        vKodeBarang = 0
        vNamaBarang = 1
    End Enum
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If

        Session("CurrentFolder") = "DMgm"
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
            pbuFillDstWarehouse(DstWarehouse, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnNotaFind_Click(sender As Object, e As EventArgs) Handles BtnNotaFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvList(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub psFillGrvList(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String
        Dim vnCust As String = fbuFormatString(Trim(TxtListCustomer.Text))
        vnCriteria = "      Where 1=1"
        If DstCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and mj.CompanyCode='" & DstCompany.SelectedValue & "'"
        End If
        If DstWarehouse.SelectedIndex > 0 Then
            vnCriteria += vbCrLf & "            and mj.WarehouseOID=" & DstWarehouse.SelectedValue
        End If
        If Trim(TxtListCustomer.Text) <> "" Then
            vnCriteria += vbCrLf & "            and (mj.kode_cust like '%" & vnCust & "%' or mj.CUSTOMER like '%" & vnCust & "%')"
        End If
        If Trim(TxtNotaNo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and mj.no_nota like '%" & fbuFormatString(Trim(TxtNotaNo.Text)) & "%'"
        End If
        If IsDate(TxtListStart.Text) Then
            vnCriteria += vbCrLf & "            and mj.tanggal >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnCriteria += vbCrLf & "            and mj.tanggal <= '" & TxtListEnd.Text & "'"
        End If
        If IsDate(TxtListUploadStart.Text) Then
            vnCriteria += vbCrLf & "            and cast(mj.UploadDatetime as date) >= '" & TxtListUploadStart.Text & "'"
        End If
        If IsDate(TxtListUploadEnd.Text) Then
            vnCriteria += vbCrLf & "            and cast(mj.UploadDatetime as date) <= '" & TxtListUploadEnd.Text & "'"
        End If
        If ChkListIsPicklist_Yes.Checked = False And ChkListIsPicklist_No.Checked = False Then
            ChkListIsPicklist_Yes.Checked = True
            ChkListIsPicklist_No.Checked = True
        End If
        If Not (ChkListIsPicklist_Yes.Checked = True And ChkListIsPicklist_No.Checked = True) Then
            If ChkListIsPicklist_Yes.Checked = True Then
                vnCriteria += vbCrLf & "            and abs(mj.IsPickListClosed)=1"
            Else
                vnCriteria += vbCrLf & "            and abs(mj.IsPickListClosed)=0"
            End If
        End If

        vnQuery = "Select Distinct mj.CompanyCode,mj.no_nota,convert(varchar(11),mj.tanggal,106)vtanggal,mj.kode_cust,mj.CUSTOMER,wm.WarehouseName,mj.ALAMAT,mj.kota,"
        vnQuery += vbCrLf & "            case when abs(mj.NotaDOT)=1 then 'Y' else 'N' end vDOTitip,"
        vnQuery += vbCrLf & "            case when abs(mj.IsPickListClosed)=1 then 'Y' else 'N' end vIsPickListClosed,"
        vnQuery += vbCrLf & "            convert(varchar(11),mj.UploadDatetime,106)+' '+convert(varchar(5),mj.UploadDatetime,108)vUploadDatetime,"
        vnQuery += vbCrLf & "            pm.PriorityName,"
        vnQuery += vbCrLf & "            convert(varchar(11),mj.NotaPRIODatetime,106)+' '+convert(varchar(5),mj.NotaPRIODatetime,108)vNotaPRIODatetime,"
        vnQuery += vbCrLf & "            case when abs(mj.NotaCancel)=1 then 'Y' else 'N' end vNotaCancel,"
        vnQuery += vbCrLf & "            mj.NotaCancelNote,mj.NotaCancelReturNo,"
        vnQuery += vbCrLf & "            convert(varchar(11),mj.NotaCancelDatetime,106)+' '+convert(varchar(5),mj.NotaCancelDatetime,108)vNotaCancelDatetime,"
        vnQuery += vbCrLf & "            case when mj.NotaTKF=1 then 'Faktur Ditukar' when mj.NotaTKF=2 then 'Faktur Pengganti' Else '' End vNotaTKF,"
        vnQuery += vbCrLf & "            nh.NotaNo vNotaNo_Baru,mj.NotaRJSNo"
        vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmJUAL mj with(nolock)"
        vnQuery += vbCrLf & "            inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA wm with(nolock) on wm.OID=mj.WarehouseOID"
        vnQuery += vbCrLf & "            left outer join " & vnDBDcm & "Sys_DcmInvPriority_MA pm with(nolock) on pm.OID=mj.PrioTypeOID"
        vnQuery += vbCrLf & "            left outer join " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock) on nh.OID=mj.NotaHOID_TKF"

        If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=mj.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by mj.CompanyCode,mj.no_nota"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvList(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvList.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)
        If e.CommandName = "no_nota" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnNotaNo As String = DirectCast(vnGRow.Cells(ensColList.no_nota).Controls(0), LinkButton).Text

            TxtNotaNo1.Text = vnNotaNo
            TxtNotaCompany.Text = vnGRow.Cells(ensColList.CompanyCode).Text
            TxtNotaDate.Text = vnGRow.Cells(ensColList.vtanggal).Text
            TxtNotaCustomer.Text = vnGRow.Cells(ensColList.kode_cust).Text & " " & vnGRow.Cells(ensColList.CUSTOMER).Text

            If ChkSummary.Checked Then
                psFillGrvSumNota(vnGRow.Cells(ensColList.CompanyCode).Text, vnNotaNo, vnSQLConn)
            Else
                psFillGrvNota(vnGRow.Cells(ensColList.CompanyCode).Text, vnNotaNo, vnSQLConn)
            End If

            psShowNota(True)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub psFillGrvNota(vriCompCode As String, vriNotaNo As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select nd.NotaHOID,nd.OID vNotaDOID,nd.KodeBarang KODE_BARANG,nd.NamaBarang NAMA_BARANG,nd.QTY,nd.QTYBONUS,nd.QtyOnPKDOT,nd.QtyOnPickList,nd.SATUAN,nd.NoRef NO_REF,nd.SALESMAN"
        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) on nd.NotaHOID=nh.OID"
        vnQuery += vbCrLf & "     Where nh.CompanyCode='" & vriCompCode & "' and nh.NotaNo='" & fbuFormatString(vriNotaNo) & "'"
        vnQuery += vbCrLf & "Order by nd.KodeBarang"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            vnDtb = New DataTable
            vnQuery = "Select 0 NotaHOID,0 vNotaDOID,KODE_BARANG,NAMA_BARANG,QTY,QTYBONUS,0 QtyOnPKDOT,0 QtyOnPickList,SATUAN,NO_REF,SALESMAN"
            vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmJual with(nolock)"
            vnQuery += vbCrLf & "     Where CompanyCode='" & vriCompCode & "' and NO_NOTA='" & fbuFormatString(vriNotaNo) & "'"
            vnQuery += vbCrLf & "Order by KODE_BARANG"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvNota.DataSource = vnDtb
            GrvNota.DataBind()
        Else
            GrvNota.DataSource = vnDtb
            GrvNota.DataBind()

            psFillGrvNota_ByBarang(vriCompCode, vnDtb.Rows(0).Item("NotaHOID"), vriSQLConn)
        End If

        vnDtb.Dispose()
    End Sub

    Private Sub psFillGrvNota_ByBarang(vriCompanyCode As String, vriNotaHOID As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select " & vriNotaHOID & " NotaHOID,nd.KodeBarang,mb.BRGNAME,nd.TotalQty,nd.TotalQtyBonus,nd.TotalQtyOnPKDOT,nd.TotalQtyOnPickList"
        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR nd with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=nd.KodeBarang and mb.CompanyCode='" & vriCompanyCode & "'"
        vnQuery += vbCrLf & "     Where nd.NotaHOID=" & vriNotaHOID
        vnQuery += vbCrLf & "Order by KodeBarang"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvNota_ByBarang.DataSource = vnDtb
        GrvNota_ByBarang.DataBind()

        vnDtb.Dispose()
    End Sub

    Private Sub psFillGrvSumNota(vriCompCode As String, vriNotaNo As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select vKodeBarang,vNamaBarang,vSatuan,vQtyNota,vQtyNotaBonus,"
        vnQuery += vbCrLf & "     vQtySKK_Closed,vQtySKK_NotClosed,"
        vnQuery += vbCrLf & "     vQtySJ_Closed,vQtySJ_NotClosed,"
        vnQuery += vbCrLf & "     vQtyPL_Closed,vQtyPL_NotClosed,vQtySisa,'Detail'vDetail"
        vnQuery += vbCrLf & "From " & vnDBDcm & "fnTbl_NotaQtySisa('" & vriCompCode & "','" & fbuFormatString(vriNotaNo) & "')"
        vnQuery += vbCrLf & "order by vNamaBarang"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSumNota.DataSource = vnDtb
        GrvSumNota.DataBind()

        vnDtb.Dispose()
    End Sub

    Private Sub psFillGrvDetailNota(vriCompCode As String, vriNotaNo As String, vriKodeBarang As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select vTransOID,vTransCode,vTransName,vTransNo,vTransDate,vTransStatus,vTransStatusName,vTransQty"
        vnQuery += vbCrLf & "From " & vnDBDcm & "fnTbl_NotaQtySisa_Detail('" & vriCompCode & "','" & fbuFormatString(vriNotaNo) & "','" & vriKodeBarang & "')"
        vnQuery += vbCrLf & "order by vTransDate"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvDetailNota.DataSource = vnDtb
        GrvDetailNota.DataBind()

        vnDtb.Dispose()
    End Sub
    Protected Sub BtnRefreshCustomer_Click(sender As Object, e As EventArgs) Handles BtnRefreshCustomer.Click
        Dim vnSQLConnDWH As New SqlConnection
        If Not fbuConnectSQLDWH(vnSQLConnDWH) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTransDWH As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnUserCompanyCode As String = Session("UserCompanyCode")
            Dim vnDBDcm As String = fbuGetDBDcm()

            Dim vnQuery As String
            Dim vnCriteria As String = ""
            Dim vnCount1 As Integer
            Dim vnCount2 As Integer
            vnCriteria = "      Where 1=1"
            If DstCompany.SelectedValue <> "" Then
                vnCriteria += vbCrLf & "            and mj.CompanyCode='" & DstCompany.SelectedValue & "'"
            End If
            If Trim(TxtListCustomer.Text) <> "" Then
                vnCriteria += vbCrLf & "            and mj.CUSTOMER like '%" & fbuFormatString(Trim(TxtListCustomer.Text)) & "%'"
            End If
            If Trim(TxtNotaNo.Text) <> "" Then
                vnCriteria += vbCrLf & "            and mj.no_nota like '%" & fbuFormatString(Trim(TxtNotaNo.Text)) & "%'"
            End If
            If IsDate(TxtListStart.Text) Then
                vnCriteria += vbCrLf & "            and mj.tanggal >= '" & TxtListStart.Text & "'"
            End If
            If IsDate(TxtListEnd.Text) Then
                vnCriteria += vbCrLf & "            and mj.tanggal <= '" & TxtListEnd.Text & "'"
            End If
            vnCriteria += vbCrLf & "            and not mj.CompanyCode+'x'+mj.KODE_CUST in (Select b.CompanyCode+'x'+b.CUSTSUB From " & fbuGetDBMaster() & "Sys_MstCustomer_MA b)"

            vnSQLTransDWH = vnSQLConnDWH.BeginTransaction()
            vnBeginTrans = True
            vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomer_MA("
            vnQuery += vbCrLf & "CompanyCode,CUSTSUB,CUSTNAME,CUSTPERSON,"
            vnQuery += vbCrLf & "CUSTPHONE,CUSTFAX,CUSTADDRESS,CUSTPOST,"
            vnQuery += vbCrLf & "CUSTKOTA,CUSTKOTANAMA,CustXlsFileOID)"

            vnQuery += vbCrLf & "Select distinct CompanyCode,KODE_CUST,CUSTOMER,''CUSTPERSON,"
            vnQuery += vbCrLf & "       ''CUSTPHONE,''CUSTFAX,ALAMAT CUSTADDRESS,''CUSTPOST,"
            vnQuery += vbCrLf & "       kota CUSTKOTA,kota CUSTKOTANAMA,0 CustXlsFileOID"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmJUAL mj"

            If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
                vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=mj.CompanyCode and mu.UserOID=" & Session("UserOID")
            End If

            vnQuery += vbCrLf & vnCriteria
            vnCount1 = fbuExecuteSQLTransScalar(vnQuery, vnSQLConnDWH, vnSQLTransDWH)

            vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomer_MA("
            vnQuery += vbCrLf & "CompanyCode,CUSTSUB,CUSTNAME,CUSTPERSON,"
            vnQuery += vbCrLf & "CUSTPHONE,CUSTFAX,CUSTADDRESS,CUSTPOST,"
            vnQuery += vbCrLf & "CUSTKOTA,CUSTKOTANAMA,CustXlsFileOID)"

            vnQuery += vbCrLf & "Select distinct CompanyCode,KODE_CUST,CUSTOMER,''CUSTPERSON,"
            vnQuery += vbCrLf & "       ''CUSTPHONE,''CUSTFAX,ALAMAT CUSTADDRESS,''CUSTPOST,"
            vnQuery += vbCrLf & "       kota CUSTKOTA,kota CUSTKOTANAMA,0 CustXlsFileOID"
            vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "View_Winacc_Invoice mj"

            If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
                vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=mj.CompanyCode and mu.UserOID=" & Session("UserOID")
            End If

            vnQuery += vbCrLf & vnCriteria
            '01 Feb 2023 matiin dulu
            'The server principal ""stockopname"" is not able to access the database ""WINACC_SRB"" under the current security context.
            'vnCount2 = fbuExecuteSQLTransScalar(vnQuery, vnSQLConnDWH, vnSQLTransDWH)

            vnSQLTransDWH.Commit()
            vnSQLTransDWH.Dispose()
            vnSQLTransDWH = Nothing
            vnBeginTrans = False

            LblMsgRefreshCustomer.Text = "Refresh Customer Selesai " & CStr(vnCount1 + vnCount2) & " Data"

            vnSQLConnDWH.Close()
            vnSQLConnDWH.Dispose()
            vnSQLConnDWH = Nothing

        Catch ex As Exception
            LblMsgError.Text = ex.Message

            If vnBeginTrans Then
                vnSQLTransDWH.Rollback()
                vnSQLTransDWH.Dispose()
                vnSQLTransDWH = Nothing
            End If

            vnSQLConnDWH.Close()
            vnSQLConnDWH.Dispose()
            vnSQLConnDWH = Nothing
        End Try
    End Sub

    Protected Sub BtnNotaClose_Click(sender As Object, e As EventArgs) Handles BtnNotaClose.Click
        psShowNota(False)
    End Sub

    Private Sub psShowNota(vriBo As Boolean)
        If vriBo Then
            DivNota.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivNota.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psDefaultDisplay()
        DivNota.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        'PanNota.Style(HtmlTextWriterStyle.Position) = "absolute"
        'PanNota.Style(HtmlTextWriterStyle.Top) = "200px"
    End Sub

    Protected Sub ChkSummary_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSummary.CheckedChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If ChkSummary.Checked Then
            psFillGrvSumNota(TxtNotaCompany.Text, TxtNotaNo1.Text, vnSQLConn)

            GrvNota.Visible = False
            GrvNota_ByBarang.Visible = False
            GrvSumNota.Visible = True
            GrvDetailNota.Visible = False
        Else
            psFillGrvNota(TxtNotaCompany.Text, TxtNotaNo1.Text, vnSQLConn)

            GrvNota.Visible = True
            GrvNota_ByBarang.Visible = True

            GrvSumNota.Visible = False
            GrvDetailNota.Visible = False
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Protected Sub GrvSumNota_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSumNota.SelectedIndexChanged

    End Sub

    Private Sub GrvSumNota_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumNota.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvSumNota.Rows(vnIdx)
        If e.CommandName = "vDetail" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnKodeBarang As String = vnGRow.Cells(ensColSumNota.vKodeBarang).Text

            LblDetailNota.Text = vnKodeBarang & " - " & vnGRow.Cells(ensColSumNota.vNamaBarang).Text

            psFillGrvDetailNota(TxtNotaCompany.Text, TxtNotaNo1.Text, vnKodeBarang, vnSQLConn)

            GrvDetailNota.Visible = True

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Protected Sub BtnSummary_Click(sender As Object, e As EventArgs) Handles BtnSummary.Click
        Response.Redirect("~/Reporting/WbfDmInvSummary.aspx")
    End Sub
End Class