Imports System.Data.SqlClient
Imports System.IO
Public Class WbfSsoSummaryStock
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSummaryStock"

    Enum ensColList
        CompanyCode = 0
        WarehouseOID = 1
        WarehouseName = 2
        BRGCODE = 3
        BRGNAME = 4
    End Enum

    Private Sub psDefaultDisplay()
        DivStCard.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanStCard.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivInv.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanInv.Style(HtmlTextWriterStyle.Position) = "absolute"


        DivTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanTRB.Style(HtmlTextWriterStyle.Position) = "absolute"

    End Sub
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Reporting"
        If Not IsPostBack Then
            psDefaultDisplay()
            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -2, Date.Now), "dd MMM yyyy")


            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If
            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWarehouse, False, vnSQLConn)
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnCompanyCode As String = DstListCompany.SelectedValue

        Dim vnCutOfDate As String = fbuGetCutOfDate(vnCompanyCode, vnSQLConn)

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnParam As String
        vnParam = "'" & vnCompanyCode & "'," & DstListWarehouse.SelectedValue & ",'" & vnCutOfDate & "','" & TxtListStart.Text & "'"

        vnQuery = "spSsoInvNotPickDone_ByCompanyWhsDate_Table " & vnParam

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        LblMsgListWarehouse.Text = ""
        LblMsgListCompany.Text = ""
        LblMsgListStart.Text = ""
        If Val(DstListWarehouse.SelectedValue) = 0 Then
            LblMsgListWarehouse.Text = "Pilih Warehouse"
            Exit Sub
        End If
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If
        If Not IsDate(TxtListStart.Text) Then
            LblMsgListStart.Text = "Isi Tanggal"
            Exit Sub
        End If
        psFillGrvList()
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvList.Rows.Count Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)

        Dim vnCompanyCode As String = vnGRow.Cells(ensColList.CompanyCode).Text
        Dim vnWarehouseOID As String = vnGRow.Cells(ensColList.WarehouseOID).Text
        Dim vnBrgCode As String = vnGRow.Cells(ensColList.BRGCODE).Text

        If e.CommandName = "vQty_StockCard" Then
            LblStCardTitle.Text = "STOCK CARD " & vnBrgCode

            psFillGrvStCard(vnCompanyCode, vnWarehouseOID, vnBrgCode, TxtListStart.Text, vnSQLConn)

            psShowStCard(True)

        ElseIf e.CommandName = "vQtyInv_Belum_PickingDone" Then
            LblStCardTitle.Text = "INVOICE BELUM PICKING DONE " & vnBrgCode

            psFillGrvInv(vnCompanyCode, vnWarehouseOID, vnBrgCode, TxtListStart.Text, vnSQLConn)

            psShowInv(True)

        ElseIf e.CommandName = "vQtyTRB_Belum_PickingDone" Then
            LblStCardTitle.Text = "TRB BELUM PICKING DONE " & vnBrgCode

            psFillGrvTRB(vnCompanyCode, vnWarehouseOID, vnBrgCode, TxtListStart.Text, vnSQLConn)

            psShowTRB(True)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvStCard(vriCompanyCode As String, vriWarehouseOID As String, vriBrgCode As String, vriStatusDate As String, vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnCutOfDate As String = fbuGetCutOfDate(vriCompanyCode, vriSQLConn)

        Dim vnDtbStCard As New DataTable
        Dim vnCriteria As String
        vnCriteria = "Where cast(sm.CreationDatetime as date)>='" & vnCutOfDate & "' and cast(sm.CreationDatetime as date)<='" & vriStatusDate & "'"
        vnCriteria += vbCrLf & "            and sm.CompanyCode='" & vriCompanyCode & "' and pm.WarehouseOID=" & vriWarehouseOID
        vnCriteria += vbCrLf & "            and sm.BRGCODE='" & vriBrgCode & "'"

        Dim vnQuery As String
        vnQuery = "Select * From ("
        vnQuery += vbCrLf & "Select sm.OID,pm.vStorageInfoHtml,sm.TransCode,tm.TransName,sm.TransOID,"
        vnQuery += vbCrLf & "     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,"
        vnQuery += vbCrLf & "     sm.TransQty"
        vnQuery += vbCrLf & " From Sys_SsoStockCard_TR sm with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm on pm.vStorageOID=sm.StorageOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoTransName_MA tm on tm.TransCode=sm.TransCode"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "UNION"
        vnQuery += vbCrLf & "Select Null OID,''vStorageInfoHtml,''TransCode,''TransName,Null TransOID,"
        vnQuery += vbCrLf & "     'TOTAL'vCreationDatetime,"
        vnQuery += vbCrLf & "     sum(sm.TransQty)"
        vnQuery += vbCrLf & " From Sys_SsoStockCard_TR sm with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm on pm.vStorageOID=sm.StorageOID"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & ")tb"

        vnQuery += vbCrLf & "Order by case when vCreationDatetime='TOTAL' then 19 else 8 end,case when TransCode='STOB' then 4 else 5 end,OID"

        pbuFillDtbSQL(vnDtbStCard, vnQuery, vriSQLConn)
        GrvStCard.DataSource = vnDtbStCard
        GrvStCard.DataBind()

        If vnDtbStCard.Rows.Count > 0 Then
            GrvStCard.Rows(GrvStCard.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
        End If
    End Sub

    Private Sub psFillGrvInv(vriCompanyCode As String, vriWarehouseOID As String, vriBrgCode As String, vriStatusDate As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnCutOfDate As String = fbuGetCutOfDate(vriCompanyCode, vriSQLConn)

        Dim vnDtbInv As New DataTable
        Dim vnCriteria As String
        vnCriteria = "Where sm.IsPickListClosed=0 and sm.NotaCancel=0 and cast(sm.UploadDatetime as date)>='" & vnCutOfDate & "' and cast(sm.UploadDatetime as date)<='" & vriStatusDate & "'"
        vnCriteria += vbCrLf & "    and sm.CompanyCode='" & vriCompanyCode & "' and sm.WarehouseOID=" & vriWarehouseOID
        vnCriteria += vbCrLf & "    and sm.KODE_BARANG='" & vriBrgCode & "'"
        vnCriteria += vbCrLf & "    and not sm.CompanyCode+'x'+cast(sm.WarehouseOID as varchar)+'x'+sm.NO_NOTA in("
        vnCriteria += vbCrLf & "        Select dj.CompanyCode+'x'+cast(pch.WarehouseOID as varchar)+'x'+pch.PCLRefHNo"
        vnCriteria += vbCrLf & "		  From " & vnDBDcm & "Sys_DcmJUAL dj with(nolock)"
        vnCriteria += vbCrLf & "               inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.PCLCompanyCode=dj.CompanyCode and pch.WarehouseOID=dj.WarehouseOID and pch.PCLRefHNo=dj.NO_NOTA"
        vnCriteria += vbCrLf & "               inner join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=pch.OID"
        vnCriteria += vbCrLf & "         Where dj.GDG in('KEA01','PRA01') and pch.SchDTypeOID=" & enuSchDType.Invoice & " and pch.TransStatus>0 and"
        vnCriteria += vbCrLf & "               pck.PickDoneDatetime is not null and cast(pck.PickDoneDatetime as date)<='" & vriStatusDate & "' and"
        vnCriteria += vbCrLf & "               cast(dj.UploadDatetime as date)>='" & vnCutOfDate & "' and cast(dj.UploadDatetime as date)<='" & vriStatusDate & "' and"
        vnCriteria += vbCrLf & "               dj.CompanyCode='" & vriCompanyCode & "' and dj.WarehouseOID=" & vriWarehouseOID & " and"
        vnCriteria += vbCrLf & "               dj.KODE_BARANG='" & vriBrgCode & "')"

        Dim vnQuery As String
        vnQuery = "Select * From ("
        vnQuery += vbCrLf & "Select sm.NO_NOTA,convert(varchar(11),sm.TANGGAL,106) vTANGGAL,sm.KODE_CUST+' '+sm.CUSTOMER vCustomer,"
        vnQuery += vbCrLf & "     convert(varchar(11),sm.UploadDatetime,106)+' '+convert(varchar(8),sm.UploadDatetime,108)vUploadDatetime,"
        vnQuery += vbCrLf & "     cast(sm.QTY as int)QTY,cast(sm.QTYBONUS as int)QTYBONUS"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmJUAL sm with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoGudangAktif_MA gd with(nolock) on gd.CompanyCode=rtrim(sm.CompanyCode) and gd.GdgCode=sm.GDG"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "UNION"
        vnQuery += vbCrLf & "Select Null NO_NOTA,''vTANGGAL,''vCustomer,"
        vnQuery += vbCrLf & "     'TOTAL'vUploadDatetime,"
        vnQuery += vbCrLf & "     cast(sum(sm.QTY)as int),cast(sum(sm.QTYBONUS)as int)"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmJUAL sm with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoGudangAktif_MA gd with(nolock) on gd.CompanyCode=rtrim(sm.CompanyCode) and gd.GdgCode=sm.GDG"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & ")tb"

        vnQuery += vbCrLf & "Order by case when vUploadDatetime='TOTAL' then 19 else 8 end,NO_NOTA"

        pbuFillDtbSQL(vnDtbInv, vnQuery, vriSQLConn)
        GrvInv.DataSource = vnDtbInv
        GrvInv.DataBind()

        If vnDtbInv.Rows.Count > 0 Then
            GrvInv.Rows(GrvInv.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
        End If
    End Sub



    Private Sub psFillGrvTRB(vriCompanyCode As String, vriWarehouseOID As String, vriBrgCode As String, vriStatusDate As String, vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnCutOfDate As String = fbuGetCutOfDate(vriCompanyCode, vriSQLConn)

        Dim vnDtbInv As New DataTable
        Dim vnCriteria As String
        Dim vnQuery As String
        vnQuery = "	DECLARE @vriCompanycode AS varchar(25)	"
        vnQuery += vbCrLf & "	DECLARE @vriWarehouseOID AS varchar(10)	"
        vnQuery += vbCrLf & "	DECLARE @vriCutOfDate AS date	"
        vnQuery += vbCrLf & "	DECLARE @vriStatusDate AS date	"
        vnQuery += vbCrLf & "	DECLARE @vriKodeBarang AS varchar(10)	"
        vnQuery += vbCrLf & "	SET @vriCompanycode = '" & vriCompanyCode & "'	"
        vnQuery += vbCrLf & "	SET @vriWarehouseOID = '" & vriWarehouseOID & "'	"
        vnQuery += vbCrLf & "	SET @vriCutOfDate = '" & vriStatusDate & "'	"
        vnQuery += vbCrLf & "	SET @vriStatusDate = '" & vriStatusDate & "'	"
        vnQuery += vbCrLf & "	SET @vriKodeBarang = '" & vriBrgCode & "'	"
        vnQuery += vbCrLf & "	IF OBJECT_ID('tempdb..#Sys_DcmTRB_Temp') IS NOT NULL	"
        vnQuery += vbCrLf & "	drop table #Sys_DcmTRB_Temp	"
        vnQuery += vbCrLf & "	IF OBJECT_ID('tempdb..#Sys_SsoTRBNotPickDone') IS NOT NULL	"
        vnQuery += vbCrLf & "	drop table #Sys_SsoTRBNotPickDone	"
        vnQuery += vbCrLf & "	IF OBJECT_ID('tempdb..#Sys_SsoTRBNotPickDone') IS NOT NULL	"
        vnQuery += vbCrLf & "	drop table #Sys_SsoTRBNotPickDone	"
        vnQuery += vbCrLf & "	Select trb.* into #Sys_DcmTRB_Temp	"
        vnQuery += vbCrLf & "	   From " & vnDBDcm & "Sys_DcmTRB trb with(nolock)	"
        vnQuery += vbCrLf & "	        inner join " & vnDBDcm & "Sys_DcmTRBHeader_TR trh with(nolock) on trh.CompanyCode=trb.CompanyCode and trh.NoBukti=trb.NoBukti and trh.IsPickListClosed=0	"
        vnQuery += vbCrLf & "	  Where cast(trb.UploadDatetime as date)>=@vriCutOfDate and cast(trb.UploadDatetime as date)<=@vriStatusDate and	"
        vnQuery += vbCrLf & "	trb.CompanyCode=@vriCompanycode and trb.WarehouseAsalOID=@vriWarehouseOID	"
        vnQuery += vbCrLf & "	Select trb.CompanyCode,trb.WarehouseAsalOID,trb.KodeBrg,cast(sum(-1 * trb.QTY)as int) vQtyTRB_Belum_PickingDone into #Sys_SsoTRBNotPickDone	"
        vnQuery += vbCrLf & "	   From #Sys_DcmTRB_Temp trb	"
        vnQuery += vbCrLf & "	  Where not trb.CompanyCode+'x'+cast(trb.WarehouseAsalOID as varchar)+'x'+trb.NoBukti in	"
        vnQuery += vbCrLf & "	(	"
        vnQuery += vbCrLf & "	Select trm.CompanyCode+'x'+cast(pch.WarehouseOID as varchar)+'x'+pch.PCLRefHNo	"
        vnQuery += vbCrLf & "	   From #Sys_DcmTRB_Temp trm	"
        vnQuery += vbCrLf & "	inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.PCLCompanyCode=trm.CompanyCode and pch.WarehouseOID=trm.WarehouseAsalOID and pch.PCLRefHNo=trm.NoBukti	"
        vnQuery += vbCrLf & "	inner join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=pch.OID	"
        vnQuery += vbCrLf & "	  Where pch.SchDTypeOID=5 and pch.TransStatus>0 and pck.PickDoneDatetime is not null and cast(pck.PickDoneDatetime as date)<=@vriStatusDate	"
        vnQuery += vbCrLf & "	)	"
        vnQuery += vbCrLf & "	Group by trb.CompanyCode,trb.WarehouseAsalOID,trb.KodeBrg	"
        vnQuery += vbCrLf & "	Insert into #Sys_SsoTRBNotPickDone	"
        vnQuery += vbCrLf & "	Select @vriCompanycode,@vriWarehouseOID,pck_s.BRGCODE,sum(-1 * pck_s.PCKScanQty)vQtyTRB_Belum_PickingDone	"
        vnQuery += vbCrLf & "	   From Sys_SsoPCKHeader_TR pck_h with(nolock)	"
        vnQuery += vbCrLf & "	inner join Sys_SsoPCKScan_TR pck_s with(nolock) on pck_s.PCKHOID=pck_h.OID	"
        vnQuery += vbCrLf & "	inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sto on sto.vStorageOID=pck_s.StorageOID	"
        vnQuery += vbCrLf & "	inner join Sys_SsoPCLHeader_TR pcl_h with(nolock) on pcl_h.OID=pck_h.PCLHOID	"
        vnQuery += vbCrLf & "	inner join " & vnDBDcm & "Sys_DcmTRBHeader_TR nh with(nolock) on nh.OID=pcl_h.PCLRefHOID	"
        vnQuery += vbCrLf & "	  Where pck_h.PCKCompanyCode=@vriCompanycode and pck_h.WarehouseOID=@vriWarehouseOID and	"
        vnQuery += vbCrLf & "	        sto.StorageTypeOID in(1002,1003) and	"
        vnQuery += vbCrLf & "	cast(isnull(pck_h.PickDoneDatetime,dateadd(d,2,getdate())) as date) > @vriStatusDate and pck_s.PCKScanDeleted=0 and	"
        vnQuery += vbCrLf & "	pcl_h.SchDTypeOID=5 and cast(nh.Tanggal as date)<= @vriStatusDate	"
        vnQuery += vbCrLf & "	  Group by pck_s.BRGCODE	"
        vnQuery += vbCrLf & "	Insert into #Sys_SsoTRBNotPickDone	"
        vnQuery += vbCrLf & "	Select @vriCompanycode,@vriWarehouseOID,pck_s.BRGCODE,sum(pck_s.PCKScanQty)vQtyTRB_Belum_PickingDone	"
        vnQuery += vbCrLf & "	       From Sys_SsoPCKHeader_TR pck_h with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPCKScan_TR pck_s with(nolock) on pck_s.PCKHOID=pck_h.OID	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPCLHeader_TR pcl_h with(nolock) on pcl_h.OID=pck_h.PCLHOID	"
        vnQuery += vbCrLf & "	inner join " & vnDBDcm & "Sys_DcmTRBHeader_TR nh with(nolock) on nh.OID=pcl_h.PCLRefHOID	"
        vnQuery += vbCrLf & "	left outer join Sys_SsoDSRPick_TR dsr_p with(nolock) on dsr_p.PCKHOID=pck_h.OID	"
        vnQuery += vbCrLf & "	left outer join Sys_SsoDSRHeader_TR dsr_h with(nolock) on dsr_h.OID=dsr_p.DSRHOID	"
        vnQuery += vbCrLf & "	  Where pck_h.PCKCompanyCode=@vriCompanycode and	"
        vnQuery += vbCrLf & "	        pcl_h.WarehouseOID_Dest<>pcl_h.WarehouseOID and pcl_h.WarehouseOID_Dest=@vriWarehouseOID and	"
        vnQuery += vbCrLf & "	cast(isnull(dsr_h.DispatchRcvDoneDatetime,dateadd(d,2,getdate())) as date) > @vriStatusDate and pck_s.PCKScanDeleted=0 and	"
        vnQuery += vbCrLf & "	pcl_h.SchDTypeOID=5 and cast(nh.Tanggal as date)<= @vriStatusDate	"
        vnQuery += vbCrLf & "	  Group by pck_s.BRGCODE	"
        vnQuery += vbCrLf & "	  SELECT * FROM #Sys_SsoTRBNotPickDone where kodeBrg = @vriKodeBarang	"


        pbuFillDtbSQL(vnDtbInv, vnQuery, vriSQLConn)
        GrvInv.DataSource = vnDtbInv
        GrvInv.DataBind()

        If vnDtbInv.Rows.Count > 0 Then
            GrvInv.Rows(GrvInv.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
        End If
    End Sub



    Private Sub psShowStCard(vriBo As Boolean)
        If vriBo Then
            DivStCard.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivStCard.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowInv(vriBo As Boolean)
        If vriBo Then
            DivInv.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivInv.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowTRB(vriBo As Boolean)
        If vriBo Then
            DivTRB.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
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
        pbuCreateXlsx_SummaryStock1(vnFileName, Session("UserOID"), DstListWarehouse, DstListCompany, TxtListStart, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
    Private Sub BtnStCardClose_Click(sender As Object, e As EventArgs) Handles BtnStCardClose.Click
        psShowStCard(False)
    End Sub

    Protected Sub BtnInvClose_Click(sender As Object, e As EventArgs) Handles BtnInvClose.Click
        psShowInv(False)
    End Sub
    Protected Sub BtnTRBClose_Click(sender As Object, e As EventArgs) Handles BtnTRBClose.Click
        psShowTRB(False)
    End Sub
End Class