Imports System.Data.SqlClient
Public Class WbfSsoMonStatusInvoice
    Inherits System.Web.UI.Page

    Private Sub psDefaultDisplay()
        'DivLsRcvPO.Style(HtmlTextWriterStyle.MarginTop) = "-175px"


        'DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        'DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        'PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

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

            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub DstListWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstListWarehouse.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        'If DstListWarehouse.SelectedValue = "0" Then
        '    pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)
        'Else
        '    pbuFillDstBuilding_ByWarehouse(DstListBuilding, True, DstListWarehouse.SelectedValue, vnSQLConn)
        'End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""

    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
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
        Dim Transsum As Integer = 0
        Dim vnCompany As String = fbuFormatString(Trim(DstCompany.SelectedValue))
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "	select distinct	"
        vnQuery += vbCrLf & "	mj.CompanyCode,whs.WarehouseName, whs.OID, DATEDIFF(HOUR,mj.UploadDatetime,skh.BackDatetime) as [Durasi_Start_to_End]	"
        vnQuery += vbCrLf & "	, mj.NO_NOTA, mj.TANGGAL, mj.KODE_CUST, mj.CUSTOMER, mj.UploadDatetime	"
        vnQuery += vbCrLf & "	, pch.PCLNo, pch.PCLDate, pch.PCLScheduleDate, pch.CreationDatetime as [Time_Create_Picklist], usr_pch.UserName, pch.PreparedDatetime	"
        vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pch.CreationDatetime) as [Durasi_Upload_to_Create_Picklist]	"
        vnQuery += vbCrLf & "	, pck.PCKNo, pck.PCKDate, pck.CreationDatetime 'Picking_Created_Date_Time', pck.PickDoneDatetime, pch.PCLRefHOID, pch.PCLRefHNo	"
        vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pch.CreationDatetime,pck.PickDoneDatetime) as [Durasi_Picklist_Created_to_Picking_Done]	"
        vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pck.PickDoneDatetime) as [Durasi_Upload_to_Picking_Done]	"
        vnQuery += vbCrLf & "	, dsh.DSPNo, dsh.DSPDate, dsh.CreationDatetime 'Dispatch_Created_Date_Time', dsh.DispatchDoneDatetime 'Dispatch_Created_Date'	"
        vnQuery += vbCrLf & "	, dsh.DriverConfirmDatetime 	"
        vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pck.PickDoneDatetime,dsh.DriverConfirmDatetime) as [Durasi_Picking_Done_to_Dispatch]	"
        vnQuery += vbCrLf & "	,drv.DcmDriverName,skh.BackDatetime, dsh.CancelledDatetime, pch.TransCode , pch.TransStatus, sstsm.TransStatusDescr	"
        vnQuery += vbCrLf & "	from 	"

        vnQuery += vbCrLf & "	(select	"
        vnQuery += vbCrLf & "	ju.CompanyCode, ju.WarehouseOID, ju.NO_NOTA, ju.TANGGAL, ju.KODE_CUST, ju.CUSTOMER, max(ju.UploadDatetime) as uploadDatetime	"
        vnQuery += vbCrLf & "	from " & vnDBDcm & "Sys_DcmJUAL	ju with(nolock)"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=ju.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "	group by CompanyCode, WarehouseOID, NO_NOTA, TANGGAL, KODE_CUST, CUSTOMER	"
        vnQuery += vbCrLf & "	) as mj	"

        vnQuery += vbCrLf & "	left join Sys_SsoPCLHeader_TR pch with(nolock) on pch.PCLRefHNo = mj.NO_NOTA	"
        vnQuery += vbCrLf & "	left join Sys_SsoUser_MA usr_pch with(nolock) on usr_pch.OID = pch.CreationUserOID	"
        vnQuery += vbCrLf & "	left join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID = mj.WarehouseOID	"
        vnQuery += vbCrLf & "	left join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID = pch.OID	"
        vnQuery += vbCrLf & "	left join Sys_SsoDSPPick_TR dsp with(nolock) on dsp.PCKHOID= pck.OID	"
        vnQuery += vbCrLf & "	left join Sys_SsoDSPHeader_TR dsh with(nolock) on dsh.OID= dsp.DSPHOID	"
        vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmDriver_MA drv with(nolock) on drv.OID= dsh.DcmSchDriverOID	"
        vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmScheduleDetail_TR skd with(nolock) on skd.NotaNo=mj.NO_NOTA and skd.SchDTypeOID=1	"
        vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmScheduleHeader_TR skh with(nolock) on skh.OID=skd.DcmSchHOID	"
        vnQuery += vbCrLf & "	LEFT JOIN Sys_SsoTransStatus_MA sstsm with(nolock) ON pch.TransCode = sstsm.TransCode AND pch.TransStatus = sstsm.TransStatus	"
        vnQuery += vbCrLf & "Where 1=1 and LEFT(NO_NOTA,1) <> 'P'"
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and mj.TANGGAL >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and mj.TANGGAL <= '" & TxtListEnd.Text & "'"
        End If
        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & " and mj.WarehouseOID = " & DstListWarehouse.SelectedValue & " "
        End If
        If DstCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and mj.CompanyCode = '" & vnCompany & "'"
        End If
        If Trim(TxtInvoiceNo.Text) <> "" Then
            vnQuery += vbCrLf & " and mj.NO_NOTA like '%" & fbuFormatString(Trim(TxtInvoiceNo.Text)) & "%'"
        End If
        If Trim(TxtPCLNo.Text) <> "" Then
            vnQuery += vbCrLf & " and pch.PCLNo like '%" & fbuFormatString(Trim(TxtPCLNo.Text)) & "%'"
        End If
        If Trim(TxtPCLRefNo.Text) <> "" Then
            vnQuery += vbCrLf & " and pch.PCLRefHNo like '%" & fbuFormatString(Trim(TxtPCLNo.Text)) & "%'"
        End If
        If Trim(TxtPickNo.Text) <> "" Then
            vnQuery += vbCrLf & " and pck.PCKNo like '%" & fbuFormatString(Trim(TxtPickNo.Text)) & "%'"
        End If
        If Trim(TxtDispatchNo.Text) <> "" Then
            vnQuery += vbCrLf & " and dsh.DSPNo like '%" & fbuFormatString(Trim(TxtDispatchNo.Text)) & "%'"
        End If

        If Chk_Upload.Checked = True Then
            vnQuery += vbCrLf & " and mj.UploadDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and mj.UploadDatetime is null "
        End If
        If Chk_Picklist.Checked = True Then
            vnQuery += vbCrLf & " and pch.creationdatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and pch.creationdatetime is null "
        End If

        If Chk_PickilistPrepared.Checked = True Then
            vnQuery += vbCrLf & " and pch.PreparedDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and pch.PreparedDatetime is null "
        End If
        If Chk_Picking.Checked = True Then
            vnQuery += vbCrLf & " and pck.CreationDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and pck.CreationDatetime is null "
        End If
        If Chk_PickingDone.Checked = True Then
            vnQuery += vbCrLf & " and pck.PickDoneDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and pck.PickDoneDatetime is null "
        End If

        If Chk_Dispatch.Checked = True Then
            vnQuery += vbCrLf & " and dsh.CreationDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and dsh.CreationDatetime is null "
        End If
        If Chk_DispatchDone.Checked = True Then
            vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is null "
        End If
        If Chk_DriverConfirm.Checked = True Then
            vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is not null "
            Transsum = Transsum + 1
        Else
            vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is null "
        End If
        vnQuery += vbCrLf & " ORDER BY mj.TANGGAL DESC  "
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
        pbuCreateXlsx_MonInvoice(vnFileName, Session("UserOID"), DstListWarehouse, DstCompany, TxtInvoiceNo, TxtPCLNo, TxtPCLRefNo, TxtPickNo, TxtDispatchNo, TxtListStart, TxtListEnd, Chk_Upload, Chk_Picklist, Chk_PickilistPrepared, Chk_Picking, Chk_PickingDone, Chk_Dispatch, Chk_DispatchDone, Chk_DriverConfirm, Chk_Back, vnSQLConn)
        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub


End Class