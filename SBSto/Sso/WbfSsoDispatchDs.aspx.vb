Imports System.Data.SqlClient
Public Class WbfSsoDispatchDs
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoDispatchDs"
    Const csTNoPrefix = "DSP"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String

    Enum ensColList
        OID = 0
    End Enum

    Enum ensColLsPick
        vPCKHOID = 0
        PCKNo = 1
        vPCKDate = 2
        PCLNo = 3
        PCLRefHNo = 4
        SchDTypeName = 5
        vIsQtyConfirm = 6
        vDelItem = 7
    End Enum
    Enum ensColData
        vDPSOID = 0
        BrgCode = 1
        BrgName = 2
        RcvPONo = 3
        RcvPOHOID = 4
        vSumPCKScanQty = 5
        DSPScanQty = 6
        vConfirm = 7
        vNotConfirm = 8
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransDriver.Text = ""
        TxtTransNo.Text = ""

        HdfTransStatus.Value = "0"
        HdfTransStatus.Value = enuTCSSOH.Baru
    End Sub
    Enum ensColLsScan
        vRcvPOScanDeleted = 5
    End Enum
    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrOption.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo

        BtnCancelDSP.Visible = Not vriBo

        BtnList.Visible = Not vriBo

    End Sub
    Private Sub psEnableInput(vriBo As Boolean)
        TxtTransDriver.Visible = Not vriBo
        TxtTransVehicle.Visible = Not vriBo

        DstDriver.Visible = vriBo
        DstVehicle.Visible = vriBo

        TxtNote.ReadOnly = Not vriBo

        If vriBo Then
            TxtNote.BackColor = Drawing.Color.White
            DstDriver.SelectedValue = HdfDriverOID.Value
            DstVehicle.SelectedValue = HdfVehicleOID.Value
        Else
            TxtNote.BackColor = TxtTransNo.BackColor
        End If
    End Sub
    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Trim(TxtTransID.Text) = "" Then Exit Sub
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
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

        Dim vnTransStatus As Integer
        vnTransStatus = fbuGetDSPTransStatus(TxtTransID.Text, vnSQLConn)
        If vnTransStatus < enuTCDISP.On_Dispatch Or vnTransStatus > enuTCDISP.Dispatch_Done Then
            LblMsgError.Text = "Status Sudah Driver Confirm/Cancelled"
            LblMsgError.Visible = True
        Else
            psEnableSave(True)
            psEnableInput(True)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Sso"
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoDispatch, vnSQLConn)
            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, False, vnSQLConn)

            pbuFillDstDcmDriver(DstDriver, False, vnSQLConn)
            pbuFillDstDcmVehicle(DstVehicle, False, vnSQLConn)

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
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkSt_OnDispatch.Checked = False And ChkSt_DispatchDone.Checked = False And ChkSt_DriverConfirm.Checked = False And ChkSt_Closed.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_OnDispatch.Checked = True
            ChkSt_DispatchDone.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_OnDispatch.Checked = True Then
            vnCrStatus += enuTCDISP.On_Dispatch & ","
        End If
        If ChkSt_DispatchDone.Checked = True Then
            vnCrStatus += enuTCDISP.Dispatch_Done & ","
        End If
        If ChkSt_DriverConfirm.Checked = True Then
            vnCrStatus += enuTCDISP.Driver_Confirm & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCDISP.Closed & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCDISP.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select pwh.OID,pwh.DSPCompanyCode,pwh.DSPNo,convert(varchar(11),pwh.DSPDate,106)vDSPDate,"
        vnQuery += vbCrLf & "       wha.WarehouseName,dm.DcmDriverName,vm.VehicleNo,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.DispatchDoneDatetime,106)+' '+convert(varchar(5),pwh.DispatchDoneDatetime,108)+' '+ RD.UserName vDispatchDone,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.DriverConfirmDatetime,106)+' '+convert(varchar(5),pwh.DriverConfirmDatetime,108)vDriverConfirm"
        vnQuery += vbCrLf & "  From Sys_SsoDSPHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA wha on wha.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "       left outer join " & vnDBDcm & "Sys_DcmDriver_MA dm with(nolock) on dm.OID=pwh.DcmSchDriverOID"
        vnQuery += vbCrLf & "       left outer join " & vnDBDcm & "Sys_DcmVehicle_MA vm with(nolock) on vm.OID=pwh.DcmVehicleOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       inner join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA RD with(nolock) on RD.OID=pwh.DispatchDoneUserOID"

        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.DSPCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and pwh.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and pwh.DSPNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.DSPDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.DSPDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "Order by pwh.DSPNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        psFillGrvList()
    End Sub

    Protected Sub BtnListClose_Click(sender As Object, e As EventArgs) Handles BtnListClose.Click
        psShowList(False)
    End Sub

    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowList(vriBo As Boolean)
        If vriBo Then
            DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psFillGrvList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
        End If
    End Sub

    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        If Not IsDate(TxtListStart.Text) Then
            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -14, Date.Now), "dd MMM yyyy")
        End If
        If Not IsDate(TxtListEnd.Text) Then
            TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")
        End If
        psShowList(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblMsgDcsDriver.Text = ""
        LblMsgDcsVehicle.Text = ""
        LblXlsProses.Text = ""
    End Sub

    Private Sub psShowPrOption(vriBo As Boolean)
        If vriBo Then
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoDispatch
            vnParam += "&vqTrNo=" & TxtTransNo.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)

            'Dim vnWinOpen As String
            'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
            'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
            'vnClientScript = Nothing
        End If
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()

        vnQuery = "Select pwh.*,convert(varchar(11),pwh.DSPDate,106)vDSPDate,"
        vnQuery += vbCrLf & "cmp.CompanyName,wha.WarehouseName,dm.DcmDriverName,vm.VehicleNo,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoDSPHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "DimCompany cmp on cmp.CompanyCode=pwh.DSPCompanyCode"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA wha on wha.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBDcm & "Sys_DcmDriver_MA dm with(nolock) on dm.OID=pwh.DcmSchDriverOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBDcm & "Sys_DcmVehicle_MA vm with(nolock) on vm.OID=pwh.DcmVehicleOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=pwh.TransStatus and ST.TransCode='" & stuTransCode.SsoDispatch & "'"

        vnQuery += vbCrLf & "     Where pwh.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vDSPDate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("DSPNo")

            HdfCompanyCode.Value = vnDtb.Rows(0).Item("DSPCompanyCode")
            TxtCompany.Text = vnDtb.Rows(0).Item("CompanyName")
            TxtWhs.Text = vnDtb.Rows(0).Item("WarehouseName")

            TxtTransDriver.Text = vnDtb.Rows(0).Item("DcmDriverName")
            TxtTransVehicle.Text = vnDtb.Rows(0).Item("VehicleNo")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")
            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            HdfDriverOID.Value = vnDtb.Rows(0).Item("DcmSchDriverOID")
            HdfVehicleOID.Value = vnDtb.Rows(0).Item("DcmVehicleOID")
        End If

        psFillGrvLsPick(0, vriSQLConn)

        psButtonStatus()

        If BtnCancelDSP.Enabled Then
            vnQuery = "Select count(1) FROM Sys_SsoDSPPick_TR Where DSPHOID=" & TxtTransID.Text
            If fbuGetDataNumSQL(vnQuery, vriSQLConn) > 0 Then
                BtnCancelDSP.Enabled = False
                BtnCancelDSP.Visible = BtnCancelDSP.Enabled = False
            End If
        End If

        PanData.Visible = False
        vnDtb.Dispose()
    End Sub

    Private Sub psButtonStatusDefault()
        BtnEdit.Enabled = False
        BtnSimpan.Enabled = False
        BtnBatal.Enabled = False

        BtnCancelDSP.Enabled = False
        BtnList.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "Select" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)
            TxtTransID.Text = vnRow.Cells(ensColList.OID).Text

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

            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowList(False)
        End If
    End Sub
    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnProCancel_Click(sender As Object, e As EventArgs) Handles BtnProCancel.Click
        psShowPrOption(False)
    End Sub

    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
    End Sub

    Private Sub psFillGrvLsPick(vriEmpty As Byte, vriSQLConn As SqlConnection)
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''PCKNo,''vPCKDate,''PCLNo,''PCLRefHNo,''SchDTypeName,''vIsQtyConfirm,''vDelItem Where 1=2"
        Else
            vnQuery = "Select pwh.OID vPCKHOID,pwh.PCKNo,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
            vnQuery += vbCrLf & "       pch.PCLNo,pch.PCLRefHNo,msc.SchDTypeName,case when abs(pcp.IsQtyConfirm)=1 then 'Confirm' else 'Not Confirm' end vIsQtyConfirm,"
            vnQuery += vbCrLf & "       ''vDelItem"
            vnQuery += vbCrLf & "  From Sys_SsoPCKHeader_TR pwh with(nolock)"
            vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBDcm() & "Sys_DcmSchDType_MA msc with(nolock) on msc.OID=pch.SchDTypeOID"
            vnQuery += vbCrLf & "       inner join Sys_SsoDSPPick_TR pcp with(nolock) on pcp.PCKHOID=pwh.OID"
            vnQuery += vbCrLf & " Where pcp.DSPHOID=" & TxtTransID.Text
            vnQuery += vbCrLf & " Order by pch.PCLNo"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvLsPick.DataSource = vnDtb
        GrvLsPick.DataBind()
    End Sub

    Private Sub psFillGrvData(vriEmpty As Byte, vriPCKHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 vDPSOID,''BrgCode,''BrgName,''RcvPONo,0 RcvPOHOID,0 vSumPCKScanQty,0 DSPScanQty,"
            vnQuery += vbCrLf & "       ''vConfirm,"
            vnQuery += vbCrLf & "       ''vNotConfirm"
            vnQuery += vbCrLf & "	    Where 1=2"
        Else
            vnQuery = "Select isnull(dps.OID,0)vDPSOID,sc.BrgCode,mbr.BrgName,rcv.RcvPONo,sc.RcvPOHOID,vSumPCKScanQty,isnull(DSPScanQty,0)DSPScanQty,"
            vnQuery += vbCrLf & "       ''vConfirm,"
            vnQuery += vbCrLf & "       ''vNotConfirm"
            vnQuery += vbCrLf & "  From (Select isc.BrgCode,isc.RcvPOHOID,sum(isc.PCKScanQty)vSumPCKScanQty From Sys_SsoPCKScan_TR isc with(nolock) where abs(isc.PCKScanDeleted)=0 and isc.PCKHOID=" & vriPCKHOID & " group by isc.BrgCode,isc.RcvPOHOID)sc"
            vnQuery += vbCrLf & "       inner join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mbr with(nolock)on mbr.BrgCode=sc.BrgCode and mbr.CompanyCode=rcv.RcvPOCompanyCode"
            vnQuery += vbCrLf & "       left outer join (Select idps.OID,idps.BrgCode,idps.RcvPOHOID,DSPScanQty From Sys_SsoDSPScan_TR idps with(nolock) where idps.DSPHOID=" & TxtTransID.Text & " and idps.PCKHOID=" & vriPCKHOID & ")dps on dps.BRGCODE=sc.BRGCODE and dps.RcvPOHOID=sc.RcvPOHOID"

            vnQuery += vbCrLf & " Order by sc.BrgCode,rcv.RcvPONo"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData.DataSource = vnDtb
        GrvData.DataBind()

        Dim vn As Integer
        For vn = 0 To GrvData.Rows.Count - 1
            If GrvData.Rows(vn).Cells(ensColData.DSPScanQty).Text <> GrvData.Rows(vn).Cells(ensColData.vSumPCKScanQty).Text Then
                GrvData.Rows(vn).ForeColor = Drawing.Color.Red
            End If
        Next
    End Sub

    Private Sub GrvLsPick_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsPick.RowCommand
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvLsPick.Rows(vnIdx)
        HdfPCKHOID.Value = vnGRow.Cells(ensColLsPick.vPCKHOID).Text

        If e.CommandName = "PCKNo" Then
            LblDataTitle.Text = HdfPCKHOID.Value & " " & DirectCast(vnGRow.Cells(ensColLsPick.PCKNo).Controls(0), LinkButton).Text & " " & vnGRow.Cells(ensColLsPick.PCLRefHNo).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                pbMsgError = ""
                Exit Sub
            End If

            psFillGrvData(0, HdfPCKHOID.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanData.Visible = True
        End If
    End Sub

    Protected Sub GrvLsPick_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsPick.SelectedIndexChanged

    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Protected Sub BtnCancelDSP_Click(sender As Object, e As EventArgs) Handles BtnCancelDSP.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Dispatch No. " & TxtTransID.Text & " ?<br />WARNING : Batal Dispatch Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelDSP"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub
    Private Sub psShowConfirm(vriBo As Boolean)
        If vriBo Then
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "visible"
            BtnConfirmYes.Visible = True
            BtnConfirmNo.Text = "NO"
        Else
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Protected Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelDSP" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelDSP()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub
    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnEdit.Enabled = (HdfTransStatus.Value >= enuTCDISP.On_Dispatch And HdfTransStatus.Value <= enuTCDISP.Dispatch_Done)
            BtnCancelDSP.Enabled = (HdfTransStatus.Value = enuTCDISP.On_Dispatch)

            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnEdit.Visible = BtnEdit.Enabled
        BtnCancelDSP.Visible = BtnCancelDSP.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psCancelDSP()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True

            vsTextStream.WriteLine("Open SQL Connection Error....")
            vsTextStream.WriteLine(pbMsgError)
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnDSPHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnCount As Integer
            Dim vnQuery As String
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("fbuGetDSPTransStatus")
            vnTransStatus = fbuGetDSPTransStatus(vnDSPHOID, vnSQLConn)

            If vnTransStatus <> enuTCDISP.On_Dispatch Then
                LblMsgError.Text = "Status Sudah Batal atau Done"
                LblMsgError.Visible = True

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                psDisplayData(vnSQLConn)
                Exit Sub
            End If

            vnQuery = "Select count(OID) FROM Sys_SsoDSPPick_TR Where DSPHOID=" & vnDSPHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnCount = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnCount > 0 Then
                LblMsgError.Text = "Sudah ada picklist dipilih"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                psDisplayData(vnSQLConn)
                Exit Sub
            End If

            vnQuery = "Select count(OID) FROM Sys_SsoDSPScan_TR Where DSPHOID=" & vnDSPHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnCount = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnCount > 0 Then
                LblMsgError.Text = "Sudah ada picklist dipilih"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                psDisplayData(vnSQLConn)
                Exit Sub
            End If

            Dim vnDtbDSPH As New DataTable
            pbuFillDtbDSPHeader(vnDtbDSPH, vnDSPHOID, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            pbuSsoProcessDataKey(HdfProcessDataKey.Value, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoDSPHeader_TR set TransStatus=" & enuTCDISP.Cancelled & ",DSPCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusDSP...Start")
            pbuInsertStatusDSP(vnDSPHOID, vnDtbDSPH.Rows(0).Item("DcmSchDriverOID"), vnDtbDSPH.Rows(0).Item("DcmVehicleOID"), fbuFormatString(Trim(TxtConfirmNote.Text)), enuTCDISP.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusDSP...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psDisplayData(vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        psSaveEdit()
    End Sub

    Protected Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
        psEnableSave(False)
        psButtonVisible()

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psDisplayData(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psSaveEdit()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If

        Dim vnSave As Boolean = True
        psClearMessage()
        If DstDriver.SelectedValue = "0" Then
            LblMsgDcsDriver.Text = "Pilih Driver"
            vnSave = False
        End If
        If DstVehicle.SelectedValue = "0" Then
            LblMsgDcsVehicle.Text = "Pilih Kendaraan"
            vnSave = False
        End If

        If Not vnSave Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnDSPHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            vnTransStatus = fbuGetDSPTransStatus(vnDSPHOID, vnSQLConn)

            If vnTransStatus < enuTCDISP.On_Dispatch Or vnTransStatus > enuTCDISP.Dispatch_Done Then
                LblMsgError.Text = "Status Sudah Batal atau Driver Confirm"
                LblMsgError.Visible = True

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Exit Sub
            End If

            Dim vnQuery As String
            Dim vnDriverOID As String = DstDriver.SelectedValue
            Dim vnVehicleOID As String = DstVehicle.SelectedValue

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoDSPHeader_TR set"
            vnQuery += vbCrLf & "DcmSchDriverOID=" & vnDriverOID & ","
            vnQuery += vbCrLf & "DcmVehicleOID=" & vnVehicleOID & ","
            vnQuery += vbCrLf & "DSPNote='" & fbuFormatString(Trim(TxtNote.Text)) & "',"
            vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
            vnQuery += vbCrLf & "Where OID=" & vnDSPHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            If vnTransStatus = enuTCDISP.Driver_Confirm Then
                vnTransStatus = enuTCDISP.Dispatch_Done
            End If
            pbuInsertStatusDSP(vnDSPHOID, vnDriverOID, vnVehicleOID, fbuFormatString(Trim(TxtNote.Text)), vnTransStatus, Session("UserOID"), vnSQLConn, vnSQLTrans)

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            psEnableInput(False)
            psEnableSave(False)

            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub
End Class