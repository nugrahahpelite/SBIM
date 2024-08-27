Imports System.Data.SqlClient
Public Class WbfSsoPickDs
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoPickDs"
    Const csTNoPrefix = "PCK"

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

    Enum ensColSumm
        BRGCODE = 0
        vBRGCODE = 1
        BRGNAME = 2
        ReservedQty = 3
        PCKScanQty = 4
        vQtyVarian = 5
        RcvPONo = 6
        vStorageInfoHtml = 7
        StorageOID = 8
        StorageStockOID = 9
    End Enum
    Enum ensColData
        OID = 0
        BrgCode = 1
        vStorageInfoHtml = 2
        RcvPONo = 3
        PCKScanQty = 4
        vQtySerialNo = 5
        vPCKScanNote = 6
        vPCKScanUser = 7
        vPCKScanTime = 8
        vDelItem = 9
        vPCKScanDeleted = 10
        PCKScanDeletedNote = 11
        vPCKScanDeletedTime = 12
        StorageOID = 13
        RcvPOHOID = 14
        StorageStockOID = 15
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransPLNo.Text = ""
        TxtTransNo.Text = ""

        HdfTransStatus.Value = "0"
        HdfTransStatus.Value = enuTCSSOH.Baru

        BtnCancelPCK.Enabled = False
        BtnCancelPCK.Visible = BtnCancelPCK.Enabled
    End Sub
    Enum ensColLsScan
        vRcvPOScanDeleted = 5
    End Enum
    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPicking, vnSQLConn)

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, False, vnSQLConn)
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
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkSt_OnPicking.Checked = False And ChkSt_PickingDone.Checked = False And ChkSt_OnDispatch.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_Void.Checked = False And ChkSt_PtwVoid.Checked = False And ChkSt_MoveAntarStgOut.Checked = False Then
            ChkSt_OnPicking.Checked = True
            ChkSt_PickingDone.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_OnPicking.Checked = True Then
            vnCrStatus += enuTCPCKG.On_Picking & ","
        End If
        If ChkSt_PickingDone.Checked = True Then
            vnCrStatus += enuTCPCKG.Picking_Done & ","
        End If
        If ChkSt_OnDispatch.Checked = True Then
            vnCrStatus += enuTCPCKG.On_Dispatch_Putaway & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPCKG.Cancelled & ","
        End If
        If ChkSt_DispatchDone.Checked = True Then
            vnCrStatus += enuTCPCKG.Putaway_Dispatch_Done & ","
        End If
        If ChkSt_MoveAntarStgOut.Checked Then
            vnCrStatus += enuTCPCKG.Move_Antar_StagingOut_Done & ","
        End If
        If ChkSt_Void.Checked = True Then
            vnCrStatus += enuTCPCKG.Void & ","
        End If
        If ChkSt_PtwVoid.Checked = True Then
            vnCrStatus += enuTCPCKG.On_Putaway_Void & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select pwh.OID,pwh.PCKCompanyCode,pwh.PCKNo,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
        vnQuery += vbCrLf & "       pch.PCLNo,pch.PCLRefHNo,"
        vnQuery += vbCrLf & "       whs.WarehouseName,pwh.StorageOID,sto.vStorageInfoHtml,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.PickDoneDatetime,106)+' '+convert(varchar(5),pwh.PickDoneDatetime,108)+' '+ RD.UserName vPickDone,PCKCancelNote"
        vnQuery += vbCrLf & "  From Sys_SsoPCKHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "       left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sto on sto.vStorageOID=pwh.StorageOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       inner join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA RD with(nolock) on RD.OID=pwh.PickDoneUserOID"

        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PCKCompanyCode and uc.UserOID=" & Session("UserOID")
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
            vnQuery += vbCrLf & " and (pwh.PCKNo like '%" & Trim(TxtListNo.Text) & "%' or pch.PCLNo like '%" & Trim(TxtListNo.Text) & "%' or pch.PCLRefHNo like '%" & Trim(TxtListNo.Text) & "%')"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PCKDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PCKDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "Order by pwh.PCKNo"
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

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoPicking
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

        vnQuery = "Select pwh.*,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
        vnQuery += vbCrLf & "cmp.CompanyName,wha.WarehouseName,pch.PCLNo,pch.PCLRefHNo,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoPCKHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "DimCompany cmp on cmp.CompanyCode=pwh.PCKCompanyCode"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA wha on wha.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=pwh.TransStatus and ST.TransCode='" & stuTransCode.SsoPicking & "'"

        vnQuery += vbCrLf & "     Where pwh.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vPCKDate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("PCKNo")
            TxtTransPLNo.Text = vnDtb.Rows(0).Item("PCLNo")
            TxtTransPLRef.Text = vnDtb.Rows(0).Item("PCLRefHNo")

            HdfPCLHOID.Value = vnDtb.Rows(0).Item("PCLHOID")
            HdfCompanyCode.Value = vnDtb.Rows(0).Item("PCKCompanyCode")
            TxtCompany.Text = vnDtb.Rows(0).Item("CompanyName")
            TxtWhs.Text = vnDtb.Rows(0).Item("WarehouseName")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        psFillGrvSumm(0, Val(TxtTransID.Text), vriSQLConn)
        PanData.Visible = False
        vnDtb.Dispose()
    End Sub

    Private Sub psButtonStatusDefault()
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

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvSumm(0, TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriPCKHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If vriEmpty = 1 Then
            vnQuery = "Select ''BRGCODE,''BRGNAME,''vIsSN,0 ReservedQty,0 PCKScanQty,0 vQtyVarian,''RcvPONo,''vStorageInfoHtml,0 StorageOID,0 StorageStockOID Where 1=2"
        Else
            vnQuery = "Select pcs.BRGCODE,mb.BRGNAME,case when abs(mb.IsSN)=1 then 'Y' else 'N' end vIsSN,"
            vnQuery += vbCrLf & "       pcs.ReservedQty,pcks.PCKScanQty,(pcs.ReservedQty - isnull(pcks.PCKScanQty,0))vQtyVarian,"
            vnQuery += vbCrLf & "       rch.RcvPONo,sti.vStorageInfoHtml,pcs.StorageOID,pcs.StorageStockOID"
            vnQuery += vbCrLf & "       From Sys_SsoPCLReserve_TR pcs with(nolock)"
            vnQuery += vbCrLf & "            inner join Sys_SsoStorageStock_MA sto with(nolock) on sto.OID=pcs.StorageStockOID"
            vnQuery += vbCrLf & "            inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sto.RcvPOHOID"
            vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pcs.BRGCODE and mb.CompanyCode='" & HdfCompanyCode.Value & "'"
            vnQuery += vbCrLf & "            inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') sti on sti.vStorageOID=pcs.StorageOID"
            vnQuery += vbCrLf & "			 inner join Sys_SsoPCKHeader_TR pckh with(nolock) on pckh.PCLHOID=pcs.PCLHOID"
            vnQuery += vbCrLf & "			 left outer join Sys_SsoPCKScan_TR pcks with(nolock) on pcks.PCKHOID=pckh.OID and pcks.StorageStockOID=pcs.StorageStockOID and abs(pcks.PCKScanDeleted)=0"

            vnQuery += vbCrLf & " Where pckh.OID=" & vriPCKHOID & " and (mb.BRGCODE like '%" & vnCriteria & "%' or mb.BRGNAME like '%" & vnCriteria & "%')"

            vnQuery += vbCrLf & "order by case when isnull(pcks.PCKScanQty,0)=0 then 5 else 1 end,mb.BRGCODE"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
    End Sub
    Private Sub GrvSumm_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumm.RowCommand
        If e.CommandName = "BRGCODE" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvSumm.Rows(vnIdx)

            HdfStorageOID.Value = vnGRow.Cells(ensColSumm.StorageOID).Text
            HdfStoKB.Value = DirectCast(vnGRow.Cells(ensColSumm.BRGCODE).Controls(0), LinkButton).Text

            LblDataTitle.Text = HdfStoKB.Value & " " & vnGRow.Cells(ensColSumm.BRGNAME).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvData(0, HdfStorageOID.Value, HdfStoKB.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanData.Visible = True
        End If
    End Sub


    Private Sub psFillGrvData(vriEmpty As Byte, vriStorageOID As String, vriBrgCode As String, vriSQLConn As SqlClient.SqlConnection)
        If ChkSt_DelNo.Checked = False And ChkSt_DelYes.Checked = False Then
            ChkSt_DelNo.Checked = True
            ChkSt_DelYes.Checked = True
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''BrgCode,''vStorageInfoHtml,''RcvPONo,0 PCKScanQty,0 vQtySerialNo,"
            vnQuery += vbCrLf & "       ''PCKScanNote,"
            vnQuery += vbCrLf & "       ''vPCKScanUser,"
            vnQuery += vbCrLf & "	    ''vPCKScanTime,"
            vnQuery += vbCrLf & "	    ''vDelItem,"
            vnQuery += vbCrLf & "	    ''vPCKScanDeleted,''PCKScanDeletedNote,"
            vnQuery += vbCrLf & "	    ''vPCKScanDeletedTime,"
            vnQuery += vbCrLf & "	    0 StorageOID,0 RcvPOHOID,0 StorageStockOID"
            vnQuery += vbCrLf & "	    Where 1=2"
        Else
            vnQuery = "Select sc.OID,sc.BrgCode,st.vStorageInfoHtml,rcv.RcvPONo,sc.PCKScanQty,isnull(vQtySerialNo,0)vQtySerialNo,"
            vnQuery += vbCrLf & "       sc.PCKScanNote,"
            vnQuery += vbCrLf & "       mu.UserID vPCKScanUser,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PCKScanDatetime,106) + ' ' + convert(varchar(5),sc.PCKScanDatetime,108)vPCKScanTime,"
            vnQuery += vbCrLf & "	    ''vDelItem,"
            vnQuery += vbCrLf & "	    case when abs(PCKScanDeleted)=1 then 'Y' else 'N' end vPCKScanDeleted,PCKScanDeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PCKScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.PCKScanDeletedDatetime,108)vPCKScanDeletedTime,"
            vnQuery += vbCrLf & "	    sc.StorageOID,sc.RcvPOHOID,sc.StorageStockOID"
            vnQuery += vbCrLf & "  From Sys_SsoPCKScan_TR sc with(nolock)"
            vnQuery += vbCrLf & "       left outer join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "       left outer join (Select PCKHOID,PCKSOID,count(1)vQtySerialNo From Sys_SsoPCKScanSN_TR sno with(nolock) where abs(PCKScanSNDeleted)=0 and PCKHOID=" & TxtTransID.Text & " group by PCKHOID,PCKSOID)sno on sno.PCKSOID=sc.OID"
            vnQuery += vbCrLf & "	    inner join " & vnDBMaster & "fnTbl_SsoStorageData('') st on st.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.PCKScanUserOID"
            vnQuery += vbCrLf & " Where sc.PCKHOID=" & Val(TxtTransID.Text)
            vnQuery += vbCrLf & "       and sc.PCKScanNote like '%" & vnCriteria & "%'"

            If Val(vriStorageOID) > 0 Then
                vnQuery += vbCrLf & "       and sc.StorageOID=" & vriStorageOID
            End If
            If Trim(vriBrgCode) <> "" Then
                vnQuery += vbCrLf & "       and sc.BrgCode='" & vriBrgCode & "'"
            End If

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(sc.PCKScanDeleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(sc.PCKScanDeleted)=1"
                End If
            End If
            vnQuery += vbCrLf & " Order by sc.OID"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData.DataSource = vnDtb
        GrvData.DataBind()

        If ChkSt_DelYes.Checked = True Then
            Dim vn As Integer
            For vn = 0 To GrvData.Rows.Count - 1
                If GrvData.Rows(vn).Cells(ensColData.vPCKScanDeleted).Text = "Y" Then
                    GrvData.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Protected Sub ChkSt_DelYes_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelYes.CheckedChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvData(0, HdfStorageOID.Value, HdfStoKB.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkSt_DelNo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelNo.CheckedChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvData(0, HdfStorageOID.Value, HdfStoKB.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnCancelPCK_Click(sender As Object, e As EventArgs) Handles BtnCancelPCK.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psClearMessage()
        LblConfirmMessage.Text = "Anda Membatalkan Picking " & TxtTransNo.Text & " ?<br />WARNING : Batal Picking Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelPCK"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        psStatusRefresh()
        If HdfProcess.Value = "CancelPCK" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelPCK()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub
    Private Sub psStatusRefresh()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfTransStatus.Value = fbuGetPCKTransStatus(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnCancelPCK.Enabled = (HdfTransStatus.Value = enuTCPCKG.On_Picking)
            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnCancelPCK.Visible = BtnCancelPCK.Enabled
    End Sub

    Private Sub psCancelPCK()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psCancelPCK", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnPCKHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoPCKHeader_TR with(nolock) Where OID=" & vnPCKHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus <> enuTCPCKG.On_Picking Then
                LblMsgError.Text = "Status <> ON Picking (" & enuTCPCKG.On_Picking & ")"
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

            vnQuery = "Select count(1) From Sys_SsoPCKScan_TR with(nolock) Where PCKHOID=" & vnPCKHOID & " and PCKScanDeleted=0"
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgError.Text = "Batal Picking Gagal... Ada Item Belum Dihapus"
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

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True
            vnQuery = "Update Sys_SsoPCKHeader_TR set TransStatus=" & enuTCPCKG.Cancelled & ",PCKCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnPCKHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPCK...Start")
            pbuInsertStatusPCK(vnPCKHOID, enuTCPCKG.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCK...End")

            vnQuery = "Update Sys_SsoPCLHeader_TR set TransStatus=" & enuTCPICK.Prepared & " Where OID=" & HdfPCLHOID.Value
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPCL...Start")
            pbuInsertStatusPCL(HdfPCLHOID.Value, enuTCPICK.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCL...End")

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

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub
End Class