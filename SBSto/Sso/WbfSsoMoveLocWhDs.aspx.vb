Imports System.Data.SqlClient
Public Class WbfSsoMoveLocWhDs
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoMoveLocWhDs"
    Const csTNoPrefix = "PDW"

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
        RcvPOHOID = 0
        RcvPONo = 1
        BRGCODE = 2
        BRGNAME = 3
        vSumPDWScan1Qty = 4
        PDWReceiveQty = 5
        vSumPDWScan2Qty = 6
    End Enum
    Enum ensColData1
        OID = 0
        vStorageInfoHtml = 1
        RcvPONo = 2
        PDWScan1Qty = 3
        PDWScan1Note = 4
        vPDWScan1User = 5
        vPDWScan1Time = 6
        vDelItem1 = 7
        vPDWScan1Deleted = 8
        vPDWScan1DeletedUser = 9
        PDWScan1DeletedNote = 10
        vPDWScan1DeletedTime = 11
    End Enum

    Enum ensColData2
        OID = 0
        vStorageInfoHtml = 1
        RcvPONo = 2
        PDWScan2Qty = 3
        vPDWScan2Note = 4
        vPDWScan2User = 5
        vPDWScan2Time = 6
        vDelItem2 = 7
        vPDWScan2Deleted = 8
        vPDWScan2DeletedUser = 9
        PDWScan2DeletedNote = 10
        vPDWScan2DeletedTime = 11
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransNo.Text = ""
        TxtTransWhsName.Text = ""
        TxtCompany.Text = ""

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPindahLokasi_Antar_Wh, vnSQLConn)

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
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkSt_Baru.Checked = False And ChkSt_OnDelivery.Checked = False And ChkSt_OnMovement.Checked = False And ChkSt_MovementDone.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_OnDelivery.Checked = True
            ChkSt_OnMovement.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCPDLW.Baru & ","
        End If
        If ChkSt_OnDelivery.Checked = True Then
            vnCrStatus += enuTCPDLW.On_Delivery_Movement & ","
        End If
        If ChkSt_OnMovement.Checked = True Then
            vnCrStatus += enuTCPDLW.On_Movement & ","
        End If
        If ChkSt_MovementDone.Checked = True Then
            vnCrStatus += enuTCPDLW.Movement_Done & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPDTW.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select pwh.OID,pwh.PDWCompanyCode,pwh.PDWNo,convert(varchar(11),pwh.PDWDate,106)vPDWDate,"
        vnQuery += vbCrLf & "       whs.WarehouseName,whd.WarehouseName vWarehouseName_Dest,pwh.PDWCancelNote,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnDeliveryMovementDatetime,106)+' '+convert(varchar(5),pwh.OnDeliveryMovementDatetime,108)+' '+ OD.UserName vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnMovementDatetime,106)+' '+convert(varchar(5),pwh.OnMovementDatetime,108)+' '+ OP.UserName vOnMovement,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.MovementDoneDatetime,106)+' '+convert(varchar(5),pwh.MovementDoneDatetime,108)+' '+ PD.UserName vMovementDone"
        vnQuery += vbCrLf & "  From Sys_SsoPDWHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA whd with(nolock) on whd.OID=pwh.WarehouseOID_Dest"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OD with(nolock) on OD.OID=pwh.OnDeliveryMovementUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnMovementUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.MovementDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PDWCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and pwh.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and pwh.PDWNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDWDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDWDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "     Order by pwh.PDWNo"

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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPindahLokasi_Antar_Wh
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

        vnQuery = "Select pwh.OID,pwh.PDWNo,pwh.PDWCompanyCode,pwh.WarehouseOID,convert(varchar(11),pwh.PDWDate,106)vPDWDate,"
        vnQuery += vbCrLf & "      whs.WarehouseName,wht.WarehouseName vWarehouseDest,pwh.PDWDoneNote,pwh.TransStatus,ST.TransStatusDescr"
        vnQuery += vbCrLf & " From Sys_SsoPDWHeader_TR pwh"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA whs on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA wht on wht.OID=pwh.WarehouseOID_Dest"
        vnQuery += vbCrLf & "      inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=pwh.TransStatus and ST.TransCode='" & stuTransCode.SsoPindahLokasi_Antar_Wh & "'"
        vnQuery += vbCrLf & "Where pwh.OID=" & TxtTransID.Text

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vPDWDate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("PDWNo")
            TxtTransWhsName.Text = vnDtb.Rows(0).Item("WarehouseName")
            TxtTransWhsNameDest.Text = vnDtb.Rows(0).Item("vWarehouseDest")

            TxtCompany.Text = vnDtb.Rows(0).Item("PDWCompanyCode")

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
        BtnCancelPDW.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "PDWNo" Then
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

    Protected Sub BtnProCancel_Click(sender As Object, e As EventArgs) Handles BtnProCancel.Click
        psShowPrOption(False)
    End Sub


    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
    End Sub

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriPDWHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If vriEmpty = 1 Then
            vnQuery = "Select 0 RcvPOHOID,''RcvPONo,''BRGCODE,''BRGNAME,0 vSumPDWScan1Qty,0 PDWReceiveQty,0 vSumPDWScan2Qty Where 1=2"
        Else
            vnQuery = "Select pws1.RcvPOHOID,pws1.RcvPONo,mb.BRGCODE,mb.BRGNAME,pws1.vSumPDWScan1Qty,pwr.PDWReceiveQty,pws2.vSumPDWScan2Qty"
            vnQuery += vbCrLf & "From fnTbl_SsoPDWHeaderScan1(" & vriPDWHOID & ",0) pws1"
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws1.PDWCompanyCode"
            vnQuery += vbCrLf & "     left outer join Sys_SsoPDWReceive_TR pwr on pwr.PDWHOID=pws1.PDWHOID and pwr.BRGCODE=pws1.BRGCODE and pwr.RcvPOHOID=pws1.RcvPOHOID"
            vnQuery += vbCrLf & "     left outer join fnTbl_SsoPDWHeaderScan2(" & vriPDWHOID & ",0) pws2 on pws2.PDWHOID=pws1.PDWHOID and pws2.RcvPOHOID=pws1.RcvPOHOID and pws2.BRGCODE=pws1.BRGCODE and pws2.PDWCompanyCode=pws1.PDWCompanyCode"

            vnQuery += vbCrLf & " Where (mb.BRGCODE like '%" & vnCriteria & "%' or mb.BRGNAME like '%" & vnCriteria & "%')"
            vnQuery += vbCrLf & "order by case when isnull(vSumPDWScan2Qty,0)=0 then 5 else 1 end,mb.BRGCODE"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
    End Sub

    Protected Sub GrvSumm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSumm.SelectedIndexChanged

    End Sub

    Private Sub GrvSumm_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumm.RowCommand
        If e.CommandName = "BRGCODE" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvSumm.Rows(vnIdx)

            HdfStoKB.Value = DirectCast(vnGRow.Cells(ensColSumm.BRGCODE).Controls(0), LinkButton).Text
            HdfRcvPOHOID.Value = vnGRow.Cells(ensColSumm.RcvPOHOID).Text

            LblDataTitle.Text = vnGRow.Cells(ensColSumm.RcvPONo).Text & " " & HdfStoKB.Value & " " & vnGRow.Cells(ensColSumm.BRGNAME).Text

            psFillGrvData()
        End If
    End Sub

    Private Sub psFillGrvData()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If RdbDataScan.SelectedValue = "S1" Then
            psFillGrvData1(0, HdfRcvPOHOID.Value, HdfStoKB.Value, vnSQLConn)
            GrvData1.Visible = True
            GrvData2.Visible = False
        Else
            psFillGrvData2(0, HdfRcvPOHOID.Value, HdfStoKB.Value, vnSQLConn)
            GrvData1.Visible = False
            GrvData2.Visible = True
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        PanData.Visible = True
    End Sub

    Protected Sub ChkSt_DelYes_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelYes.CheckedChanged
        psFillGrvData()
    End Sub

    Protected Sub ChkSt_DelNo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelNo.CheckedChanged
        psFillGrvData()
    End Sub

    Protected Sub RdbDataScan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdbDataScan.SelectedIndexChanged
        psFillGrvData()
    End Sub


    Private Sub psFillGrvData1(vriEmpty As Byte, vriRcvPOHOID As String, vriBrgCode As String, vriSQLConn As SqlClient.SqlConnection)
        If ChkSt_DelNo.Checked = False And ChkSt_DelYes.Checked = False Then
            ChkSt_DelNo.Checked = True
            ChkSt_DelYes.Checked = True
        End If

        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''vStorageInfoHtml,''RcvPONo,0 PDWScan1Qty,"
            vnQuery += vbCrLf & "       ''PDWScan1Note,"
            vnQuery += vbCrLf & "       ''vPDWScan1User,"
            vnQuery += vbCrLf & "	    ''vPDWScan1Time,"
            vnQuery += vbCrLf & "	    ''vPDWScan1Deleted,''PDWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    ''vPDWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem1 Where 1=2"
        Else
            vnQuery = "Select sc.OID,st.vStorageInfoHtml,rcv.RcvPONo,sc.PDWScan1Qty,"
            vnQuery += vbCrLf & "       sc.PDWScan1Note,"
            vnQuery += vbCrLf & "       mu.UserID vPDWScan1User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PDWScan1Datetime,106) + ' ' + convert(varchar(5),sc.PDWScan1Datetime,108)vPDWScan1Time,"
            vnQuery += vbCrLf & "	    case when abs(PDWScan1Deleted)=1 then 'Y' else 'N' end vPDWScan1Deleted,PDWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PDWScan1DeletedDatetime,106) + ' ' + convert(varchar(5),sc.PDWScan1DeletedDatetime,108)vPDWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem1"
            vnQuery += vbCrLf & "  From Sys_SsoPDWScan1_TR sc with(nolock)"
            vnQuery += vbCrLf & "       left outer join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('') st on st.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.PDWScan1UserOID"
            vnQuery += vbCrLf & " Where sc.PDWHOID=" & TxtTransID.Text & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
            vnQuery += vbCrLf & "       and sc.PDWScan1Note like '%" & vnCriteria & "%' and sc.RcvPOHOID=" & vriRcvPOHOID

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(PDWScan1Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(PDWScan1Deleted)=1"
                End If
            End If
            vnQuery += vbCrLf & " Order by sc.OID"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData1.DataSource = vnDtb
        GrvData1.DataBind()

        If ChkSt_DelYes.Checked = True Then
            Dim vn As Integer
            For vn = 0 To GrvData1.Rows.Count - 1
                If GrvData1.Rows(vn).Cells(ensColData1.vPDWScan1Deleted).Text = "Y" Then
                    GrvData1.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Private Sub psFillGrvData2(vriEmpty As Byte, vriRcvPOHOID As String, vriBrgCode As String, vriSQLConn As SqlClient.SqlConnection)
        If ChkSt_DelNo.Checked = False And ChkSt_DelYes.Checked = False Then
            ChkSt_DelNo.Checked = True
            ChkSt_DelYes.Checked = True
        End If

        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''vStorageInfoHtml,''RcvPONo,0 PDWScan2Qty,"
            vnQuery += vbCrLf & "       ''PDWScan2Note,"
            vnQuery += vbCrLf & "       ''vPDWScan2User,"
            vnQuery += vbCrLf & "	    ''vPDWScan2Time,"
            vnQuery += vbCrLf & "	    ''vPDWScan2Deleted,''PDWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    ''vPDWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem2 Where 1=2"
        Else
            vnQuery = "Select sc.OID,st.vStorageInfoHtml,rcv.RcvPONo,sc.PDWScan2Qty,"
            vnQuery += vbCrLf & "       sc.PDWScan2Note,"
            vnQuery += vbCrLf & "       mu.UserID vPDWScan2User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PDWScan2Datetime,106) + ' ' + convert(varchar(5),sc.PDWScan2Datetime,108)vPDWScan2Time,"
            vnQuery += vbCrLf & "	    case when abs(PDWScan2Deleted)=1 then 'Y' else 'N' end vPDWScan2Deleted,PDWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.PDWScan2DeletedDatetime,106) + ' ' + convert(varchar(5),sc.PDWScan2DeletedDatetime,108)vPDWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem2"
            vnQuery += vbCrLf & "  From Sys_SsoPDWScan2_TR sc with(nolock)"
            vnQuery += vbCrLf & "       left outer join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('') st on st.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.PDWScan2UserOID"
            vnQuery += vbCrLf & " Where sc.PDWHOID=" & TxtTransID.Text & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
            vnQuery += vbCrLf & "       and sc.PDWScan2Note like '%" & vnCriteria & "%' and sc.RcvPOHOID=" & vriRcvPOHOID

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(PDWScan2Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(PDWScan2Deleted)=1"
                End If
            End If
            vnQuery += vbCrLf & " Order by sc.OID"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData2.DataSource = vnDtb
        GrvData2.DataBind()

        If ChkSt_DelYes.Checked = True Then
            Dim vn As Integer
            For vn = 0 To GrvData2.Rows.Count - 1
                If GrvData2.Rows(vn).Cells(ensColData2.vPDWScan2Deleted).Text = "Y" Then
                    GrvData2.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Protected Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
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

    Protected Sub BtnCancelPDW_Click(sender As Object, e As EventArgs) Handles BtnCancelPDW.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Pindah Lokasi No. " & TxtTransID.Text & " ?<br />WARNING : Batal Pindah Lokasi Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelPDW"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelPDW" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelPDW()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnCancelPDW.Enabled = (HdfTransStatus.Value = enuTCPDLW.Baru Or HdfTransStatus.Value = enuTCPDLW.On_Delivery_Movement)
            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnCancelPDW.Visible = BtnCancelPDW.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub

    Private Sub psCancelPDW()
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
            Dim vnPDWHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnCount1 As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoPDWHeader_TR with(nolock) Where OID=" & vnPDWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus = enuTCPDLW.Cancelled Or vnTransStatus > enuTCPDLW.On_Movement Then
                LblMsgError.Text = "Status Sudah Batal atau Done"
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

            vnQuery = "Select count(1) FROM Sys_SsoPDWScan1_TR Where PDWHOID=" & vnPDWHOID & " and abs(PDWScan1Deleted)=0"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnCount1 = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnCount1 > 0 Then
                LblMsgError.Text = "Sudah ada Barang di scan"
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

            pbuSsoProcessDataKey(HdfProcessDataKey.Value, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoPDWHeader_TR set TransStatus=" & enuTCPDLW.Cancelled & ",PDWCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnPDWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusPDW...Start")
            pbuInsertStatusPDW(vnPDWHOID, enuTCPDLW.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPDW...End")

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
End Class