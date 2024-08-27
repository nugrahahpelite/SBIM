Imports System.Data.SqlClient
Public Class WbfSsoPutAwayDsw
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoPutAwayDsw"
    Const csTNoPrefix = "DSW"

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

    Enum ensColData1
        OID = 0
        RcvPONo = 1
        DSWScan1Qty = 2
        vDSWScan1Note = 3
        vDSWScan1User = 4
        vDSWScan1Time = 5
        vDelItem1 = 6
        vDSWScan1Deleted = 7
        DSWScan1DeletedNote = 8
        vDSWScan1DeletedTime = 9
    End Enum

    Enum ensColData2
        OID = 0
        vStorageInfoHtml = 1
        RcvPONo = 2
        PWScan2Qty = 3
        vPWScan2Note = 4
        vPWScan2User = 5
        vPWScan2Time = 6
        vDelItem2 = 7
        vDSWScan2Deleted = 8
        DSWScan2DeletedNote = 9
        vDSWScan2DeletedTime = 10
    End Enum

    Enum ensColSumm
        RcvPOHOID = 0
        RcvPONo = 1
        BRGCODE = 2
        BRGNAME = 3
        vSumDSWScan1Qty = 4

        vSumDSWScan2Qty = 5
        vConfirm = 6
        vNotConfirm = 7
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransRefNo.Text = ""

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPutaway_Antar_Wh, vnSQLConn)

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
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()

        If ChkSt_Baru.Checked = False And ChkSt_OnDelivery.Checked = False And ChkSt_OnPutaway.Checked = False And ChkSt_PutawayDone.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_OnDelivery.Checked = True
            ChkSt_OnPutaway.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCPYAY.Baru & ","
        End If
        If ChkSt_OnDelivery.Checked = True Then
            vnCrStatus += enuTCPYAY.On_Delivery_Putaway & ","
        End If
        If ChkSt_OnPutaway.Checked = True Then
            vnCrStatus += enuTCPYAY.On_Putaway & ","
        End If
        If ChkSt_PutawayDone.Checked = True Then
            vnCrStatus += enuTCPYAY.Putaway_Done & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPDTW.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	Select pwh.OID,pwh.DSWCompanyCode,pwh.DSWNo,convert(varchar(11),pwh.DSWDate,106)vDSWDate,dsr.DSRNo,	"
        vnQuery += vbCrLf & "	       str.WarehouseName,mdv.DcmDriverName,mvh.VehicleNo,pwh.DSWCancelNote,stn.TransStatusDescr	"
        vnQuery += vbCrLf & "	  From Sys_SsoDSWHeader_TR pwh with(nolock)	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoDSRHeader_TR dsr with(nolock) on dsr.OID=pwh.DSRHOID	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoDSPHeader_TR dsp with(nolock) on dsp.OID=dsr.DSPHOID	"
        vnQuery += vbCrLf & "	       inner join " & vnDBDcm & "Sys_DcmDriver_MA mdv with(nolock) on mdv.OID=dsp.DcmSchDriverOID	"
        vnQuery += vbCrLf & "	       inner join " & vnDBDcm & "Sys_DcmVehicle_MA mvh with(nolock) on mvh.OID=dsp.DcmVehicleOID	"
        vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=pwh.StorageOID	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.DSWCompanyCode	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus	"

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        vnQuery += vbCrLf & "and usc.UserOID=" & Session("UserOID")

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and pwh.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and ( pwh.DSWNo like '%" & Trim(TxtListNo.Text) & "%' OR dsr.DSWNo like '%" & Trim(TxtListNo.Text) & "%' )"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.DSWDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.DSWDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and ( STR.WarehouseOID = " & DstListWhs.SelectedValue & ")"
        End If

        vnQuery += vbCrLf & "     Order by pwh.DSWDate DESC"

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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPutaway_Antar_Wh
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

        vnQuery = "	SELECT TOP 1 pwh.OID,pwh.DSWNo,pwh.DSWCompanyCode,pwh.WarehouseOID,pwh.StorageOID,convert(varchar(11),pwh.DSWDate,106)vDSWDate,	"
        vnQuery += vbCrLf & "	      whs.WarehouseName, pwh.DSWDoneNote,pwh.DSRHOID,pck.PCKNo,pwh.TransStatus,stn.TransStatusDescr	"
        vnQuery += vbCrLf & "	 From Sys_SsoDSWHeader_TR pwh with(nolock)	"
        vnQuery += vbCrLf & "	      inner join Sys_SsoPCKHeader_TR pck with(nolock) on pck.OID=pwh.DSRHOID	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoDSRHeader_TR dsr with(nolock) on dsr.OID=pwh.DSRHOID	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoDSPHeader_TR dsp with(nolock) on dsp.OID=dsr.DSPHOID	"
        vnQuery += vbCrLf & "	      inner join " & vnDBMaster & "Sys_Warehouse_MA whs on whs.OID=pwh.WarehouseOID	"
        vnQuery += vbCrLf & "	  inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=pwh.StorageOID	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.DSWCompanyCode	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus	"
        vnQuery += vbCrLf & "	Where pwh.OID=" & TxtTransID.Text
        vnQuery += vbCrLf & "	ORDER BY pwh.DSWDate DESC	"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vDSWDate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("DSWNo")
            TxtTransRefNo.Text = vnDtb.Rows(0).Item("PCKNo")


            TxtTransWhsName.Text = vnDtb.Rows(0).Item("WarehouseName")

            TxtCompany.Text = vnDtb.Rows(0).Item("DSWCompanyCode")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        psFillGrvSumm(0, TxtTransID.Text, vriSQLConn)
        PanData.Visible = False
        vnDtb.Dispose()
    End Sub

    Private Sub psButtonStatusDefault()
        BtnList.Enabled = True
        BtnCancelDSW.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "DSWNo" Then
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

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriDSWHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 RcvPOHOID,''RcvPONo,''BRGCODE,''BRGNAME,0 vSumDSWScan1Qty,0 vSumDSWScan2Qty,''vConfirm,''vNotConfirm Where 1=2"
        Else
            vnQuery = "Select pws1.RcvPOHOID,pws1.RcvPONo,mb.BRGCODE,mb.BRGNAME,pws1.vSumDSWScan1Qty,pws2.vSumDSWScan2Qty"
            vnQuery += vbCrLf & "From fnTbl_SsoDSWHeaderScan1(" & vriDSWHOID & ",0) pws1"
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws1.DSWCompanyCode"
            vnQuery += vbCrLf & "     left outer join fnTbl_SsoDSWHeaderScan2(" & vriDSWHOID & ",0) pws2 on pws2.DSWHOID=pws1.DSWHOID and pws1.RcvPOHOID=pws2.RcvPOHOID and pws2.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws2.DSWCompanyCode"

            vnQuery += vbCrLf & "Where (mb.BRGCODE like '%" & vnCriteria & "%' or mb.BRGNAME like '%" & vnCriteria & "%')"
            vnQuery += vbCrLf & "order by case when isnull(vSumDSWScan2Qty,0)=0 then 5 else 1 end,mb.BRGCODE"
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

            LblDataTitle.Text = HdfStoKB.Value & " " & vnGRow.Cells(ensColSumm.RcvPONo).Text & " " & vnGRow.Cells(ensColSumm.BRGNAME).Text

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
            psFillGrvData2(0, 0, HdfRcvPOHOID.Value, HdfStoKB.Value, vnSQLConn)
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
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''RcvPONo,0 DSWScan1Qty,"
            vnQuery += vbCrLf & "       ''DSWScan1Note,"
            vnQuery += vbCrLf & "       ''vDSWScan1User,"
            vnQuery += vbCrLf & "	    ''vDSWScan1Time,"
            vnQuery += vbCrLf & "	    ''vDSWScan1Deleted,''DSWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    ''vDSWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem1 Where 1=2"
        Else
            vnQuery = "Select sc.OID,rch.RcvPONo,sc.DSWScan1Qty,"
            vnQuery += vbCrLf & "       sc.DSWScan1Note,"
            vnQuery += vbCrLf & "       mu.UserID vDSWScan1User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DSWScan1Datetime,106) + ' ' + convert(varchar(5),sc.DSWScan1Datetime,108)vDSWScan1Time,"
            vnQuery += vbCrLf & "	    case when abs(DSWScan1Deleted)=1 then 'Y' else 'N' end vDSWScan1Deleted,DSWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DSWScan1DeletedDatetime,106) + ' ' + convert(varchar(5),sc.DSWScan1DeletedDatetime,108)vDSWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    case when abs(DSWScan1Deleted)=1 then '' else 'Hapus' end vDelItem1"
            vnQuery += vbCrLf & "  From Sys_SsoDSWScan1_TR sc with(nolock)"
            vnQuery += vbCrLf & "	    inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.DSWScan1UserOID"
            vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA md with(nolock) on md.OID=sc.DSWScan1DeletedUserOID"
            vnQuery += vbCrLf & " Where sc.DSWHOID=" & TxtTransID.Text & " and sc.RcvPOHOID=" & vriRcvPOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(sc.DSWScan1Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(sc.DSWScan1Deleted)=1"
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
                If GrvData1.Rows(vn).Cells(ensColData1.vDSWScan1Deleted).Text = "Y" Then
                    GrvData1.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Private Sub psFillGrvData2(vriEmpty As Byte, vriStorageOID As String, vriRcvPOHOID As String, vriBrgCode As String, vriSQLConn As SqlClient.SqlConnection)
        If ChkSt_DelNo.Checked = False And ChkSt_DelYes.Checked = False Then
            ChkSt_DelNo.Checked = True
            ChkSt_DelYes.Checked = True
        End If

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''RcvPONo,''vStorageInfoHtml,0 DSWScan2Qty,"
            vnQuery += vbCrLf & "       ''vDSWScan2Note,"
            vnQuery += vbCrLf & "       ''vDSWScan2User,"
            vnQuery += vbCrLf & "	    ''vDSWScan2Time,"
            vnQuery += vbCrLf & "	    ''vDSWScan2Deleted,''DSWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    ''vDSWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem2 Where 1=2"
        Else
            vnQuery = "Select sc.OID,rch.RcvPONo,st.vStorageInfoHtml,sc.DSWScan2Qty,"
            vnQuery += vbCrLf & "       sc.DSWScan2Note,"
            vnQuery += vbCrLf & "       mu.UserID vDSWScan2User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DSWScan2Datetime,106) + ' ' + convert(varchar(5),sc.DSWScan2Datetime,108)vDSWScan2Time,"
            vnQuery += vbCrLf & "	    case when abs(DSWScan2Deleted)=1 then 'Y' else 'N' end vDSWScan2Deleted,DSWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DSWScan2DeletedDatetime,106) + ' ' + convert(varchar(5),sc.DSWScan2DeletedDatetime,108)vDSWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    case when abs(DSWScan2Deleted)=1 then '' else 'Hapus' end vDelItem2"
            vnQuery += vbCrLf & "  From Sys_SsoDSWScan2_TR sc with(nolock)"
            vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('') st on st.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.DSWScan2UserOID"
            vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA md with(nolock) on md.OID=sc.DSWScan2DeletedUserOID"
            vnQuery += vbCrLf & " Where sc.DSWHOID=" & TxtTransID.Text & " and sc.RcvPOHOID=" & vriRcvPOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"

            If Val(vriStorageOID) > 0 Then
                vnQuery += vbCrLf & "       and sc.StorageOID=" & vriStorageOID
            End If

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(DSWScan2Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(DSWScan2Deleted)=1"
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
                If GrvData2.Rows(vn).Cells(ensColData2.vDSWScan2Deleted).Text = "Y" Then
                    GrvData2.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If

        If Val(vriStorageOID) = 0 Then
            GrvData2.Columns(ensColData2.vStorageInfoHtml).HeaderStyle.CssClass = ""
            GrvData2.Columns(ensColData2.vStorageInfoHtml).ItemStyle.CssClass = ""
        Else
            GrvData2.Columns(ensColData2.vStorageInfoHtml).HeaderStyle.CssClass = "myDisplayNone"
            GrvData2.Columns(ensColData2.vStorageInfoHtml).ItemStyle.CssClass = "myDisplayNone"
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

    Protected Sub BtnCancelDSW_Click(sender As Object, e As EventArgs) Handles BtnCancelDSW.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Putaway No. " & TxtTransID.Text & " ?<br />WARNING : Batal Putaway Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelDSW"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelDSW" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelDSW()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnCancelDSW.Enabled = (HdfTransStatus.Value = enuTCPDSW.Baru Or HdfTransStatus.Value = enuTCPDSW.On_Putaway)
            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnCancelDSW.Visible = BtnCancelDSW.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub

    Private Sub psCancelDSW()
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
            Dim vnDSWHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnCount1 As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoDSWHeader_TR with(nolock) Where OID=" & vnDSWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus = enuTCPDSW.Cancelled Or vnTransStatus > enuTCPDSW.On_Putaway Then
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

            vnQuery = "Select count(1) FROM Sys_SsoDSWScan1_TR Where DSWHOID=" & vnDSWHOID & " and abs(DSWScan1Deleted)=0"
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

            vnQuery = "Update Sys_SsoDSWHeader_TR set TransStatus=" & enuTCPDSW.Cancelled & ",DSWCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnDSWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusDSW...Start")
            pbuInsertStatusDSW(vnDSWHOID, enuTCPDSW.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusDSW...End")

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