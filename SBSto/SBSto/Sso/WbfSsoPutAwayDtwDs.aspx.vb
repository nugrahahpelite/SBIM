Imports System.Data.SqlClient
Imports System.Threading
Public Class WbfSsoPutAwayDtwDs
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoPutAwayDtwDs"
    Const csTNoPrefix = "DTW"

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
        RcvPONo = 0
        BRGCODE = 1
        BRGNAME = 2
        vSumDTWScan1Qty = 3
        vSumDTWScan2Qty = 4
        RcvPOHOID = 5
    End Enum
    Enum ensColData1
        OID = 0
        RcvPONo = 1
        DTWScan1Qty = 2
        vDTWScan1Note = 3
        vDTWScan1User = 4
        vDTWScan1Time = 5
        vDelItem1 = 6
        vDTWScan1Deleted = 7
        DTWScan1DeletedNote = 8
        vDTWScan1DeletedTime = 9
    End Enum

    Enum ensColData2
        OID = 0
        vStorageInfoHtml = 1
        RcvPONo = 2
        DTWScan2Qty = 3
        vDTWScan2Note = 4
        vDTWScan2User = 5
        vDTWScan2Time = 6
        vDelItem2 = 7
        vDTWScan2Deleted = 8
        DTWScan2DeletedNote = 9
        vDTWScan2DeletedTime = 10
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransRefNo.Text = ""
        TxtTransInvNo.Text = ""
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPutaway, vnSQLConn)

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

        If ChkSt_Baru.Checked = False And ChkSt_OnPutaway.Checked = False And ChkSt_PutawayDone.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_OnPutaway.Checked = True
            ChkSt_PutawayDone.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCPDTW.Baru & ","
        End If
        If ChkSt_OnPutaway.Checked = True Then
            vnCrStatus += enuTCPDTW.On_Putaway & ","
        End If
        If ChkSt_PutawayDone.Checked = True Then
            vnCrStatus += enuTCPDTW.Putaway_Done & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPDTW.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select pwh.OID,pwh.DTWCompanyCode,pwh.DTWNo,convert(varchar(11),pwh.DTWDate,106)vDTWDate,pck.PCKNo,pch.PCLRefHNo,"
        vnQuery += vbCrLf & "       str.WarehouseName,pwh.DTWCancelNote,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnPutawayDatetime,106)+' '+convert(varchar(5),pwh.OnPutawayDatetime,108)+' '+ OP.UserName vOnPutaway,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.PutawayDoneDatetime,106)+' '+convert(varchar(5),pwh.PutawayDoneDatetime,108)+' '+ PD.UserName vPutawayDone"
        vnQuery += vbCrLf & "  From Sys_SsoDTWHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoPCKHeader_TR pck with(nolock)on pck.OID=pwh.PCKHOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=pwh.StorageOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnPutawayUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.PutawayDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.DTWCompanyCode and uc.UserOID=" & Session("UserOID")
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
            vnQuery += vbCrLf & " and pwh.DTWNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.DTWDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.DTWDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "     Order by pwh.DTWNo"

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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPutaway
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

        vnQuery = "Select pwh.OID,pwh.DTWNo,pwh.DTWCompanyCode,pwh.WarehouseOID,pwh.StorageOID,convert(varchar(11),pwh.DTWDate,106)vDTWDate,"
        vnQuery += vbCrLf & "      whs.WarehouseName,pwh.DTWDoneNote,pwh.PCKHOID,pck.PCKNo,pch.PCLRefHNo,pwh.TransStatus,ST.TransStatusDescr"
        vnQuery += vbCrLf & " From Sys_SsoDTWHeader_TR pwh"
        vnQuery += vbCrLf & "      inner join Sys_SsoPCKHeader_TR pck on pck.OID=pwh.PCKHOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA whs on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=pwh.TransStatus and ST.TransCode='" & stuTransCode.SsoPutaway & "'"
        vnQuery += vbCrLf & "Where pwh.OID=" & TxtTransID.Text

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vDTWDate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("DTWNo")
            TxtTransWhsName.Text = vnDtb.Rows(0).Item("WarehouseName")

            TxtTransRefNo.Text = vnDtb.Rows(0).Item("PCKNo")
            TxtTransInvNo.Text = vnDtb.Rows(0).Item("PCLRefHNo")

            TxtCompany.Text = vnDtb.Rows(0).Item("DTWCompanyCode")

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
        BtnCancelDTW.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "DTWNo" Then
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

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriPWHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select ''RcvPONo,''BRGCODE,''BRGNAME,0 vSumDTWScan1Qty,0 vSumDTWScan2Qty,0 RcvPOHOID Where 1=2"
        Else
            vnQuery = "Select pws1.RcvPONo,mb.BRGCODE,mb.BRGNAME,pws1.vSumDTWScan1Qty,pws2.vSumDTWScan2Qty,pws1.RcvPOHOID"
            vnQuery += vbCrLf & "From fnTbl_SsoDTWHeaderScan1(" & vriPWHOID & ",0) pws1"
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws1.DTWCompanyCode"
            vnQuery += vbCrLf & "     left outer join fnTbl_SsoDTWHeaderScan2(" & vriPWHOID & ",0) pws2 on pws2.DTWHOID=pws1.DTWHOID and pws2.BRGCODE=pws1.BRGCODE and pws1.RcvPOHOID=pws2.RcvPOHOID"

            vnQuery += vbCrLf & " Where (mb.BRGCODE like '%" & vnCriteria & "%' or mb.BRGNAME like '%" & vnCriteria & "%')"
            vnQuery += vbCrLf & "order by case when isnull(vSumDTWScan2Qty,0)=0 then 5 else 1 end,mb.BRGCODE"
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

        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''RcvPONo,0 DTWScan1Qty,"
            vnQuery += vbCrLf & "       ''DTWScan1Note,"
            vnQuery += vbCrLf & "       ''vDTWScan1User,"
            vnQuery += vbCrLf & "	    ''vDTWScan1Time,"
            vnQuery += vbCrLf & "	    ''vDTWScan1Deleted,''DTWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    ''vDTWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem1 Where 1=2"
        Else
            vnQuery = "Select sc.OID,rch.RcvPONo,sc.DTWScan1Qty,"
            vnQuery += vbCrLf & "       sc.DTWScan1Note,"
            vnQuery += vbCrLf & "       mu.UserID vDTWScan1User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DTWScan1Datetime,106) + ' ' + convert(varchar(5),sc.DTWScan1Datetime,108)vDTWScan1Time,"
            vnQuery += vbCrLf & "	    case when abs(DTWScan1Deleted)=1 then 'Y' else 'N' end vDTWScan1Deleted,DTWScan1DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DTWScan1DeletedDatetime,106) + ' ' + convert(varchar(5),sc.DTWScan1DeletedDatetime,108)vDTWScan1DeletedTime,"
            vnQuery += vbCrLf & "	    case when abs(DTWScan1Deleted)=1 then '' else 'Hapus' end vDelItem1"
            vnQuery += vbCrLf & "  From Sys_SsoDTWScan1_TR sc with(nolock)"
            vnQuery += vbCrLf & "	    inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.DTWScan1UserOID"
            vnQuery += vbCrLf & " Where sc.DTWHOID=" & TxtTransID.Text & " and sc.RcvPOHOID=" & vriRcvPOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
            vnQuery += vbCrLf & "       and sc.DTWScan1Note like '%" & vnCriteria & "%'"

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(DTWScan1Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(DTWScan1Deleted)=1"
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
                If GrvData1.Rows(vn).Cells(ensColData1.vDTWScan1Deleted).Text = "Y" Then
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

        Dim vnCriteria As String = ""
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''RcvPONo,''vStorageInfoHtml,0 DTWScan2Qty,"
            vnQuery += vbCrLf & "       ''DTWScan2Note,"
            vnQuery += vbCrLf & "       ''vDTWScan2User,"
            vnQuery += vbCrLf & "	    ''vDTWScan2Time,"
            vnQuery += vbCrLf & "	    ''vDTWScan2Deleted,''DTWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    ''vDTWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem2 Where 1=2"
        Else
            vnQuery = "Select sc.OID,rch.RcvPONo,st.vStorageInfoHtml,sc.DTWScan2Qty,"
            vnQuery += vbCrLf & "       sc.DTWScan2Note,"
            vnQuery += vbCrLf & "       mu.UserID vDTWScan2User,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DTWScan2Datetime,106) + ' ' + convert(varchar(5),sc.DTWScan2Datetime,108)vDTWScan2Time,"
            vnQuery += vbCrLf & "	    case when abs(DTWScan2Deleted)=1 then 'Y' else 'N' end vDTWScan2Deleted,DTWScan2DeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.DTWScan2DeletedDatetime,106) + ' ' + convert(varchar(5),sc.DTWScan2DeletedDatetime,108)vDTWScan2DeletedTime,"
            vnQuery += vbCrLf & "	    case when abs(DTWScan2Deleted)=1 then '' else 'Hapus' end vDelItem2"
            vnQuery += vbCrLf & "  From Sys_SsoDTWScan2_TR sc with(nolock)"
            vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('') st on st.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sc.RcvPOHOID"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.DTWScan2UserOID"
            vnQuery += vbCrLf & " Where sc.DTWHOID=" & TxtTransID.Text & " and sc.RcvPOHOID=" & vriRcvPOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"

            If Val(vriStorageOID) > 0 Then
                vnQuery += vbCrLf & "       and sc.StorageOID=" & vriStorageOID
            End If

            vnQuery += vbCrLf & "       and sc.DTWScan2Note like '%" & vnCriteria & "%'"

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(DTWScan2Deleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(DTWScan2Deleted)=1"
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
                If GrvData2.Rows(vn).Cells(ensColData2.vDTWScan2Deleted).Text = "Y" Then
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

    Protected Sub BtnCancelDTW_Click(sender As Object, e As EventArgs) Handles BtnCancelDTW.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Putaway No. " & TxtTransID.Text & " ?<br />WARNING : Batal Putaway Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelDTW"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelDTW" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelDTW()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnCancelDTW.Enabled = (HdfTransStatus.Value = enuTCPWAY.Baru Or HdfTransStatus.Value = enuTCPWAY.On_Putaway)
            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnCancelDTW.Visible = BtnCancelDTW.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub

    Private Sub psCancelDTW()
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
            Dim vnDTWHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnCount1 As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoDTWHeader_TR with(nolock) Where OID=" & vnDTWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus = enuTCPDTW.Cancelled Or vnTransStatus > enuTCPDTW.On_Putaway Then
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

            vnQuery = "Select count(1) FROM Sys_SsoDTWScan1_TR Where DTWHOID=" & vnDTWHOID & " and abs(DTWScan1Deleted)=0"
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

            vnQuery = "Update Sys_SsoDTWHeader_TR set TransStatus=" & enuTCPDTW.Cancelled & ",DTWCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnDTWHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusDTW...Start")
            pbuInsertStatusDTW(vnDTWHOID, enuTCPDTW.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusDTW...End")

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