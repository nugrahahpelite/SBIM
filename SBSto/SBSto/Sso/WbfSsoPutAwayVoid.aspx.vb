Imports System.Data.SqlClient
Public Class WbfSsoPutAwayVoid
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoPutAwayVoid"
    Const csTNoPrefix = "PTV"


    Enum ensColList
        OID = 0
    End Enum
    Public Enum enuTCPDTV 'Putaway Void
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum
    Enum ensColData1
        OID = 0
        PYScan1Qty = 1
        vPYScan1Note = 2
        vPYScan1User = 3
        vPYScan1Time = 4
        vDelItem1 = 5
        vPYScan1Deleted = 6
        vPYScan1DeletedUser = 7
        PYScan1DeletedNote = 8
        vPYScan1DeletedTime = 9
    End Enum
    Public Enum enuTCPCKG 'Picking
        Void = -6
        Cancelled = -2
        None = 0
        On_Picking = 16
        Picking_Done = 18
        On_Dispatch_Putaway = 19
        On_Putaway_Void = 21
        Putaway_Dispatch_Done = 24
    End Enum
    Enum ensColData2
        OID = 0
        vStorageInfoHtml = 1
        PWScan2Qty = 2
        vPWScan2Note = 3
        vPWScan2User = 4
        vPWScan2Time = 5
        vDelItem2 = 6
        vPYScan2Deleted = 7
        vPYScan2DeletedUser = 8
        PYScan2DeletedNote = 9
        vPYScan2DeletedTime = 10
    End Enum

    Enum ensColSumm
        BRGCODE = 0
        BRGNAME = 1
        vSumPYScan1Qty = 2
        PYReceiveQty = 3
        vSumPYScan2Qty = 4
        vConfirm = 5
        vNotConfirm = 6
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        'TxtTransRefNo.Text = ""
        TxtTransNo.Text = ""
        TxtTransWhsName.Text = ""
        'TxtTransWhsNameDest.Text = ""
        TxtCompany.Text = ""

        HdfTransStatus.Value = enuTCSSOH.Baru
    End Sub
    Enum ensColLsScan
        vRcvPOScanDeleted = 5
    End Enum
    Enum ensColListItemNo
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
    End Enum
    Enum ensColLsPCK
        vPTVOID = 0
        PCKCompanyCode = 1
        vPCKOID = 2
        PCKNo = 3
        vPCKDate = 4
        PCLNo = 5
        PCLRefHNo = 6
        WarehouseName = 7
        TransStatusDescr = 8
    End Enum
    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"
        PanHdf.Style(HtmlTextWriterStyle.Visibility) = "hidden"

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
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")
        Dim vnDBMaster As String = fbuGetDBMaster()
        If ChkSt_OnDelivery.Checked = False And ChkSt_OnPutaway.Checked = False And ChkSt_PutawayDone.Checked = False Then
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
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable
        Dim vnCrList As String = fbuFormatString(Trim(TxtListFind.Text))

        vnQuery = "Select DISTINCT pwh.OID,pwh.PTVCompanyCode,pwh.PTVNo,convert(varchar(11),pwh.PTVDate,106)vPTVDate,pck.PCKNo,pch.PCLRefHNo,"
        vnQuery += vbCrLf & "       str.WarehouseName,stn.TransStatusDescr"
        vnQuery += vbCrLf & "  From Sys_SsoPTVHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoPCKHeader_TR pck with(nolock)on pck.OID=pwh.PCKHOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=pwh.StorageOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.PTVCompanyCode"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        vnQuery += vbCrLf & "and usc.UserOID=" & Session("UserOID")

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and pwh.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and ( pwh.PTVNo like '%" & Trim(TxtListNo.Text) & "%' OR pck.PCKNo like '%" & Trim(TxtListNo.Text) & "%' OR pch.PCLRefHNo like '%" & Trim(TxtListNo.Text) & "%' )"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PTVDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PTVDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and ( STR.WarehouseOID = " & DstListWhs.SelectedValue & " ) "
        End If
        'vnQuery += vbCrLf & " Where usc.UserOID=" & Session("UserOID") & " and str.WarehouseOID='" & Session("LoginWhsOID") & "'"
        'If vnCrList = "" Then
        '    vnQuery += vbCrLf & "           and pwh.PTVDate >= cast(dateadd(d,-17,getdate())as date)"
        'Else
        '    vnQuery += vbCrLf & "           and (pwh.PTVNo like '%" & vnCrList & "%' or rcv.RcvPONo like '%" & vnCrList & "%')"
        'End If
        'vnQuery += vbCrLf & "     Order by pwh.PTVNo desc"
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

    'Private Sub psShowPreview(vriBo As Boolean)
    '    If vriBo Then
    '        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
    '    Else
    '        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
    '    End If
    'End Sub
    'Private Sub psShowList(vriBo As Boolean)
    '    If vriBo Then
    '        DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
    '        tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
    '        psFillGrvList()
    '    Else
    '        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
    '        tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
    '    End If
    'End Sub

    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        psClearData()
        psClearMessage()
        psClearOption()

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
        psClearData()
        psClearMessage()
        psShowList(True)
        psFillGrvList()
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblXlsProses.Text = ""
    End Sub
    Private Sub psClearOption()
        TxtListTransID.Text = ""
        TxtListStart.Text = ""
        TxtListEnd.Text = ""

        TxtListNo.Text = ""
        TxtListTransID.Text = ""
    End Sub

    'Private Sub psShowPrOption(vriBo As Boolean)
    '    If vriBo Then
    '        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "visible"
    '    Else
    '        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
    '    End If
    'End Sub

    'Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
    '    If Not IsNumeric(TxtTransID.Text) Then Exit Sub
    '    If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

    '    Dim vnName1 As String = "Preview"
    '    Dim vnType As Type = Me.GetType()
    '    Dim vnClientScript As ClientScriptManager = Page.ClientScript
    '    If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
    '        Dim vnParam As String
    '        vnParam = "vqTrOID=" & TxtTransID.Text
    '        vnParam += "&vqTrCode=" & stuTransCode.SsoPutaway_Antar_Wh
    '        vnParam += "&vqTrNo=" & TxtTransNo.Text

    '        vbuPreviewOnClose = "0"

    '        'ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
    '        'psShowPreview(True)

    '        'Dim vnWinOpen As String
    '        'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
    '        'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
    '        'vnClientScript = Nothing
    '    End If
    'End Sub

    Private Sub psButtonStatusDefault()
        BtnList.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If e.CommandName = "PTVNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)
            TxtListTransID.Text = vnRow.Cells(ensColList.OID).Text
            HdfTransOID.Value = vnRow.Cells(ensColList.OID).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgConfirm.Text = pbMsgError
                LblMsgConfirm.Visible = True
                pbMsgError = ""
                Exit Sub
            End If

            psDisplayData(HdfTransOID.Value, vnSQLConn)


            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing


            tbTrans.Visible = True
            GrvSumm.Visible = True

            psShowList(False)

        End If
    End Sub


    'Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
    '    vbuPreviewOnClose = "1"
    '    psShowPreview(False)
    'End Sub

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

    'Protected Sub BtnProCancel_Click(sender As Object, e As EventArgs) Handles BtnProCancel.Click
    '    psShowPrOption(False)
    'End Sub

    'Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
    '    psClearMessage()
    '    If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
    'End Sub

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriPTVHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnDBMaster As String = fbuGetDBMaster()

        If vriEmpty = 1 Then
            vnQuery = "Select 0 vPTVOID,'' vPCKOID,'' PCKCompanyCode,''PCKNo,''vPCKDate,''PCLNo,''PCLRefHNo,''WarehouseName,''TransStatusDescr Where 1=2"
        Else
            vnQuery = "	Select DISTINCT pwh.OID vPTVOID ,pck.OID vPCKOID,pck.PCKCompanyCode,pck.PCKNo,convert(varchar(11),pck.PCKDate,106)vPCKDate,	"
            vnQuery += vbCrLf & "	       pch.PCLNo,pch.PCLRefHNo,	"
            vnQuery += vbCrLf & "	       whs.WarehouseName,stn.TransStatusDescr	"
            vnQuery += vbCrLf & "	  From Sys_SsoPCKHeader_TR pck with(nolock)	"
            vnQuery += vbCrLf & "	   inner join Sys_SsoPTVHeader_TR pwh with(nolock) on pck.OID=pwh.PCKHOID	"
            vnQuery += vbCrLf & "	   inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID	"
            vnQuery += vbCrLf & "	   inner join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID	"
            vnQuery += vbCrLf & "	   inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pck.PCKCompanyCode	"
            vnQuery += vbCrLf & "	   inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus	"
            vnQuery += vbCrLf & "	WHERE pck.PCKNo = '" & vriPTVHOID & "' AND pwh.OID = " & TxtTransID.Text & " 	"



        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
        GrvSumm.Visible = True
        tbTrans.Visible = True

    End Sub


    Protected Sub GrvSumm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSumm.SelectedIndexChanged

    End Sub

    Private Sub GrvSumm_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumm.RowCommand
        GrvData1.Visible = True
        PanData.Visible = True
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgConfirm.Text = pbMsgError
            LblMsgConfirm.Visible = True
            pbMsgError = ""
            Exit Sub
        End If

        If e.CommandName = "PCKNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvSumm.Rows(vnIdx)
            Dim vnGRow As GridViewRow = GrvSumm.Rows(vnIdx)
            'TxtFind.Text = vnRow.Cells(ensColList.OID).Text
            TxPCKOID3.Text = vnRow.Cells(ensColList.OID).Text
            HdfTransOID.Value = vnRow.Cells(ensColList.OID).Text
            'TxtStoPCKNo.Text = DirectCast(vnGRow.Cells(ensColLsPCK.PCKNo).Controls(0), LinkButton).Text
            TxtStoInvNo.Text = vnGRow.Cells(ensColLsPCK.PCLRefHNo).Text
            HdfPCKHOID.Value = vnGRow.Cells(ensColLsPCK.vPCKOID).Text
            HdfCompanyCode.Value = vnGRow.Cells(ensColLsPCK.PCKCompanyCode).Text





            TxtStoPCKNo.Text = DirectCast(vnGRow.Cells(ensColLsPCK.PCKNo).Controls(0), LinkButton).Text
            TxtStoInvNo.Text = vnGRow.Cells(ensColLsPCK.PCLRefHNo).Text
            HdfPCKHOID.Value = vnGRow.Cells(ensColLsPCK.vPCKOID).Text
            HdfCompanyCode.Value = vnGRow.Cells(ensColLsPCK.PCKCompanyCode).Text

            psFillGrvData1(0, TxtListTransID.Text, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
            GrvData1.Visible = True
            PanData.Visible = True

            tbTrans.Visible = True
            GrvSumm.Visible = True

            psShowList(False)

        End If
    End Sub

    Private Sub psDisplayData(vriOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()

        GrvSumm.Visible = True
        tbTrans.Visible = True
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select pwh.OID,pwh.PTVNo,pwh.PTVDate,pwh.PTVCompanyCode,pwh.WarehouseOID,pwh.StorageOID,str.WarehouseName, dc.CompanyName,"
        vnQuery += vbCrLf & "      pwh.PCKHOID,pck.PCKNo,pch.PCLRefHNo,pwh.TransStatus"
        vnQuery += vbCrLf & " From Sys_SsoPTVHeader_TR pwh"
        vnQuery += vbCrLf & "      inner join Sys_SsoPCKHeader_TR pck on pck.OID=pwh.PCKHOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.PTVCompanyCode"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=pwh.StorageOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "DimCompany dc with(nolock) on usc.CompanyCode=dc.CompanyCode"
        vnQuery += vbCrLf & "Where pwh.OID=" & vriOID
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""

            HdfDefGRHNo.Value = ""
            HdfDefGRHOID.Value = "0"

            PanData.Visible = False
        Else
            tbTrans.Visible = True
            GrvSumm.Visible = True
            tbSumm.Visible = True
            Dim vnDRow As DataRow = vnDtb.Rows(0)
            Dim vnDtb2 As New DataTable
            Dim vnQuery2 As String
            TxtTransOID.Text = vriOID
            TxtTransID.Text = vriOID
            TxtListFind.Text = vriOID

            TxtTransOID.Text = vriOID

            TxtListNo.Text = vnDtb.Rows(0).Item("PTVNo")
            TxtTransNo.Text = vnDtb.Rows(0).Item("PTVNo")
            TxtStoPCKNo.Text = vnDtb.Rows(0).Item("PCKNo")
            TxtTransDate.Text = vnDtb.Rows(0).Item("PTVDate")
            TxtCompany.Text = vnDtb.Rows(0).Item("CompanyName")
            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatus")
            TxtStoInvNo.Text = vnDtb.Rows(0).Item("PCLRefHNo")
            TxtTransWhsName.Text = vnDtb.Rows(0).Item("WarehouseName")
            HdfTransOID.Value = vriOID
            HdfWarehouseOID.Value = vnDtb.Rows(0).Item("WarehouseOID")


            HdfStorageOID.Value = vnDtb.Rows(0).Item("StorageOID")
            HdfCompanyCode.Value = vnDtb.Rows(0).Item("PTVCompanyCode")
            HdfPCKHOID.Value = vnDtb.Rows(0).Item("PCKHOID")

            vnQuery2 = "	Select DISTINCT pwh.OID vPTVOID ,pck.OID vPCKOID,pck.PCKCompanyCode,pck.PCKNo,convert(varchar(11),pck.PCKDate,106)vPCKDate,	"
            vnQuery2 += vbCrLf & "	       pch.PCLNo,pch.PCLRefHNo,	"
            vnQuery2 += vbCrLf & "	       whs.WarehouseName,stn.TransStatusDescr	"
            vnQuery2 += vbCrLf & "	  From Sys_SsoPCKHeader_TR pck with(nolock)	"
            vnQuery2 += vbCrLf & "	       inner join Sys_SsoPTVHeader_TR pwh with(nolock) on pck.OID=pwh.PCKHOID	"
            vnQuery2 += vbCrLf & "	       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pck.PCLHOID	"
            vnQuery2 += vbCrLf & "	       inner join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID	"
            vnQuery2 += vbCrLf & "	       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pck.PCKCompanyCode	"
            vnQuery2 += vbCrLf & "	       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus	"
            vnQuery2 += vbCrLf & "	 WHERE pck.PCKNo = '" & TxtStoPCKNo.Text & "' AND pwh.OID = " & TxtTransID.Text & " 	"

            pbuFillDtbSQL(vnDtb2, vnQuery2, vriSQLConn)
            GrvSumm.DataSource = vnDtb2
            GrvSumm.DataBind()
            GrvSumm.Visible = True
            tbTrans.Visible = True

            PanData.Visible = False
        End If
    End Sub

    Private Sub psShowList(vriBo As Boolean)
        If vriBo Then
            DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psFillGrvList()
            TxtListFind.Focus()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
        End If
    End Sub

    Private Sub psFillGrvLsPCK(vriEmpty As Byte)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblListError.Text = pbMsgError
            LblListError.Visible = True
            pbMsgError = ""


            Exit Sub
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''PCKCompanyCode,''PCKNo,''vPCKDate,''PCLNo,''PCLRefHNo,''WarehouseName,''TransStatusDescr Where 1=2"
        Else
            If Val(HdfStorageOID.Value) = 0 Then
                vnQuery = "Select 0 OID,''PCKCompanyCode,''PCKNo,''vPCKDate,''PCLNo,''PCLRefHNo,''WarehouseName,''TransStatusDescr Where 1=2"
            Else
                vnQuery = "Select pwh.OID,pwh.PCKCompanyCode,pwh.PCKNo,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
                vnQuery += vbCrLf & "       pch.PCLNo,pch.PCLRefHNo,"
                vnQuery += vbCrLf & "       whs.WarehouseName,stn.TransStatusDescr"
                vnQuery += vbCrLf & "  From Sys_SsoPCKHeader_TR pwh with(nolock)"
                vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
                vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID"
                vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.PCKCompanyCode"
                vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
                vnQuery += vbCrLf & " Where pwh.StorageOID=" & HdfStorageOID.Value & " and usc.UserOID=" & Session("UserOID") & " and pwh.WarehouseOID='" & Session("LoginWhsOID") & "' and"
                vnQuery += vbCrLf & "       pwh.TransStatus in(" & enuTCPCKG.Void & "," & enuTCPCKG.On_Putaway_Void & ")"
                vnQuery += vbCrLf & " Order by pwh.PCKNo desc"
            End If
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        GrvSumm.Visible = True
        tbTrans.Visible = True

    End Sub

    Private Sub psShowBrg(vriBo As Boolean)
        If vriBo Then
            GrvData1.Style(HtmlTextWriterStyle.Visibility) = "visible"
            PanData.Style(HtmlTextWriterStyle.Visibility) = "visible"

            PanData.Visible = True
            GrvData1.Visible = True


        Else
            GrvData1.Style(HtmlTextWriterStyle.Visibility) = "hidden"

            PanData.Style(HtmlTextWriterStyle.Visibility) = "hidden"


            PanData.Visible = False
            GrvData1.Visible = False


        End If
    End Sub

    Private Sub psShowSumm(vriBo As Boolean)
        If vriBo Then
            GrvSumm.Style(HtmlTextWriterStyle.Visibility) = "visible"

            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"

        Else
            GrvSumm.Style(HtmlTextWriterStyle.Visibility) = "hidden"

            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psShowBrg(False)
        End If
    End Sub
    Private Sub psFillGrvData1(vriEmpty As Byte, vriPTVHOID As String, vriSQLConn As SqlClient.SqlConnection)

        GrvData1.Visible = True
        PanData.Visible = True

        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select ''RcvPONo,''BRGCODE,''BRGNAME,0 vSumPTVScan1Qty,0 vSumPTVScan2Qty Where 1=2"
        Else
            vnQuery = "Select pws1.RcvPONo,mb.BRGCODE,mb.BRGNAME,pws1.vSumPTVScan1Qty,pws2.vSumPTVScan2Qty"
            vnQuery += vbCrLf & "From fnTbl_SsoPTVHeaderScan1(" & TxtListTransID.Text & ",0) pws1"
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws1.PTVCompanyCode"
            vnQuery += vbCrLf & "     left outer join fnTbl_SsoPTVHeaderScan2(" & TxtListTransID.Text & ",0) pws2 on pws2.PTVHOID=pws1.PTVHOID and pws1.RcvPOHOID=pws2.RcvPOHOID and pws2.BRGCODE=pws1.BRGCODE and mb.CompanyCode=pws2.PTVCompanyCode"

            vnQuery += vbCrLf & " Where (mb.BRGCODE like '%" & vnCriteria & "%' or mb.BRGNAME like '%" & vnCriteria & "%')"

            If RdbSumm.SelectedValue = 2 Then
                vnQuery += vbCrLf & "       and pws1.vSumPTVScan1Qty<>isnull(pws2.vSumPTVScan2Qty,0)"
            End If
            vnQuery += vbCrLf & "order by case when isnull(vSumPTVScan2Qty,0)=0 then 5 else 1 end,mb.BRGCODE"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData1.DataSource = vnDtb
        GrvData1.DataBind()
        GrvData1.Visible = True
        PanData.Visible = True
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub
End Class