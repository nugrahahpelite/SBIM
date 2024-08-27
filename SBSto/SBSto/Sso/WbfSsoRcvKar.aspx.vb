Imports System.Data.SqlClient
Public Class WbfSsoRcvKar
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoRcvKar"
    Const csTNoPrefix = "KR"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
    End Enum

    Enum ensColListBrg
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
        vQtyKrOutstanding = 3
        STKRHOID = 4
    End Enum
    Enum ensColDetail
        OID = 0
        vAddItem = 1
        BRGCODE = 2
        BRGNAME = 3
        RcvKRQty = 4
        TxtRcvKRQty = 5
        vDelItem = 6
        STKRHOID = 7
        vMessage = 8
    End Enum

    Enum ensColLsRcvPO
        RcvPONo = 0
        vRcvPODate = 1
        RcvPOSupplierName = 2
        RcvPOTypeName = 3
        OID = 4
        RcvPORefTypeOID = 5
        RcvRefTypeOID = 6
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPenerimaanDate.Text = ""
        TxtPenerimaanNo.Text = ""
        TxtRcvPONo.Text = ""

        TxtPenerimaanNote.Text = ""
        HdfTransStatus.Value = enuTCRCKR.Baru
        HdfWarehouseOID.Value = 0
        HdfRcvPOHOID.Value = "0"
        HdfRcvKRTypeOID.Value = "0"
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsRcvPO.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "DMgm"
        If Not IsPostBack Then
            psDefaultDisplay()
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPenerimaanKarantina, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstWhs, False, vnSQLConn)

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, True, vnSQLConn)
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
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_OnReceive.Checked = False And ChkSt_Prepared.Checked = False And ChkSt_Approved.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_Prepared.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCRCKR.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCRCKR.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCRCKR.Prepared & ","
        End If
        If ChkSt_Approved.Checked = True Then
            vnCrStatus += enuTCRCKR.Approved & ","
        End If
        If ChkSt_OnReceive.Checked = True Then
            vnCrStatus += enuTCRCKR.On_Receive & ","
        End If
        If ChkSt_ReceiveDone.Checked = True Then
            vnCrStatus += enuTCRCKR.Receive_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	SELECT PM.OID,PM.RcvKRNo,convert(varchar(11),PM.RcvKRDate,106)vRcvKRDate,SM.RcvPONo,SM.RcvPOSupplierName,KR.RcvKRTypeName,"
        vnQuery += vbCrLf & "	  PM.WarehouseOID,PW.WarehouseName,PM.RcvKRCompanyCode,"
        vnQuery += vbCrLf & "	  PM.RcvKRNote,ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ApprovedDatetime,106)+' '+convert(varchar(5),PM.ApprovedDatetime,108)+' '+ AP.UserName vApproved"
        vnQuery += vbCrLf & "FROM Sys_SsoRcvKRHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "	  inner join Sys_SsoRcvPOHeader_TR SM with(nolock) on SM.OID=PM.RcvPOHOID"
        vnQuery += vbCrLf & "	  inner join Sys_SsoRcvKRType_MA KR with(nolock) on KR.OID=PM.RcvKRTypeOID"
        vnQuery += vbCrLf & "	  inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA PW with(nolock) on PW.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "	  inner join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "	  left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"
        vnQuery += vbCrLf & "	  left outer join Sys_SsoUser_MA AP with(nolock) on AP.OID=PM.ApprovedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.RcvKRCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.RcvKRNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvKRDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvKRDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWhs.SelectedValue
        End If
        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.RcvKRCompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.RcvKRNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnOID As String = "0"
        Dim vnvAddItem As String = "..."
        Dim vnBRGCODE As String = ""
        Dim vnBRGNAME As String = ""

        Dim vnvDelItem As String = ""
        Dim vnvMessage As String = ""

        Dim vnSTKRHOID As Integer = 0
        Dim vnRcvKRQty As Integer = 0

        If vriHOID = "0" Then
            vnQuery = "	 Select '' OID,'' vAddItem,"
            vnQuery += vbCrLf & "	     '' BRGCODE,'' BRGNAME,0 RcvKRQty,"
            vnQuery += vbCrLf & "	    '' vDelItem,0 STKRHOID,''vMessage"
            vnQuery += vbCrLf & "	  From Sys_SsoRcvKRDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        Else
            vnQuery = "	 Select pld.OID,'' vAddItem,"
            vnQuery += vbCrLf & "	     pld.BRGCODE,pld.BRGNAME,pld.RcvKRQty,"
            vnQuery += vbCrLf & "	    'Hapus Item'vDelItem, pld.STKRHOID,''vMessage"
            vnQuery += vbCrLf & "	  From Sys_SsoRcvKRDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & " Where pld.RcvKRHOID =" & vriHOID & " "
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        End If

        Dim vn As Integer
        If HdfActionStatus.Value = cbuActionNorm Then
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

            If HdfTransStatus.Value = enuTCPLSP.Baru Then
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = ""
            Else
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"
            End If
            GrvDetail.Columns(ensColDetail.RcvKRQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.RcvKRQty).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtRcvKRQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtRcvKRQty).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.RcvKRQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.RcvKRQty).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtRcvKRQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtRcvKRQty).ItemStyle.CssClass = ""
            For vn = 0 To 40
                vnDtb.Rows.Add(New Object() {vnOID, vnvAddItem, vnBRGCODE, vnBRGNAME, vnRcvKRQty, vnvDelItem, vnSTKRHOID, vnvMessage})
            Next
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()

        Dim vnGRow As GridViewRow
        Dim vnTxtRcvKRQty As TextBox

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtRcvKRQty = vnGRow.FindControl("TxtRcvKRQty")
            vnTxtRcvKRQty.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.RcvKRQty).Text)
        Next
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
            psFillGrvList()
            psButtonShowList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psButtonStatus()
        End If
    End Sub

    Private Sub psShowLsRcvPO(vriBo As Boolean)
        If vriBo Then
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psSetTransNo(vriCompanyCode As String, vriWarehouseCode As String, vriSQLConn As SqlConnection)
        Dim vnQuery As String
        vnQuery = "Select '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(RcvKRNo,len(RcvKRNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(RcvKRNo,len(RcvKRNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoRcvKRHeader_TR with(nolock)"
        vnQuery += vbCrLf & "       Where substring(RcvKRNo,1,len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'))="
        vnQuery += vbCrLf & "                                     '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "                                 and len(RcvKRNo)=len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/')+4"
        TxtPenerimaanNo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnCancelPCL.Enabled = False
        BtnPrepare.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True
    End Sub
    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psShowList(True)
        psButtonVisible()
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        DivBrg.Visible = True
        GrvDetail.Visible = True

        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
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

        Session(csModuleName & stuSession.Simpan) = ""

        psClearData()

        If DstCompany.Items.Count > 0 Then
            DstCompany.SelectedIndex = 0
        End If

        TxtPenerimaanDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)


    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgType.Text = ""
        LblMsgWhs.Text = ""

        LblMsgPenerimaanDate.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPenerimaanNo.ReadOnly = Not vriBo
        TxtPenerimaanNote.ReadOnly = Not vriBo

        RdbTypeRelease.Enabled = vriBo
        RdbTypeMinus.Enabled = vriBo
        RdbTypePlus.Enabled = vriBo

        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelPCL.Visible = Not vriBo
        BtnPrepare.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo


        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        DivBrg.Visible = True
        GrvDetail.Visible = True
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)

        If e.CommandName = "vAddItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True

                Exit Sub
            End If
            HdfDetailRowIdx.Value = vnIdx

            psShowListBrg(True)

        ElseIf e.CommandName = "vDelItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If
            HdfDetailRowIdx.Value = vnIdx
            LblConfirmMessage.Text = "Anda Hapus Item " & vnGRow.Cells(ensColDetail.BRGNAME).Text & " ?"
            HdfProcess.Value = "vDelItem"
            tbConfirmNote.Visible = False
            psShowConfirm(True)

        End If
    End Sub

    Private Sub psShowConfirm(vriBo As Boolean)
        If vriBo Then
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "visible"
            BtnConfirmYes.Visible = True
            BtnConfirmNo.Text = "NO"
            DivBrg.Visible = True
            GrvDetail.Visible = True
        Else
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            DivBrg.Visible = True
            GrvDetail.Visible = True
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPenerimaanKarantina
            vnParam += "&vqTrNo=" & TxtPenerimaanNo.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)
        End If
    End Sub

    Private Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
        psEnableSave(False)
        psButtonVisible()

        HdfActionStatus.Value = cbuActionNorm
        If TxtTransID.Text = "" Then
            psClearData()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvDetail(0, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        Else
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
        End If
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        DivBrg.Visible = True
        GrvDetail.Visible = True
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,convert(varchar(11),PM.RcvKRDate,106)vRcvKRDate,SM.RcvPONo,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoRcvKRHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "	  inner join Sys_SsoRcvPOHeader_TR SM with(nolock) on SM.OID=PM.RcvPOHOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoPenerimaanKarantina & "'"

        vnQuery += vbCrLf & "     Where PM.OID =" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtPenerimaanDate.Text = vnDtb.Rows(0).Item("vRcvKRDate")
            TxtPenerimaanNo.Text = vnDtb.Rows(0).Item("RcvKRNo")
            TxtPenerimaanNote.Text = vnDtb.Rows(0).Item("RcvKRNote")

            TxtRcvPONo.Text = vnDtb.Rows(0).Item("RcvPONo")
            HdfRcvPOHOID.Value = vnDtb.Rows(0).Item("RcvPOHOID")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfRcvKRTypeOID.Value = vnDtb.Rows(0).Item("RcvKRTypeOID")

            If vnDtb.Rows(0).Item("RcvKRTypeOID") = enuRcvKRType.Release_Minus Then
                RdbTypeRelease.Checked = True
                RdbTypeMinus.Checked = False
                RdbTypePlus.Checked = False
            ElseIf vnDtb.Rows(0).Item("RcvKRTypeOID") = enuRcvKRType.Receive_Minus Then
                RdbTypeRelease.Checked = False
                RdbTypeMinus.Checked = True
                RdbTypePlus.Checked = False
            ElseIf vnDtb.Rows(0).Item("RcvKRTypeOID") = enuRcvKRType.Receive_Plus Then
                RdbTypeRelease.Checked = False
                RdbTypeMinus.Checked = False
                RdbTypePlus.Checked = True
            End If

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("RcvKRCompanyCode"))
            DstWhs.SelectedValue = Trim(vnDtb.Rows(0).Item("WarehouseOID"))
            HdfWarehouseOID.Value = Trim(vnDtb.Rows(0).Item("WarehouseOID"))
            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")
            psButtonStatus()
            psButtonVisible()
            DivBrg.Visible = True
            GrvDetail.Visible = True
            BtnEdit.Visible = True
        End If

        psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelPCL.Visible = BtnCancelPCL.Enabled
        BtnPrepare.Visible = BtnPrepare.Enabled
        BtnApprove.Visible = BtnApprove.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelPCL.Enabled = False
        BtnPrepare.Enabled = False
        BtnApprove.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCRCKR.Baru)

            BtnCancelPCL.Enabled = (HdfTransStatus.Value = enuTCRCKR.Baru Or HdfTransStatus.Value = enuTCRCKR.Prepared)

            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCRCKR.Baru)
            BtnApprove.Enabled = (HdfTransStatus.Value = enuTCRCKR.Prepared)

            BtnPreview.Enabled = (HdfTransStatus.Value = enuTCRCKR.Prepared)
            BtnPreview.Enabled = False

            psButtonVisible()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Trim(TxtTransID.Text) = "" Then Exit Sub
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
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

        Session(csModuleName & stuSession.Simpan) = ""

        HdfActionStatus.Value = cbuActionEdit
        psFillGrvDetail(TxtTransID.Text, vnSQLConn)


        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)


        psEnableSave(True)


    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        psSaveBaru()
        'If HdfActionStatus.Value = cbuActionNew Then
        '    psSaveBaru()
        'Else
        '    psSaveEdit()
        'End If
    End Sub
    Private Sub psSaveBaru()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If

        Dim vnSave As Boolean = True
        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If Val(DstWhs.SelectedValue) = 0 Then
            LblMsgWhs.Text = "Pilih Warehouse"
            vnSave = False
        End If
        If Not IsDate(Trim(TxtPenerimaanDate.Text)) Then
            LblMsgPenerimaanDate.Text = "Isi Tanggal"
            vnSave = False
        End If
        If RdbTypeRelease.Checked = False And RdbTypeMinus.Checked = False And RdbTypePlus.Checked = False Then
            LblMsgType.Text = "Pilih Release Minus, Receive Karantina Plus atau Minus"
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

            Dim vnQuery As String
            Dim vnGRow As GridViewRow

            Dim vnBrgCode As String
            Dim vnBrgCode_List As String = ""

            Dim vnCriteria As String
            If RdbTypeRelease.Checked Or RdbTypeMinus.Checked Then
                vnCriteria = "and QtyKarantina<0"
            Else
                vnCriteria = "and QtyKarantina>0"
            End If

            Dim vnBrgCek As Boolean = True
            Dim vnQtyTrans As Integer
            Dim vnTxtRcvKRQty As TextBox
            For vn = 0 To GrvDetail.Rows.Count - 1
                vnGRow = GrvDetail.Rows(vn)
                vnGRow.Cells(ensColDetail.vMessage).Text = ""
                vnBrgCode = UCase(fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text))
                If vnBrgCode <> "" Then
                    If vnBrgCode_List Like "*" & vnBrgCode & "*" Then
                        vnGRow.Cells(ensColDetail.vMessage).Text = vnBrgCode & " Sudah Ada"
                        vnBrgCek = False
                    Else
                        vnTxtRcvKRQty = vnGRow.FindControl("TxtRcvKRQty")
                        vnQtyTrans = fbuValNumHtml(vnTxtRcvKRQty.Text)
                        If Val(vnQtyTrans) > 0 Then
                            vnQuery = "Select count(1) From Sys_SsoStockKarantina_TR Where BrgCode='" & vnBrgCode & "'" & vnCriteria
                            If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                                vnGRow.Cells(ensColDetail.vMessage).Text = "Item Qty Karantina Error"
                                vnBrgCek = False
                            End If
                        End If
                    End If
                    vnBrgCode_List += "-" & vnBrgCode
                End If
            Next
            If vnBrgCek = False Then
                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Exit Sub
            End If

            '27 Sep 2023 sampai sini mestinya cek Qty Release < Qty Karantina
            If RdbTypeRelease.Checked Then
                Dim vnQtySisa As Integer
                For vn = 0 To GrvDetail.Rows.Count - 1
                    vnGRow = GrvDetail.Rows(vn)
                    vnGRow.Cells(ensColDetail.vMessage).Text = ""
                    vnBrgCode = UCase(fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text))
                    If vnBrgCode <> "" Then
                        vnTxtRcvKRQty = vnGRow.FindControl("TxtRcvKRQty")
                        vnQtyTrans = fbuValNumHtml(vnTxtRcvKRQty.Text)
                        If Val(vnQtyTrans) > 0 Then
                            vnQuery = "Select abs(QtyKarantina)-abs(QtyKrRelease)-abs(QtyKrReceive) From Sys_SsoStockKarantina_TR Where BrgCode='" & vnBrgCode & "'" & vnCriteria
                            vnQtySisa = fbuGetDataNumSQL(vnQuery, vnSQLConn)
                            If vnQtySisa - vnQtyTrans < 0 Then
                                vnGRow.Cells(ensColDetail.vMessage).Text = "Qty Error...Qty Sisa = " & vnQtySisa
                                vnBrgCek = False
                            End If
                        End If
                    End If
                Next
                If vnBrgCek = False Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
            End If

            Dim vnHOID As String
            Dim vnUserNIP As String = Session("UserName")

            Dim vnRcvKRNo As String

            If HdfActionStatus.Value = cbuActionNew Then
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)
                Dim vnWarehouseOID As Integer = DstWhs.SelectedValue
                Dim warehousecode As String = fbuGetWhsCode_ByOID(vnWarehouseOID, vnSQLConn)
                psSetTransNo(vnCompanyCode, warehousecode, vnSQLConn)
                vnRcvKRNo = Trim(TxtPenerimaanNo.Text)

                vnQuery = "Select max(OID) from Sys_SsoRcvKRHeader_TR with(nolock)"
                vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoRcvKRHeader_TR(OID,RcvKRNo,RcvKRDate,"
                vnQuery += vbCrLf & "RcvKRCompanyCode,"
                vnQuery += vbCrLf & "RcvPOHOID,"
                vnQuery += vbCrLf & "RcvKRTypeOID,"
                vnQuery += vbCrLf & "WarehouseOID,"
                vnQuery += vbCrLf & "RcvKRNote,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vnRcvKRNo & "','" & TxtPenerimaanDate.Text & "',"

                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "',"
                vnQuery += vbCrLf & HdfRcvPOHOID.Value & ","

                If RdbTypeRelease.Checked Then
                    vnQuery += vbCrLf & enuRcvKRType.Release_Minus & ","
                ElseIf RdbTypeMinus.Checked Then
                    vnQuery += vbCrLf & enuRcvKRType.Receive_Minus & ","
                ElseIf RdbTypePlus.Checked Then
                    vnQuery += vbCrLf & enuRcvKRType.Receive_Plus & ","
                End If

                vnQuery += vbCrLf & "'" & Trim(vnWarehouseOID) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoPenerimaanKarantina & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnHOID, vnSQLConn, vnSQLTrans)

                If RdbTypeRelease.Checked Then
                    psUpdate_StockKarantina_QtyRelease(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
                Else
                    psUpdate_StockKarantina_QtyReceive(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
                End If

                pbuInsertStatusRcvKR(vnHOID, enuTCRCKR.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnHOID

                HdfTransStatus.Value = enuTCRCKR.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

            Else
                vnHOID = TxtTransID.Text

                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoRcvKRHeader_TR set"
                vnQuery += vbCrLf & "RcvKRNote='" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

                If RdbTypeRelease.Checked Then
                    psUpdate_StockKarantina_QtyRelease(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
                Else
                    psUpdate_StockKarantina_QtyReceive(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
                End If

                pbuInsertStatusRcvKR(vnHOID, enuTCRCKR.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"
            End If

            psEnableInput(False)
            psEnableSave(False)

            HdfActionStatus.Value = cbuActionNorm
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
    Private Sub psSaveEdit()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If
        Dim vnSave As Boolean = True
        psClearMessage()

        If Not vnSave Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgXlsProsesError.Text = pbMsgError
                LblMsgXlsProsesError.Visible = True
                Exit Sub
            End If

            Dim vnHOID As String = TxtTransID.Text
            Dim vnQuery As String
            Dim vnUserNIP As String = Session("EmpNIP")

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            If (HdfTransStatus.Value = enuTCSSOH.Baru) Then
                vnQuery = "Update Sys_SsoRcvKRHeader_TR set "
                vnQuery += vbCrLf & "RcvKRDate='" & TxtPenerimaanDate.Text & "',"
                vnQuery += vbCrLf & "RcvKRNote='" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & vnHOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

            If RdbTypeRelease.Checked Then
                psUpdate_StockKarantina_QtyRelease(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
            Else
                psUpdate_StockKarantina_QtyReceive(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
            End If

            pbuInsertStatusSSOH(vnHOID, enuTCSSOH.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            psEnableInput(False)
            psEnableSave(False)

            HdfActionStatus.Value = cbuActionNorm
            GrvDetail.PagerSettings.Visible = True
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
    Private Sub psSaveDetail(vriRcvKRHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow

        Dim vnTxtRcvKRQty As TextBox
        Dim vnBrgCode As String

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtRcvKRQty = vnGRow.FindControl("TxtRcvKRQty")

            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                    If RdbTypePlus.Checked Then
                        vnTxtRcvKRQty.Text = vnGRow.Cells(ensColDetail.RcvKRQty).Text
                    End If

                    vnQuery = "Insert into Sys_SsoRcvKRDetail_TR"
                    vnQuery += vbCrLf & "(RcvKRHOID,STKRHOID,"
                    vnQuery += vbCrLf & "BRGCODE,"
                    vnQuery += vbCrLf & "BRGNAME,"
                    vnQuery += vbCrLf & "RcvKRQty"
                    vnQuery += vbCrLf & ")"
                    vnQuery += vbCrLf & "values(" & vriRcvKRHOID & "," & vnGRow.Cells(ensColDetail.STKRHOID).Text & ","
                    vnQuery += vbCrLf & "'" & vnBrgCode & "',"
                    vnQuery += vbCrLf & "'" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGNAME).Text) & "',"
                    vnQuery += vbCrLf & Val(Replace(vnTxtRcvKRQty.Text, "", "")) & ")"

                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                End If
            Else
                vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                If RdbTypePlus.Checked Then
                    vnTxtRcvKRQty.Text = vnGRow.Cells(ensColDetail.RcvKRQty).Text
                End If

                vnQuery = "Update Sys_SsoRcvKRDetail_TR SET"
                vnQuery += vbCrLf & "RcvKRQty=" & Val(Replace(vnTxtRcvKRQty.Text, "", "")) & ""

                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            End If
        Next
    End Sub

    Private Sub psUpdate_StockKarantina_QtyRelease(vriRcvKRHOID As String, vriRcvPOHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDtbDetail As New DataTable
        Dim vnQuery As String
        vnQuery = "Select BrgCode From Sys_SsoRcvKRDetail_TR Where RcvKRHOID=" & vriRcvKRHOID
        pbuFillDtbSQLTrans(vnDtbDetail, vnQuery, vriSQLConn, vriSQLTrans)

        Dim vn As Integer
        Dim vnBrgCode As String
        Dim vnQtyRcvKR As Integer
        For vn = 0 To vnDtbDetail.Rows.Count - 1
            vnBrgCode = vnDtbDetail.Rows(vn).Item("BrgCode")
            vnQuery = "Select sum(RcvKRQty)"
            vnQuery += vbCrLf & "       From Sys_SsoRcvKRDetail_TR krd with(nolock)"
            vnQuery += vbCrLf & "	         inner join Sys_SsoRcvKRHeader_TR krh with(nolock) on krh.OID=krd.RcvKRHOID and krh.RcvKRTypeOID=" & enuRcvKRType.Release_Minus
            vnQuery += vbCrLf & "	   Where krh.TransStatus >= " & enuTCRCKR.Baru & " and krh.RcvPOHOID=" & HdfRcvPOHOID.Value & " and krd.BRGCODE='" & vnBrgCode & "'"
            vnQtyRcvKR = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

            vnQuery = "Update Sys_SsoStockKarantina_TR Set"
            vnQuery += vbCrLf & "QtyKrRelease=" & vnQtyRcvKR & " Where RcvPOHOID=" & vriRcvPOHOID & " and TransCode_Source='" & stuTransCode.SsoPenerimaanPembelian & "' and BRGCODE='" & vnBrgCode & "'"
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vriSQLTrans)

        Next
    End Sub

    Private Sub psUpdate_StockKarantina_QtyReceive(vriRcvKRHOID As String, vriRcvPOHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDtbDetail As New DataTable
        Dim vnQuery As String
        vnQuery = "Select BrgCode From Sys_SsoRcvKRDetail_TR Where RcvKRHOID=" & vriRcvKRHOID
        pbuFillDtbSQLTrans(vnDtbDetail, vnQuery, vriSQLConn, vriSQLTrans)

        Dim vn As Integer
        Dim vnBrgCode As String
        Dim vnQtyRcvKR As Integer
        For vn = 0 To vnDtbDetail.Rows.Count - 1
            vnBrgCode = vnDtbDetail.Rows(vn).Item("BrgCode")
            vnQuery = "Select sum(RcvKRQty)"
            vnQuery += vbCrLf & "       From Sys_SsoRcvKRDetail_TR krd with(nolock)"
            vnQuery += vbCrLf & "	         inner join Sys_SsoRcvKRHeader_TR krh with(nolock) on krh.OID=krd.RcvKRHOID and krh.RcvKRTypeOID in(" & enuRcvKRType.Receive_Minus & "," & enuRcvKRType.Receive_Plus & ")"
            vnQuery += vbCrLf & "	   Where krh.TransStatus >= " & enuTCRCKR.Baru & " and krh.RcvPOHOID=" & HdfRcvPOHOID.Value & " and krd.BRGCODE='" & vnBrgCode & "'"
            vnQtyRcvKR = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

            vnQuery = "Update Sys_SsoStockKarantina_TR Set"
            vnQuery += vbCrLf & "QtyKrReceive=" & vnQtyRcvKR & " Where RcvPOHOID=" & vriRcvPOHOID & " and TransCode_Source='" & stuTransCode.SsoPenerimaanPembelian & "' and BRGCODE='" & vnBrgCode & "'"
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vriSQLTrans)
        Next
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        DivBrg.Visible = True
        GrvDetail.Visible = True

        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "Select" Then
            Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvList.Rows(vnRowIdx)
            TxtTransID.Text = vnGRow.Cells(ensColList.OID).Text

            HdfCompanyCode.Value = Trim(TxtTransID.Text = vnGRow.Cells(4).Text)
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

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "vDelItem" Then
            psDeleteItem()
        ElseIf HdfProcess.Value = "CancelKR" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelKR()
        ElseIf HdfProcess.Value = "psPrepareKR" Then
            psPrepareKR()
        ElseIf HdfProcess.Value = "psApproveKR" Then
            psApproveKR()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psDeleteItem()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnHOID As String = TxtTransID.Text
            Dim vnGRow As GridViewRow
            vnGRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

            Dim vnBrgCode As String = vnGRow.Cells(ensColDetail.BRGCODE).Text

            Dim vnQuery As String
            Dim vnDtb As New DataTable
            Dim vnQtyRcvKR As Integer

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Delete Sys_SsoRcvKRDetail_TR Where RcvKRHOID=" & vnHOID & " and OID=" & vnGRow.Cells(ensColDetail.OID).Text
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)


            If HdfRcvKRTypeOID.Value = enuRcvKRType.Release_Minus Then
                vnQuery = "Select vTotal_RcvKRQty"
                vnQuery += vbCrLf & "       From (Select RcvKRHOID,BRGCODE,sum(RcvKRQty)vTotal_RcvKRQty From Sys_SsoRcvKRDetail_TR with(nolock) Where BRGCODE='" & vnBrgCode & "' group by RcvKRHOID,BRGCODE) krd"
                vnQuery += vbCrLf & "	         inner join Sys_SsoRcvKRHeader_TR krh with(nolock) on krh.OID=krd.RcvKRHOID and krh.RcvKRTypeOID=" & enuRcvKRType.Release_Minus
                vnQuery += vbCrLf & "	   Where krh.TransStatus >= " & enuTCRCKR.Baru & " and krh.RcvPOHOID=" & HdfRcvPOHOID.Value
            Else
                vnQuery = "Select vTotal_RcvKRQty"
                vnQuery += vbCrLf & "       From (Select RcvKRHOID,BRGCODE,sum(RcvKRQty)vTotal_RcvKRQty From Sys_SsoRcvKRDetail_TR with(nolock) Where BRGCODE='" & vnBrgCode & "' group by RcvKRHOID,BRGCODE) krd"
                vnQuery += vbCrLf & "	         inner join Sys_SsoRcvKRHeader_TR krh with(nolock) on krh.OID=krd.RcvKRHOID and krh.RcvKRTypeOID in(" & enuRcvKRType.Receive_Minus & "," & enuRcvKRType.Receive_Plus & ")"
                vnQuery += vbCrLf & "	   Where krh.TransStatus >= " & enuTCRCKR.Baru & " and krh.RcvPOHOID=" & HdfRcvPOHOID.Value
            End If

            vnQtyRcvKR = fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoStockKarantina_TR Set"
            If HdfRcvKRTypeOID.Value = enuRcvKRType.Release_Minus Then
                vnQuery += vbCrLf & "QtyKrRelease="
            Else
                vnQuery += vbCrLf & "QtyKrReceive="
            End If
            vnQuery += vnQtyRcvKR & " Where RcvPOHOID=" & HdfRcvPOHOID.Value & " and TransCode_Source='" & stuTransCode.SsoPenerimaanPembelian & "' and BRGCODE='" & vnBrgCode & "'"
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psFillGrvDetail(TxtTransID.Text, vnSQLConn)

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

    Private Sub BtnPrepare_Click(sender As Object, e As EventArgs) Handles BtnPrepare.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Prepare) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Prepare Penerimaan No. " & TxtPenerimaanNo.Text & " ?<br />WARNING : Prepare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "psPrepareKR"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub

    Private Sub psCancelKR()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnQuery As String
            Dim vnHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvKRHeader_TR set TransStatus=" & enuTCRCKR.Cancelled & ",RcvKRCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusRcvKR(vnHOID, enuTCRCKR.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

            If HdfRcvKRTypeOID.Value = enuRcvKRType.Release_Minus Then
                psUpdate_StockKarantina_QtyRelease(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
            Else
                psUpdate_StockKarantina_QtyReceive(vnHOID, HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans)
            End If

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

    Private Sub psPrepareKR()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psPreparePL", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnQuery As String

            vnQuery = "Select count(1) From Sys_SsoRcvKRDetail_TR Where RcvKRHOID =" & TxtTransID.Text & " and (RcvKRQty=0)"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgError.Text = "Ada Barang Dengan Qty = 0"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            Dim vnRcvKRHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvKRHeader_TR set TransStatus=" & enuTCRCKR.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnRcvKRHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusRcvKR...Start")
            pbuInsertStatusRcvKR(vnRcvKRHOID, enuTCRCKR.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusRcvKR...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Prepare Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            psDisplayData(vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-99")
            vsTextStream.WriteLine("ERROR RAISED")
            vsTextStream.WriteLine(ex.Message)

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("--------------------------------- EOF ---------------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub
    Private Sub psApproveKR()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psApproveKR", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnQuery As String
            Dim vnRcvKRHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvKRHeader_TR set TransStatus=" & enuTCRCKR.Approved & ",ApprovedUserOID=" & Session("UserOID") & ",ApprovedDatetime=getdate() Where OID=" & vnRcvKRHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusRcvKR...Start")
            pbuInsertStatusRcvKR(vnRcvKRHOID, enuTCRCKR.Approved, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusRcvKR...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Prepare Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            psDisplayData(vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-99")
            vsTextStream.WriteLine("ERROR RAISED")
            vsTextStream.WriteLine(ex.Message)

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("--------------------------------- EOF ---------------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub BtnCancelPCL_Click(sender As Object, e As EventArgs) Handles BtnCancelPCL.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan Penerimaan No. " & TxtPenerimaanNo.Text & " ?<br />WARNING : Batal Penerimaan List Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelKR"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub

    Private Sub psShowListBrg(vriBo As Boolean)
        If vriBo Then
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListBrg.Focus()
        Else
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnListPOFind_Click(sender As Object, e As EventArgs) Handles BtnListBrgFind.Click
        LblMsgListBrg.Text = ""
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListBrg(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvListBrg(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        LblMsgListBrg.Text = ""

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnBrg As String = fbuFormatString(Trim(TxtListBrg.Text))

        vnQuery = "Select mbr.BRGCODE,mbr.BRGNAME,mbr.BRGUNIT,"
        vnQuery += vbCrLf & "           abs(skr.QtyKarantina) - (abs(skr.QtyKrRelease)+abs(skr.QtyKrReceive)) vQtyKrOutstanding,skr.OID STKRHOID"
        vnQuery += vbCrLf & "      From " & vnDBMaster & "Sys_MstBarang_MA mbr with(nolock)"
        vnQuery += vbCrLf & "           inner join Sys_SsoStockKarantina_TR skr with(nolock) on skr.BRGCODE=mbr.BRGCODE"
        vnQuery += vbCrLf & "      Where mbr.CompanyCode='" & DstCompany.SelectedValue & "' and"
        vnQuery += vbCrLf & "            skr.RcvPOHOID=" & HdfRcvPOHOID.Value & " and abs(skr.QtyKarantina) > (abs(skr.QtyKrRelease)+abs(skr.QtyKrReceive)) and"

        If RdbTypeRelease.Checked Or RdbTypeMinus.Checked Then
            vnQuery += vbCrLf & "            skr.QtyKarantina < 0"
        Else
            vnQuery += vbCrLf & "            skr.QtyKarantina > 0"
        End If

        vnQuery += vbCrLf & "            and (mbr.BRGCODE like '%" & vnBrg & "%' or mbr.BRGNAME like '%" & vnBrg & "%')"

        vnQuery += vbCrLf & "Order by mbr.BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListBrg.DataSource = vnDtb
        GrvListBrg.DataBind()

        TxtListBrg.Focus()
    End Sub

    Protected Sub BtnListDocClose_Click(sender As Object, e As EventArgs) Handles BtnListBrgClose.Click
        psShowListBrg(False)
    End Sub

    Private Sub GrvListDoc_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListBrg.PageIndexChanging
        GrvListBrg.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListBrg(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListBrg.RowCommand
        DivBrg.Visible = True
        GrvDetail.Visible = True

        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRowList As GridViewRow = GrvListBrg.Rows(vnIdx)

        If e.CommandName = "BRGCODE" Then
            Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListBrg.BRGCODE).Controls(0), LinkButton).Text
            Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

            Dim vnTxtRcvKRQty As TextBox
            vnTxtRcvKRQty = vnGRowDetail.FindControl("TxtRcvKRQty")

            vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = vnKodeBarang
            vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListBrg.BRGNAME).Text

            vnGRowDetail.Cells(ensColDetail.RcvKRQty).Text = vnGRowList.Cells(ensColListBrg.vQtyKrOutstanding).Text
            vnTxtRcvKRQty.Text = vnGRowList.Cells(ensColListBrg.vQtyKrOutstanding).Text

            vnGRowDetail.Cells(ensColDetail.STKRHOID).Text = vnGRowList.Cells(ensColListBrg.STKRHOID).Text

            psShowListBrg(False)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub
    Private Sub psFillGrvLsRcvPO()
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select distinct PM.RcvPONo,Convert(varchar(11),PM.RcvPODate)vRcvPODate,PM.RcvPOSupplierName,PT.RcvPOTypeName,PM.OID,isnull(PM.RcvPORefTypeOID,999)RcvPORefTypeOID,PM.RcvRefTypeOID"
        vnQuery += vbCrLf & " From Sys_SsoRcvPOHeader_TR PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoStockKarantina_TR SK on SK.RcvPOHOID=PM.OID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOType_MA PT on PT.OID=PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & "Where PM.RcvRefTypeOID=" & enuRcvType.Pembelian & " and PM.WarehouseOID=" & DstWhs.SelectedValue & " and"

        If HdfRcvKRTypeOID.Value = enuRcvKRType.Receive_Minus Then
            vnQuery += vbCrLf & "      SK.QtyKarantina < 0 and"
        Else
            vnQuery += vbCrLf & "      SK.QtyKarantina > 0 and"
        End If

        vnQuery += vbCrLf & "      abs(SK.QtyKarantina) > (abs(SK.QtyKrRelease)+abs(SK.QtyKrReceive)) and"
        vnQuery += vbCrLf & "      PM.RcvPOCompanyCode='" & DstCompany.SelectedValue & "' and PM.RcvPONo like '%" & Trim(TxtLsRcvPONo.Text) & "%'"
        vnQuery += vbCrLf & "Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsRcvPO.DataSource = vnDtb
        GrvLsRcvPO.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvLsRcvPO_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsRcvPO.RowCommand
        If e.CommandName = "Select" Then
            psClearMessage()
            If RdbTypeRelease.Checked = False And RdbTypeMinus.Checked = False And RdbTypePlus.Checked = False Then
                LblMsgType.Text = "Pilih Receive Karantina Plus atau Minus"
                Exit Sub
            End If

            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsRcvPO.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(ensColLsRcvPO.RcvPONo).Controls(0), LinkButton).Text
            TxtRcvPONo.Text = vnValue
            HdfRcvPOHOID.Value = vnRow.Cells(ensColLsRcvPO.OID).Text
            psShowLsRcvPO(False)
        End If
    End Sub

    Protected Sub BtnLsRcvPOFind_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOFind.Click
        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If

        psFillGrvLsRcvPO()
    End Sub

    Protected Sub GrvLsRcvPO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsRcvPO.SelectedIndexChanged

    End Sub

    Private Sub GrvLsRcvPO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsRcvPO.PageIndexChanging
        GrvLsRcvPO.PageIndex = e.NewPageIndex
        psFillGrvLsRcvPO()
    End Sub

    Protected Sub BtnRcvPONo_Click(sender As Object, e As EventArgs) Handles BtnRcvPONo.Click
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psShowLsRcvPO(True)
    End Sub

    Protected Sub BtnLsRcvPOClose_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOClose.Click
        psShowLsRcvPO(False)
    End Sub

    Protected Sub RdbTypeRelease_CheckedChanged(sender As Object, e As EventArgs) Handles RdbTypeRelease.CheckedChanged
        If BtnBaru.Visible = False Then
            If RdbTypeRelease.Checked Then
                RdbTypeMinus.Checked = False
                RdbTypePlus.Checked = False
            End If
        End If
    End Sub

    Protected Sub RdbTypeMinus_CheckedChanged(sender As Object, e As EventArgs) Handles RdbTypeMinus.CheckedChanged
        If BtnBaru.Visible = False Then
            If RdbTypeMinus.Checked Then
                RdbTypeRelease.Checked = False
                RdbTypePlus.Checked = False
            End If
        End If
    End Sub

    Protected Sub RdbTypePlus_CheckedChanged(sender As Object, e As EventArgs) Handles RdbTypePlus.CheckedChanged
        If BtnBaru.Visible = False Then
            If RdbTypePlus.Checked Then
                RdbTypeMinus.Checked = False
                RdbTypeRelease.Checked = False
            End If
        End If
    End Sub

    Protected Sub BtnApprove_Click(sender As Object, e As EventArgs) Handles BtnApprove.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Approve) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Approve Penerimaan No. " & TxtPenerimaanNo.Text & " ?<br />WARNING : Approve Tidak Dapat Dibatalkan"
        HdfProcess.Value = "psApproveKR"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub
End Class