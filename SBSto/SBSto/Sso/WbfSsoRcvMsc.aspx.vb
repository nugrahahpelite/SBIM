Imports System.Data.SqlClient
Public Class WbfSsoRcvMsc
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoRcvMsc"
    Const csTNoPrefix = "MR"

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
    End Enum
    Enum ensColDetail
        OID = 0
        vAddItem = 1
        BRGCODE = 2
        BRGNAME = 3
        RcvMscQty = 4
        TxtRcvMscQty = 5
        vDelItem = 6
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPenerimaanDate.Text = ""
        TxtPenerimaanNo.Text = ""

        TxtPenerimaanNote.Text = ""
        HdfTransStatus.Value = enuTCRCMS.Baru
        HdfWarehouseOID.Value = 0
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivCheckBRG.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanCheckBRG.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPenerimaanLain2, vnSQLConn)

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

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_OnReceive.Checked = False And ChkSt_Prepared.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_Prepared.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCRCMS.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCRCMS.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCRCMS.Prepared & ","
        End If
        If ChkSt_OnReceive.Checked = True Then
            vnCrStatus += enuTCRCMS.On_Receive & ","
        End If
        If ChkSt_ReceiveDone.Checked = True Then
            vnCrStatus += enuTCRCMS.Receive_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	SELECT PM.OID,PM.RcvMscNo,convert(varchar(11),PM.RcvMscDate,106)vRcvMscDate, PM.WarehouseOID, PW.WarehouseName,	"
        vnQuery += vbCrLf & "	  PM.RcvMscCompanyCode,	"
        vnQuery += vbCrLf & "	  PM.RcvMscNote,ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared"
        vnQuery += vbCrLf & "FROM Sys_SsoRcvMscHeader_TR PM with(nolock)	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode	"
        vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA PW with(nolock) on PW.OID=PM.WarehouseOID	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID	"
        vnQuery += vbCrLf & "	  left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID	"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.RcvMscCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.RcvMscNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvMscDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvMscDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWhs.SelectedValue
        End If
        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.RcvMscCompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.RcvMscNo"
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
        Dim vnPODOID As String = "0"
        Dim vnCompanyCode As String = ""
        Dim vnPO_NO As String = ""
        Dim vnBRGCODE As String = ""
        Dim vnBRGNAME As String = ""

        Dim vnvDelItem As String = ""
        Dim vnvPrintItem As String = ""

        Dim vnRcvMscQty As Integer = 0

        If vriHOID = "0" Then
            vnQuery = "	 Select '' OID,'' vAddItem,"
            vnQuery += vbCrLf & "	     '' BRGCODE,'' BRGNAME,0 RcvMscQty,	"
            vnQuery += vbCrLf & "	    '' vDelItem"
            vnQuery += vbCrLf & "	  From Sys_SsoRcvMscDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)


        Else
            vnQuery = "	 Select pld.OID,'' vAddItem,"
            vnQuery += vbCrLf & "	     pld.BRGCODE,pld.BRGNAME,pld.RcvMscQty,	"
            vnQuery += vbCrLf & "	    'Hapus Item'vDelItem"
            vnQuery += vbCrLf & "	  From Sys_SsoRcvMscDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & " Where pld.RcvMscHOID =" & vriHOID & " "




            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " and (pld.BRGCODE like '%" & vnCr & "%' or pld.BRGNAME like '%" & vnCr & "%')"
            End If




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

                If HdfTransStatus.Value < enuTCPLSP.On_Receive Then

                Else

                End If
            End If
            GrvDetail.Columns(ensColDetail.RcvMscQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.RcvMscQty).ItemStyle.CssClass = ""

            'GrvDetail.Columns(ensColDetail.PLDSet).HeaderStyle.CssClass = ""
            'GrvDetail.Columns(ensColDetail.PLDSet).ItemStyle.CssClass = ""



            GrvDetail.Columns(ensColDetail.TxtRcvMscQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtRcvMscQty).ItemStyle.CssClass = "myDisplayNone"


        Else
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"



            GrvDetail.Columns(ensColDetail.RcvMscQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.RcvMscQty).ItemStyle.CssClass = "myDisplayNone"



            GrvDetail.Columns(ensColDetail.TxtRcvMscQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtRcvMscQty).ItemStyle.CssClass = ""




            For vn = 0 To 40
                vnDtb.Rows.Add(New Object() {vnOID, vnvAddItem, vnBRGCODE, vnBRGNAME, vnRcvMscQty, vnvDelItem})
            Next
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()

        Dim vnGRow As GridViewRow
        Dim vnTxtRcvMscQty As TextBox

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtRcvMscQty = vnGRow.FindControl("TxtRcvMscQty")
            vnTxtRcvMscQty.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.RcvMscQty).Text)
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



    Private Sub psSetTransNo(vriCompanyCode As String, vriWarehouseCode As String, vriSQLConn As SqlConnection)
        Dim vnQuery As String
        vnQuery = "Select '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(RcvMscNo,len(RcvMscNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(RcvMscNo,len(RcvMscNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoRcvMscHeader_TR with(nolock)"
        vnQuery += vbCrLf & "       Where substring(RcvMscNo,1,len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'))="
        vnQuery += vbCrLf & "                                     '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "                                 and len(RcvMscNo)=len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/')+4"
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
        'If Not IsDate(TxtListStart.Text) Then
        '    TxtListStart.Text = Format(DateAdd(DateInterval.Year, -1, Date.Now), "dd MMM yyyy")
        'End If
        'If Not IsDate(TxtListEnd.Text) Then
        '    TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")
        'End If
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
        LblMsgWhs.Text = ""

        LblMsgPenerimaanDate.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPenerimaanNo.ReadOnly = Not vriBo
        TxtPenerimaanNote.ReadOnly = Not vriBo
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPenerimaanLain2
            vnParam += "&vqTrNo=" & TxtPenerimaanNo.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)

            'Dim vnWinOpen As String
            'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
            'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
            'vnClientScript = Nothing
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
        DivBrg.Visible = False

        GrvDetail.Visible = False

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

        vnQuery = "Select PM.*,convert(varchar(11),PM.RcvMscDate,106)vRcvMscDate,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoRcvMscHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoPenerimaanLain2 & "'"

        vnQuery += vbCrLf & "     Where PM.OID =" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtPenerimaanDate.Text = vnDtb.Rows(0).Item("vRcvMscDate")
            TxtPenerimaanNo.Text = vnDtb.Rows(0).Item("RcvMscNo")
            TxtPenerimaanNote.Text = vnDtb.Rows(0).Item("RcvMscNote")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("RcvMscCompanyCode"))
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

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelPCL.Enabled = False
        BtnPrepare.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCRCMS.Baru)

            BtnCancelPCL.Enabled = (HdfTransStatus.Value = enuTCRCMS.Baru Or HdfTransStatus.Value = enuTCRCMS.Prepared)

            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCRCMS.Baru)
            BtnPreview.Enabled = (HdfTransStatus.Value = enuTCRCMS.Prepared)
            BtnPreview.Enabled = False

            If HdfTransStatus.Value = enuTCRCMS.Baru Then
                BtnPrepare.Text = "Prepare"
            ElseIf HdfTransStatus.Value = enuTCRCMS.Prepared Then
                BtnPrepare.Text = "Prepared"
            End If

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
        If HdfActionStatus.Value = cbuActionNew Then
            psSaveBaru()
        Else
            psSaveEdit()
        End If
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
            Dim vnUserNIP As String = Session("UserName")

            Dim vnRcvMscNo As String

            If HdfActionStatus.Value = cbuActionNew Then
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)
                Dim vnWarehouseOID As Integer = DstWhs.SelectedValue
                Dim warehousecode As String = fbuGetWhsCode_ByOID(vnWarehouseOID, vnSQLConn)
                psSetTransNo(vnCompanyCode, warehousecode, vnSQLConn)
                vnRcvMscNo = Trim(TxtPenerimaanNo.Text)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoRcvMscHeader_TR with(nolock)"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoRcvMscHeader_TR(OID,RcvMscNo,RcvMscDate,"
                vnQuery += vbCrLf & "RcvMscCompanyCode,"
                vnQuery += vbCrLf & "WarehouseOID,"
                vnQuery += vbCrLf & "RcvMscNote,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & vnRcvMscNo & "','" & TxtPenerimaanDate.Text & "',"

                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "',"
                vnQuery += vbCrLf & "'" & Trim(vnWarehouseOID) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoPenerimaanLain2 & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnOID, vnSQLConn, vnSQLTrans)

                pbuInsertStatusRcvMsc(vnOID, enuTCRCMS.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCRCMS.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

            Else
                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoRcvMscHeader_TR set"
                vnQuery += vbCrLf & "RcvMscNote='" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

                pbuInsertStatusRcvMsc(TxtTransID.Text, enuTCRCMS.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

            Dim vnQuery As String
            Dim vnUserNIP As String = Session("EmpNIP")

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            If (HdfTransStatus.Value = enuTCSSOH.Baru) Then
                vnQuery = "Update Sys_SsoRcvMscHeader_TR set "
                vnQuery += vbCrLf & "RcvMscDate='" & TxtPenerimaanDate.Text & "',"
                vnQuery += vbCrLf & "RcvMscNote='" & fbuFormatString(Trim(TxtPenerimaanNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSSOH(TxtTransID.Text, enuTCSSOH.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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
    Private Sub psSaveDetail(vriRcvMscHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow

        Dim vnTxtRcvMscQty As TextBox
        Dim vnBrgCode As String

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtRcvMscQty = vnGRow.FindControl("TxtRcvMscQty")

            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                    vnQuery = "Insert into Sys_SsoRcvMscDetail_TR"
                    vnQuery += vbCrLf & "(RcvMscHOID,"
                    vnQuery += vbCrLf & "BRGCODE,"
                    vnQuery += vbCrLf & "BRGNAME,"
                    vnQuery += vbCrLf & "RcvMscQty"
                    vnQuery += vbCrLf & ")"
                    vnQuery += vbCrLf & "values(" & vriRcvMscHOID & ","
                    vnQuery += vbCrLf & "'" & vnBrgCode & "',"
                    vnQuery += vbCrLf & "'" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGNAME).Text) & "',"
                    vnQuery += vbCrLf & Val(Replace(vnTxtRcvMscQty.Text, "", "")) & ")"

                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                End If
            Else
                vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                vnQuery = "Update Sys_SsoRcvMscDetail_TR SET"
                vnQuery += vbCrLf & "RcvMscQty=" & Val(Replace(vnTxtRcvMscQty.Text, "", "")) & ""

                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            End If
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
            TxtTransID.Text = vnGRow.Cells(0).Text
            TXTBrgRcvMscHOID.Text = vnGRow.Cells(0).Text

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
        ElseIf HdfProcess.Value = "CancelPL" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelMsc()
        ElseIf HdfProcess.Value = "PrepareMsc" Then
            psPrepareMsc()
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
            Dim vnGRow As GridViewRow
            vnGRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Delete Sys_SsoRcvMscDetail_TR Where RcvMscHOID=" & TxtTransID.Text & " and OID=" & vnGRow.Cells(ensColDetail.OID).Text
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
        HdfProcess.Value = "PrepareMsc"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub

    Private Sub psCancelMsc()
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
            Dim vn As Integer

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvMscHeader_TR set TransStatus=" & enuTCRCMS.Cancelled & ",RcvMscCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusRcvMsc(TxtTransID.Text, enuTCRCMS.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psPrepareMsc()
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psPrepareMsc", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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

            vnQuery = "Select count(1) From Sys_SsoRcvMscDetail_TR Where RcvMscHOID =" & TxtTransID.Text & " and (RcvMscQty=0)"
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

            Dim vnRcvMscHOID As String = TxtTransID.Text
            Dim vnWarehouseOID As Integer = DstWhs.SelectedValue

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvMscHeader_TR set TransStatus=" & enuTCRCMS.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnRcvMscHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusRcvMsc...Start")
            pbuInsertStatusRcvMsc(vnRcvMscHOID, enuTCRCMS.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusRcvMsc...End")

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
        HdfProcess.Value = "CancelPL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
        DivBrg.Visible = True
        GrvDetail.Visible = True
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        'psClearMessage()
        'If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        'Dim vnCrpFileName As String = ""
        'psGenerateCrp(vnCrpFileName)

        'Dim vnRootURL As String = ConfigurationManager.AppSettings("WebRootFolder")
        'Dim vnParam As String
        'vnParam = "vqCrpPreviewType=" & stuCrpPreviewType.ByQueryPopwin
        'vnParam += "&vqCrpFileName=" & vnCrpFileName
        'vnParam += "&vqCrpSubReport1="
        'vnParam += "&vqCrpSubReport2="
        'vnParam += "&vqCrpSubReport3="
        'vnParam += "&vqCrpSubReport4="
        'vnParam += "&vqCrpQuery=" & vbuCrpQuery
        'vnParam += "&vqCrpQuery1="
        'vnParam += "&vqCrpQuery2="
        'vnParam += "&vqCrpQuery3="
        'vnParam += "&vqCrpQuery4="
        'vnParam += "&vqCrpPreview=Pdf"

        'vbuPreviewOnClose = "0"

        'ifrPreview.Src = vnRootURL & "Preview/WbfCrpViewer.aspx?" & vnParam
        'psShowPreview(True)
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
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If
        If Trim(TxtListBrg.Text) = "" Then
            LblMsgListBrg.Text = "Pilih Barang"
            Exit Sub
        End If

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
        Dim vnCriteria As String

        Dim vnBrg As String = fbuFormatString(Trim(TxtListBrg.Text))

        vnCriteria = "      Where CompanyCode='" & DstCompany.SelectedValue & "'"
        vnCriteria += vbCrLf & "            and (BRGCODE like '%" & vnBrg & "%' or BRGNAME like '%" & vnBrg & "%')"

        vnQuery = "SELECT BRGCODE,BRGNAME, BRGUNIT FROM " & vnDBMaster & "Sys_MstBarang_MA"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by BRGCODE"
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
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRowList As GridViewRow = GrvListBrg.Rows(vnIdx)
        TxtBrgCode.Text = vnGRowList.Cells(ensColListBrg.BRGCODE).Text


        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        If (Val(TxtTransID.Text) > 0) Then
            Dim vnRcvMscHOID As Integer = Convert.ToInt32(e.CommandArgument)
            If (fsCheckBrgExist(TxtTransID.Text, TxtBrgCode.Text, vnSQLConn) = 1) Then
                LblFindProgress.Text = "Barang yang sama sudah ada di list , silahkan tambah quantity"
                LblFindProgress.Visible = True
                psShowListBrg(False)

                'If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
                'If Trim(TxtTransID.Text) = "" Then Exit Sub
                'If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                '    LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                '    LblMsgError.Visible = True
                '    Exit Sub
                'End If


                'If Not fbuConnectSQL(vnSQLConn) Then
                '    LblMsgError.Text = pbMsgError
                '    LblMsgError.Visible = True
                '    Exit Sub
                'End If

                Session(csModuleName & stuSession.Simpan) = ""

                HdfActionStatus.Value = cbuActionEdit
                psFillGrvDetail(TxtTransID.Text, vnSQLConn)

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing

                psEnableInput(True)
                psEnableSave(True)
                BtnEdit.Visible = False
                BtnBatal.Visible = False

                'psEnableInput(True)

            Else
                If e.CommandName = "BRGCODE" Then
                    Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListBrg.BRGCODE).Controls(0), LinkButton).Text

                    Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

                    vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = vnKodeBarang
                    vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListBrg.BRGNAME).Text

                    psShowListBrg(False)
                End If

            End If
        Else
            If e.CommandName = "BRGCODE" Then
                Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListBrg.BRGCODE).Controls(0), LinkButton).Text

                Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

                vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = vnKodeBarang
                vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListBrg.BRGNAME).Text

                psShowListBrg(False)
            End If
        End If

    End Sub

    Private Function fsCheckBrgExist(vriRcvMscHOID As Integer, vriBrgCode As String, vriSQLConn As SqlConnection) As Integer
        Dim vnChecked As Integer

        psClearMessage()
        DivBrg.Visible = True

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnCompanyCode As String = DstCompany.SelectedValue

        If vriRcvMscHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,0 SOStockQty,0 vSumSOScanQty,0 vSOStockScanVarian,"
            vnQuery += vbCrLf & "       ''vSOStockNote,''vSOStockNoteBy,Null vSOStockNoteDatetime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvCheckBRG.DataSource = vnDtb
            GrvCheckBRG.DataBind()
        Else
            vnQuery = "	 Select pld.OID,''vAddItem,"
            vnQuery += vbCrLf & "	     pld.BRGCODE,pld.BRGNAME,pld.RcvMscQty,	"
            vnQuery += vbCrLf & "	    'Hapus Item'vDelItem"
            vnQuery += vbCrLf & "	  From Sys_SsoRcvMscDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & " Where pld.RcvMscHOID =" & TxtTransID.Text
            vnQuery += vbCrLf & "   and  pld.BRGCODE = ' " & TxtBrgCode.Text & "'"

            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
            GrvCheckBRG.DataSource = vnDtb
            GrvCheckBRG.DataBind()
            If (GrvCheckBRG.Rows.Count) > 0 Then
                vnChecked = 1
            Else
                vnChecked = 0
            End If
        End If

        Return vnChecked
    End Function

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
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

        psFillGrvDetail(TxtTransID.Text, vnSQLConn)
    End Sub
End Class