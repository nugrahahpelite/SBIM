Imports System.Data.SqlClient
Public Class WbfSsoPL
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoPL"
    Const csTNoPrefix = "PL"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
    End Enum

    Enum ensColListSupplier
        SupplierCode = 0
        SupplierName = 1
    End Enum
    Enum ensColListPO
        CompanyCode = 0
        vPODOID = 1
        PO_NO = 2
        vPO_DATE = 3
        ChkSelect = 4
        BRG = 5
        NAMA_BARANG = 6
        QTY = 7
        QTY_PL = 8
        vQTY_PL_Sisa = 9
    End Enum

    Enum ensColDetail
        OID = 0
        vAddItem = 1
        PODOID = 2
        CompanyCode = 3
        PO_NO = 4
        BRGCODE = 5
        BRGNAME = 6
        PLDQty = 7
        TxtPLDQty = 8
        PLDSet = 9
        TxtPLDSet = 10
        PLDCtn = 11
        TxtPLDCtn = 12
        PLDNote = 13
        TxtPLDNote = 14
        vDelItem = 15
        vPrintItem = 16
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPLDate.Text = ""
        TxtPLNo.Text = ""
        TxtPLNoSupp.Text = ""
        TxtPLNote.Text = ""
        TxtTransReceive.Text = ""
        HdfRcvPONo.Value = ""

        TxtPLSupplier.Text = ""
        HdfPLSupplierCode.Value = ""
        HdfPLSupplierName.Value = ""

        HdfTransStatus.Value = enuTCPLSP.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListPO.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListSupplier.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListSupplier.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPOPackingList , vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

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
            vnCrStatus += enuTCPLSP.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPLSP.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCPLSP.Prepared & ","
        End If
        If ChkSt_OnReceive.Checked = True Then
            vnCrStatus += enuTCPLSP.On_Receive & ","
        End If
        If ChkSt_ReceiveDone.Checked = True Then
            vnCrStatus += enuTCPLSP.Receive_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.PLNo,convert(varchar(11),PM.PLDate,106)vPLDate,"
        vnQuery += vbCrLf & "     PM.PLCompanyCode,PM.PLSupplierNo,PM.PLSupplierCode,PM.PLSupplierName,"
        vnQuery += vbCrLf & "     PM.PLNote,RC.RcvPONo,convert(varchar(11),RC.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared"

        vnQuery += vbCrLf & "From Sys_SsoPLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     left outer join Sys_SsoRcvPOHeader_TR RC with(nolock) on RC.RcvPORefOID=PM.OID and RC.RcvRefTypeOID=" & enuRcvType.Pembelian & " and RC.RcvPORefTypeOID=" & enuRcvPOType.Import
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.PLCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.PLNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If
        If Trim(TxtListPONo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.OID in("
            vnQuery += vbCrLf & "              Select pod.PLHOID"
            vnQuery += vbCrLf & "                     From Sys_SsoPLDetail_TR pod with(nolock)"
            vnQuery += vbCrLf & "                          inner join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=pod.POHOID Where poh.PO_NO like '%" & Trim(TxtListPONo.Text) & "%')"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PLDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PLDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.PLCompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.PLDate Desc"
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
        Dim vnPLDQty As String = "0"
        Dim vnPLDSet As String = "0"
        Dim vnPLDCtn As String = "0"
        Dim vnPLDNote As String = ""
        Dim vnvDelItem As String = ""
        Dim vnvPrintItem As String = ""

        vnQuery = "Select pld.OID,''vAddItem,pld.PODOID,"
        vnQuery += vbCrLf & "       poh.CompanyCode,poh.PO_NO,pld.BRGCODE,pld.BRGNAME,pld.PLDQty,pld.PLDSet,pld.PLDCtn,pld.PLDNote,'Hapus Item'vDelItem,'Print'vPrintItem"
        vnQuery += vbCrLf & "  From Sys_SsoPLDetail_TR pld with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoPODetail_TR pod with(nolock) on pod.OID=pld.PODOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=pod.POHOID"
        vnQuery += vbCrLf & " Where pld.PLHOID=" & vriHOID
        vnQuery += vbCrLf & "Order by poh.PO_NO,pld.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        Dim vn As Integer
        If HdfActionStatus.Value = cbuActionNorm Then
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

            If HdfTransStatus.Value = enuTCPLSP.Baru Then
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = "myDisplayNone"
            Else
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

                If HdfTransStatus.Value < enuTCPLSP.On_Receive Then
                    GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = "myDisplayNone"
                    GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = "myDisplayNone"
                Else
                    GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = ""
                    GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = ""
                End If
            End If
            GrvDetail.Columns(ensColDetail.PLDQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.PLDQty).ItemStyle.CssClass = ""

            'GrvDetail.Columns(ensColDetail.PLDSet).HeaderStyle.CssClass = ""
            'GrvDetail.Columns(ensColDetail.PLDSet).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.PLDCtn).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.PLDCtn).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.PLDNote).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.PLDNote).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtPLDQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtPLDQty).ItemStyle.CssClass = "myDisplayNone"

            'GrvDetail.Columns(ensColDetail.TxtPLDSet).HeaderStyle.CssClass = "myDisplayNone"
            'GrvDetail.Columns(ensColDetail.TxtPLDSet).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtPLDCtn).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtPLDCtn).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtPLDNote).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtPLDNote).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.PLDQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.PLDQty).ItemStyle.CssClass = "myDisplayNone"

            'GrvDetail.Columns(ensColDetail.PLDSet).HeaderStyle.CssClass = "myDisplayNone"
            'GrvDetail.Columns(ensColDetail.PLDSet).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.PLDCtn).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.PLDCtn).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.PLDNote).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.PLDNote).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtPLDQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtPLDQty).ItemStyle.CssClass = ""

            'GrvDetail.Columns(ensColDetail.TxtPLDSet).HeaderStyle.CssClass = ""
            'GrvDetail.Columns(ensColDetail.TxtPLDSet).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtPLDCtn).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtPLDCtn).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtPLDNote).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtPLDNote).ItemStyle.CssClass = ""

            For vn = 0 To 40
                vnDtb.Rows.Add(New Object() {vnOID, vnvAddItem, vnPODOID, vnCompanyCode, vnPO_NO, vnBRGCODE, vnBRGNAME, vnPLDQty, vnPLDSet, vnPLDCtn, vnPLDNote, vnvDelItem, vnvPrintItem})
            Next
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()

        Dim vnGRow As GridViewRow
        Dim vnTxtPLDQty As TextBox
        'Dim vnTxtPLDSet As TextBox
        Dim vnTxtPLDCtn As TextBox
        Dim vnTxtPLDNote As TextBox

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtPLDQty = vnGRow.FindControl("TxtPLDQty")
            'vnTxtPLDSet = vnGRow.FindControl("TxtPLDSet")
            vnTxtPLDCtn = vnGRow.FindControl("TxtPLDCtn")
            vnTxtPLDNote = vnGRow.FindControl("TxtPLDNote")

            vnTxtPLDQty.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.PLDQty).Text)
            'vnTxtPLDSet.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.PLDSet).Text)
            vnTxtPLDCtn.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.PLDCtn).Text)
            vnTxtPLDNote.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.PLDNote).Text)
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

    Private Sub psShowListSupplier(vriBo As Boolean)
        If vriBo Then
            DivListSupplier.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListSupplier.Focus()
        Else
            DivListSupplier.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        TxtListSupplier.Focus()
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub psSetTransNo(vriCompanyCode As String, vriSQLConn As SqlConnection)
        Dim vnQuery As String
        'vnQuery = "Select '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        'vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(PLNo,len(PLNo)-3,4))),0)+1))"
        'vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(PLNo,len(PLNo)-3,4))),0)+1 as varchar)"
        'vnQuery += vbCrLf & "        From Sys_SsoPLHeader_TR with(nolock)"
        'vnQuery += vbCrLf & "       Where substring(PLNo,1,len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'))="
        'vnQuery += vbCrLf & "                                     '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        'vnQuery += vbCrLf & "                                 and len(PLNo)=len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/')+4"
        vnQuery = "Select '" & csTNoPrefix & "/" & vriCompanyCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(PLNo,len(PLNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(PLNo,len(PLNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoPLHeader_TR with(nolock)"
        vnQuery += vbCrLf & "       Where substring(PLNo,1,len('" & csTNoPrefix & "/" & vriCompanyCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'))="
        vnQuery += vbCrLf & "                                     '" & csTNoPrefix & "/" & vriCompanyCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "                                 and len(PLNo)=len('" & csTNoPrefix & "/" & vriCompanyCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/')+4"

        TxtPLNo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
    End Sub
    Private Sub psFillGrvListSupplier()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "select distinct mj.SUB SupplierCode,mj.NAMA_SUPPLIER SupplierName"
        vnQuery += vbCrLf & "From Sys_SsoPOHeader_TR mj with(nolock)"

        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=mj.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & " Where (mj.SUB Like '%" & fbuFormatString(Trim(TxtListSupplier.Text)) & "%' or mj.NAMA_SUPPLIER Like '%" & fbuFormatString(Trim(TxtListSupplier.Text)) & "%')"
        vnQuery += vbCrLf & "Order by mj.NAMA_SUPPLIER"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvListSupplier.DataSource = vnDtb
        GrvListSupplier.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        TxtListSupplier.Focus()
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
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
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

        TxtPLDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)

        BtnPLCust.Enabled = True
        BtnPLCust.Visible = True
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgPLNoSupp.Text = ""
        LblMsgPLDate.Text = ""
        LblMsgPLSupplier.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPLNo.ReadOnly = Not vriBo
        TxtPLNote.ReadOnly = Not vriBo
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
        BtnPLCust.Visible = vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
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

            TxtListPOSupplierCode.Text = HdfPLSupplierCode.Value
            TxtListPOSupplierName.Text = HdfPLSupplierName.Value

            psShowListPO(True)

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

        ElseIf e.CommandName = "vPrintItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If
            If Trim(HdfRcvPONo.Value) = "" Then
                LblMsgError.Text = "Nomor Penerimaan Kosong...Print Gagal"
                LblMsgError.Visible = True
                Exit Sub
            End If
            If fsValGenerateQR() = False Then
                Exit Sub
            End If

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnRcvPODate As String = fbuGetRcvPODate_ByPLHOID(TxtTransID.Text, "", vnSQLConn)

            psPreview_QRBarang(TxtTransID.Text, vnGRow.Cells(ensColDetail.OID).Text, vnGRow.Cells(ensColDetail.BRGCODE).Text.Trim() & vbCrLf & HdfRcvPONo.Value & vbCrLf & vnRcvPODate, vnGRow.Cells(ensColDetail.PLDCtn).Text, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPOPackingList
            vnParam += "&vqTrNo=" & TxtPLNo.Text

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
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,convert(varchar(11),PM.PLDate,106)vPLDate,RC.RcvPONo,RC.RcvPONo+' '+convert(varchar(11),RC.RcvPODate,106)vReceive,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoPLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     left outer join Sys_SsoRcvPOHeader_TR RC with(nolock) on RC.RcvPORefOID=PM.OID and RC.RcvRefTypeOID=" & enuRcvType.Pembelian & " and RC.RcvPORefTypeOID=" & enuRcvPOType.Import
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoPOPackingList & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtPLDate.Text = vnDtb.Rows(0).Item("vPLDate")
            TxtPLNo.Text = vnDtb.Rows(0).Item("PLNo")
            TxtPLNote.Text = vnDtb.Rows(0).Item("PLNote")
            TxtTransReceive.Text = fbuValStr(vnDtb.Rows(0).Item("vReceive"))
            HdfRcvPONo.Value = fbuValStr(vnDtb.Rows(0).Item("RcvPONo"))

            TxtPLSupplier.Text = vnDtb.Rows(0).Item("PLSupplierCode") & " " & vnDtb.Rows(0).Item("PLSupplierName")

            HdfPLSupplierCode.Value = vnDtb.Rows(0).Item("PLSupplierCode")
            HdfPLSupplierName.Value = vnDtb.Rows(0).Item("PLSupplierName")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("PLCompanyCode"))

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
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
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCPLSP.Baru)

            BtnCancelPCL.Enabled = (HdfTransStatus.Value = enuTCPLSP.Baru Or HdfTransStatus.Value = enuTCPLSP.Prepared)

            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCPLSP.Baru)
            BtnPreview.Enabled = (HdfTransStatus.Value = enuTCPLSP.Prepared)
            BtnPreview.Enabled = False

            If HdfTransStatus.Value = enuTCPLSP.Baru Then
                BtnPrepare.Text = "Prepare"
            ElseIf HdfTransStatus.Value = enuTCPLSP.Prepared Then
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

        If GrvDetail.Rows(0).Cells(ensColDetail.OID).Text = "0" Then
            BtnPLCust.Enabled = True
        Else
            BtnPLCust.Enabled = False
        End If

        BtnPLCust.Visible = BtnPLCust.Enabled
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        If HdfTransStatus.Value = enuTCPLSP.Baru Then
            psSaveBaru()
        End If
    End Sub

    Private Sub psSaveBaru()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If

        Dim vnSave As Boolean = True
        psClearMessage()
        If Trim(TxtPLNoSupp.Text) = "" Then
            LblMsgPLNoSupp.Text = "Isi PL No"
            vnSave = False
        End If
        If Trim(TxtPLSupplier.Text) = "" Then
            LblMsgPLSupplier.Text = "Isi Supplier"
            vnSave = False
        End If
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If Not IsDate(Trim(TxtPLDate.Text)) Then
            LblMsgPLDate.Text = "Isi Tanggal"
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
            Dim vnUserNIP As String = Session("EmpNIP")

            Dim vnPLNo As String

            If HdfActionStatus.Value = cbuActionNew Then
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)

                psSetTransNo(vnCompanyCode, vnSQLConn)
                vnPLNo = Trim(TxtPLNo.Text)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoPLHeader_TR with(nolock)"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoPLHeader_TR(OID,PLNo,PLDate,"
                vnQuery += vbCrLf & "PLCompanyCode,"
                vnQuery += vbCrLf & "PLSupplierNo,"
                vnQuery += vbCrLf & "PLSupplierCode,PLSupplierName,"
                vnQuery += vbCrLf & "PLNote,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & vnPLNo & "','" & TxtPLDate.Text & "',"

                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "',"
                vnQuery += vbCrLf & "'" & Trim(TxtPLNoSupp.Text) & "',"
                vnQuery += vbCrLf & "'" & HdfPLSupplierCode.Value & "','" & fbuFormatString(Trim(HdfPLSupplierName.Value)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPLNote.Text)) & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoPOPackingList & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                If fsSaveDetail(vnOID, vnCompanyCode, vnSQLConn, vnSQLTrans) = False Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing

                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing

                    Exit Sub
                End If

                pbuInsertStatusPL(vnOID, enuTCPLSP.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCPLSP.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

            Else
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoPLHeader_TR set"
                vnQuery += vbCrLf & "PLDate='" & TxtPLDate.Text & "',"
                vnQuery += vbCrLf & "PLSupplierNo='" & Trim(TxtPLNoSupp.Text) & "',"

                If GrvDetail.Rows(0).Cells(ensColDetail.OID).Text = "0" Then
                    vnQuery += vbCrLf & "PLSupplierCode='" & Trim(HdfPLSupplierCode.Value) & "',"
                    vnQuery += vbCrLf & "PLSupplierName='" & fbuFormatString(Trim(HdfPLSupplierName.Value)) & "',"
                End If

                vnQuery += vbCrLf & "PLNote='" & fbuFormatString(Trim(TxtPLNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                If fsSaveDetail(TxtTransID.Text, vnCompanyCode, vnSQLConn, vnSQLTrans) = False Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing

                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing

                    Exit Sub
                End If

                pbuInsertStatusPL(TxtTransID.Text, enuTCPLSP.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Function fsSaveDetail(vriPLHOID As String, vriCompanyCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow

        Dim vnPOHOID As String

        Dim vnPODOID As String

        Dim vnTxtPLDQty As TextBox
        Dim vnTxtPLDSet As TextBox
        Dim vnTxtPLDCtn As TextBox
        Dim vnTxtPLDNote As TextBox

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnBrgCode As String

        Dim vnPODOID_BrgCode_List As String = ""

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnPODOID = vnGRow.Cells(ensColDetail.PODOID).Text

            vnTxtPLDQty = vnGRow.FindControl("TxtPLDQty")
            vnTxtPLDSet = vnGRow.FindControl("TxtPLDSet")
            vnTxtPLDCtn = vnGRow.FindControl("TxtPLDCtn")
            vnTxtPLDNote = vnGRow.FindControl("TxtPLDNote")
            vnTxtPLDSet.Text = "0"
            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                    vnQuery = "Select count(1) From " & vnDBMaster & "Sys_MstBarang_MA Where CompanyCode='" & vriCompanyCode & "' and BrgCode='" & vnBrgCode & "'"
                    If fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = 0 Then
                        LblMsgError.Text = "KODE BARANG " & vnBrgCode & " TIDAK TERDAFTAR"
                        Return False
                        Exit Function
                    End If

                    If InStr(vnPODOID_BrgCode_List, vnPODOID & "x" & vnBrgCode) = 0 Then
                        vnPODOID_BrgCode_List += "," & vnPODOID & "x" & vnBrgCode

                        vnPOHOID = fbuGetPOHOID_ByPODOID(vnGRow.Cells(ensColDetail.PODOID).Text, vriSQLConn, vriSQLTrans)

                        vnQuery = "Insert into Sys_SsoPLDetail_TR"
                        vnQuery += vbCrLf & "(PLHOID,"
                        vnQuery += vbCrLf & "POHOID,PODOID,"
                        vnQuery += vbCrLf & "BRGCODE,"
                        vnQuery += vbCrLf & "BRGNAME,"
                        vnQuery += vbCrLf & "PLDQty,"
                        vnQuery += vbCrLf & "PLDSet,"
                        vnQuery += vbCrLf & "PLDCtn,"
                        vnQuery += vbCrLf & "PLDNote)"
                        vnQuery += vbCrLf & "values(" & vriPLHOID & ","
                        vnQuery += vbCrLf & vnPOHOID & "," & vnPODOID & ","
                        vnQuery += vbCrLf & "'" & vnBrgCode & "',"
                        vnQuery += vbCrLf & "'" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGNAME).Text) & "',"
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDQty.Text, ",", "")) & ","
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDSet.Text, ",", "")) & ","
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDCtn.Text, ",", "")) & ","
                        vnQuery += vbCrLf & "'" & fbuFormatString(vnTxtPLDNote.Text) & "')"
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                        vnQuery = "Update Sys_SsoPOHeader_TR Set TransStatus=" & enuTCSPPO.In_PL & " Where OID=" & vnPOHOID
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                        pbuPODetail_UpdatePLQty(vnPODOID, vriSQLConn, vriSQLTrans)
                    End If
                End If
            Else
                vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                vnPODOID_BrgCode_List += "," & vnPODOID & "x" & vnBrgCode

                vnQuery = "Update Sys_SsoPLDetail_TR set"
                vnQuery += vbCrLf & "PLDQty=" & Val(Replace(vnTxtPLDQty.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDSet=" & Val(Replace(vnTxtPLDSet.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDCtn=" & Val(Replace(vnTxtPLDCtn.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDNote='" & fbuFormatString(vnTxtPLDNote.Text) & "'"
                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                pbuPODetail_UpdatePLQty(vnPODOID, vriSQLConn, vriSQLTrans)
            End If
        Next
        Return True
    End Function
    Private Sub psSaveDetail_20230908_Orig(vriPLHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow

        Dim vnPOHOID As String

        Dim vnPODOID As String

        Dim vnTxtPLDQty As TextBox
        Dim vnTxtPLDSet As TextBox
        Dim vnTxtPLDCtn As TextBox
        Dim vnTxtPLDNote As TextBox

        Dim vnBrgCode As String

        Dim vnPODOID_BrgCode_List As String = ""

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnPODOID = vnGRow.Cells(ensColDetail.PODOID).Text

            vnTxtPLDQty = vnGRow.FindControl("TxtPLDQty")
            vnTxtPLDSet = vnGRow.FindControl("TxtPLDSet")
            vnTxtPLDCtn = vnGRow.FindControl("TxtPLDCtn")
            vnTxtPLDNote = vnGRow.FindControl("TxtPLDNote")
            vnTxtPLDSet.Text = "0"
            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                    If InStr(vnPODOID_BrgCode_List, vnPODOID & "x" & vnBrgCode) = 0 Then
                        vnPODOID_BrgCode_List += "," & vnPODOID & "x" & vnBrgCode

                        vnPOHOID = fbuGetPOHOID_ByPODOID(vnGRow.Cells(ensColDetail.PODOID).Text, vriSQLConn, vriSQLTrans)

                        vnQuery = "Insert into Sys_SsoPLDetail_TR"
                        vnQuery += vbCrLf & "(PLHOID,"
                        vnQuery += vbCrLf & "POHOID,PODOID,"
                        vnQuery += vbCrLf & "BRGCODE,"
                        vnQuery += vbCrLf & "BRGNAME,"
                        vnQuery += vbCrLf & "PLDQty,"
                        vnQuery += vbCrLf & "PLDSet,"
                        vnQuery += vbCrLf & "PLDCtn,"
                        vnQuery += vbCrLf & "PLDNote)"
                        vnQuery += vbCrLf & "values(" & vriPLHOID & ","
                        vnQuery += vbCrLf & vnPOHOID & "," & vnPODOID & ","
                        vnQuery += vbCrLf & "'" & vnBrgCode & "',"
                        vnQuery += vbCrLf & "'" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGNAME).Text) & "',"
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDQty.Text, ",", "")) & ","
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDSet.Text, ",", "")) & ","
                        vnQuery += vbCrLf & Val(Replace(vnTxtPLDCtn.Text, ",", "")) & ","
                        vnQuery += vbCrLf & "'" & fbuFormatString(vnTxtPLDNote.Text) & "')"
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                        vnQuery = "Update Sys_SsoPOHeader_TR Set TransStatus=" & enuTCSPPO.In_PL & " Where OID=" & vnPOHOID
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                        pbuPODetail_UpdatePLQty(vnPODOID, vriSQLConn, vriSQLTrans)
                    End If
                End If
            Else
                vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                vnPODOID_BrgCode_List += "," & vnPODOID & "x" & vnBrgCode

                vnQuery = "Update Sys_SsoPLDetail_TR set"
                vnQuery += vbCrLf & "PLDQty=" & Val(Replace(vnTxtPLDQty.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDSet=" & Val(Replace(vnTxtPLDSet.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDCtn=" & Val(Replace(vnTxtPLDCtn.Text, ",", "")) & ","
                vnQuery += vbCrLf & "PLDNote='" & fbuFormatString(vnTxtPLDNote.Text) & "'"
                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                pbuPODetail_UpdatePLQty(vnPODOID, vriSQLConn, vriSQLTrans)
            End If
        Next
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "Select" Then
            Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvList.Rows(vnRowIdx)
            TxtTransID.Text = vnGRow.Cells(ensColList.OID).Text

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
            psCancelPL()
        ElseIf HdfProcess.Value = "PreparePL" Then
            psPreparePL()
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

            Dim vnPOHOID As String
            Dim vnPODOID As String

            vnPODOID = vnGRow.Cells(ensColDetail.PODOID).Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnPOHOID = fbuGetPOHOID_ByPODOID(vnPODOID, vnSQLConn, vnSQLTrans)

            vnQuery = "Delete Sys_SsoPLDetail_TR Where PLHOID=" & TxtTransID.Text & " and OID=" & vnGRow.Cells(ensColDetail.OID).Text
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)

            pbuPODetail_UpdatePLQty(vnPODOID, vnSQLConn, vnSQLTrans)

            vnQuery = "Select sum(Qty_PL) From Sys_SsoPODetail_TR with(nolock) Where POHOID=" & vnPOHOID
            If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                vnQuery = "Update Sys_SsoPOHeader_TR set TransStatus=" & enuTCSPPO.Baru & " Where OID=" & vnPOHOID
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)
            End If

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
        LblConfirmMessage.Text = "Anda Prepare Packing List No. " & TxtPLNo.Text & " ?<br/>WARNING : Prepare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "PreparePL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub psCancelPL()
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

            Dim vnDtbPO As New DataTable
            Dim vnPOHOID As Integer

            vnQuery = "Select distinct pod.POHOID"
            vnQuery += vbCrLf & "       From Sys_SsoPLDetail_TR pld with(nolock)"
            vnQuery += vbCrLf & "	         inner join Sys_SsoPODetail_TR pod with(nolock) on pod.OID=pld.PODOID"
            vnQuery += vbCrLf & "	  Where pld.PLHOID=" & TxtTransID.Text
            pbuFillDtbSQL(vnDtbPO, vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPLHeader_TR set TransStatus=" & enuTCPLSP.Cancelled & ",PLCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusPL(TxtTransID.Text, enuTCPLSP.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

            For vn = 0 To vnDtbPO.Rows.Count - 1
                vnPOHOID = vnDtbPO.Rows(vn).Item("POHOID")

                vnQuery = "Select count(1)"
                vnQuery += vbCrLf & "From Sys_SsoPLHeader_TR plh with(nolock)"
                vnQuery += vbCrLf & "     inner join Sys_SsoPLDetail_TR pld with(nolock) on pld.PLHOID=plh.OID"
                vnQuery += vbCrLf & "     inner join Sys_SsoPODetail_TR pod with(nolock) on pod.OID=pld.PODOID"
                vnQuery += vbCrLf & "Where pod.POHOID=" & vnPOHOID & " and plh.TransStatus<>" & enuTCPLSP.Cancelled

                If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                    vnQuery = "Update Sys_SsoPOHeader_TR Set TransStatus=" & enuTCSPPO.Baru & " Where OID=" & vnPOHOID
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                End If
            Next

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

    Private Sub psPreparePL()
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

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
            vnQuery = "Select count(1) From Sys_SsoPLDetail_TR Where PLHOID=" & TxtTransID.Text & " and (PLDQty=0 or PLDCtn=0)"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgError.Text = "Ada Barang Dengan Qty = 0 atau Ctn = 0"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            Dim vnPLHOID As String = TxtTransID.Text
            Dim vnDtbPO As New DataTable
            Dim vnDtbPLD As New DataTable
            Dim vnPOHOID As Integer

            vnQuery = "Select pld.PLHOID,pld.OID,pld.BRGCODE,pld.BRGNAME,poh.PO_NO"
            vnQuery += vbCrLf & "From Sys_SsoPLDetail_TR pld"
            vnQuery += vbCrLf & "     inner join Sys_SsoPODetail_TR pod on pod.OID=pld.PODOID"
            vnQuery += vbCrLf & "  	  inner join Sys_SsoPOHeader_TR poh on poh.OID=pod.POHOID"
            vnQuery += vbCrLf & "Where pld.PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbPLD, vnQuery, vnSQLConn)

            vnQuery = "Select distinct sd.POHOID"
            vnQuery += vbCrLf & "       From Sys_SsoPLDetail_TR pd"
            vnQuery += vbCrLf & "	         inner join Sys_SsoPODetail_TR sd on sd.OID=pd.PODOID"
            vnQuery += vbCrLf & "	  Where pd.PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbPO, vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPLHeader_TR set TransStatus=" & enuTCPLSP.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPL...Start")
            pbuInsertStatusPL(vnPLHOID, enuTCPLSP.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPL...End")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("Loop Update POHeader...Start")
            For vn = 0 To vnDtbPO.Rows.Count - 1
                vnPOHOID = vnDtbPO.Rows(vn).Item("POHOID")

                vnQuery = "Update Sys_SsoPOHeader_TR Set TransStatus=" & enuTCSPPO.In_PL & " Where OID=" & vnPOHOID
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("3." & vn)
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
            Next
            vsTextStream.WriteLine("Loop Update POHeader...End")

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

    Private Function fsValGenerateQR() As Boolean
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

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

            Return False

            Exit Function
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnPLHOID As String = TxtTransID.Text
            Dim vnQuery As String

            vnQuery = "Select TransStatus From Sys_SsoPLHeader_TR Where OID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) <enuTCPLSP.On_Receive Then
                LblMsgError.Text="Penerimaan Pembelian BELUM Diinput"
                                                        LblMsgError.Visible= True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing

                Return False

                Exit Function
            End If

            vnQuery = "Select count(1) From Sys_SsoPLBarangQRCode_TR Where PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("QR Barang sudah ada")
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing

                Return True

                Exit Function
            End If

            Dim vnDtbRcv As New DataTable
            Dim vnRcvOID As String
            Dim vnRcvNo As String
            Dim vnRcvDate As String
            vnQuery = "Select OID,RcvPONo,convert(varchar(11),RcvPODate,106)vRcvPODate From Sys_SsoRcvPOHeader_TR Where RcvPORefTypeOID=" & enuRcvPOType.Import & " and RcvPORefOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbRcv, vnQuery, vnSQLConn)
            vnRcvOID = vnDtbRcv.Rows(0).Item("OID")
            vnRcvNo = vnDtbRcv.Rows(0).Item("RcvPONo")
            vnRcvDate = vnDtbRcv.Rows(0).Item("vRcvPODate")

            Dim vnDtbPLD As New DataTable

            vnQuery = "Select pld.PLHOID,pld.OID,pld.BRGCODE,pld.BRGNAME,poh.PO_NO"
            vnQuery += vbCrLf & "From Sys_SsoPLDetail_TR pld"
            vnQuery += vbCrLf & "     inner join Sys_SsoPODetail_TR pod on pod.OID=pld.PODOID"
            vnQuery += vbCrLf & "  	  inner join Sys_SsoPOHeader_TR poh on poh.OID=pod.POHOID"
            vnQuery += vbCrLf & "Where pld.PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbPLD, vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction("bnsrph")
            vnBeginTrans = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("Loop Generate QRCode...Start")
            Dim vnDRow As DataRow
            Dim vnPLDOID As String
            Dim vnData As String
            For vn = 0 To vnDtbPLD.Rows.Count - 1
                vnDRow = vnDtbPLD.Rows(vn)
                vnPLDOID = vnDRow.Item("OID")
                vnData = vnDRow.Item("BRGCODE") & Space(5) & Chr(10) & vnDRow.Item("BRGNAME") & Chr(10) & "No.Terima:" & vnRcvNo & Chr(10) & "Tgl Terima:" & vnRcvDate & Chr(10) & cbuQR_IDTerima & vnRcvOID

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("4." & vn)
                vsTextStream.WriteLine("vnPLHOID = " & vnPLHOID)
                vsTextStream.WriteLine("vnPLDOID  = " & vnPLDOID)
                vsTextStream.WriteLine("vnData = " & vnData)
                vsTextStream.WriteLine("fsGenBrgQRCode_PL...Start")

                If fsGenBrgQRCode_PL(vnPLHOID, vnPLDOID, vnData, vnSQLConn, vnSQLTrans) = False Then
                    LblMsgError.Text = pbMsgError
                    LblMsgError.Visible = True

                    vsTextStream.WriteLine("fsGenBrgQRCode_PL...Error")
                    vsTextStream.WriteLine(pbMsgError)
                    vnSQLTrans.Rollback()
                    vnSQLTrans = Nothing

                    vsTextStream.WriteLine("")
                    vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                    vsTextStream.WriteLine("--------------------------------- EOF ---------------------------------")

                    vsTextStream.Close()
                    vsTextStream = Nothing
                    vsFso = Nothing

                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing

                    Return False

                    Exit Function
                End If

                vsTextStream.WriteLine("fsGenBrgQRCode_PL...End")
            Next
            vsTextStream.WriteLine("Loop Generate QRCode...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Prepare Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            Return True

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

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            Return False
        End Try
    End Function

    Private Sub psPreparePL_20230526_Orig()
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

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
            vnQuery = "Select count(1) From Sys_SsoPLDetail_TR Where PLHOID=" & TxtTransID.Text & " and (PLDQty=0 or PLDSet=0 or PLDCtn=0)"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgError.Text = "Ada Barang Dengan Qty = 0 atau Set = 0 atau Ctn = 0"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            Dim vnPLHOID As String = TxtTransID.Text
            Dim vnDtbPO As New DataTable
            Dim vnDtbPLD As New DataTable
            Dim vnPOHOID As Integer

            vnQuery = "Select pld.PLHOID,pld.OID,pld.BRGCODE,pld.BRGNAME,poh.PO_NO"
            vnQuery += vbCrLf & "From Sys_SsoPLDetail_TR pld"
            vnQuery += vbCrLf & "     inner join Sys_SsoPODetail_TR pod on pod.OID=pld.PODOID"
            vnQuery += vbCrLf & "  	  inner join Sys_SsoPOHeader_TR poh on poh.OID=pod.POHOID"
            vnQuery += vbCrLf & "Where pld.PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbPLD, vnQuery, vnSQLConn)

            vnQuery = "Select distinct sd.POHOID"
            vnQuery += vbCrLf & "       From Sys_SsoPLDetail_TR pd"
            vnQuery += vbCrLf & "	         inner join Sys_SsoPODetail_TR sd on sd.OID=pd.PODOID"
            vnQuery += vbCrLf & "	  Where pd.PLHOID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtbPO, vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPLHeader_TR set TransStatus=" & enuTCPLSP.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnPLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPL...Start")
            pbuInsertStatusPL(vnPLHOID, enuTCPLSP.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPL...End")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("Loop Update POHeader...Start")
            For vn = 0 To vnDtbPO.Rows.Count - 1
                vnPOHOID = vnDtbPO.Rows(vn).Item("POHOID")

                vnQuery = "Update Sys_SsoPOHeader_TR Set TransStatus=" & enuTCSPPO.In_PL & " Where OID=" & vnPOHOID
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("3." & vn)
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
            Next
            vsTextStream.WriteLine("Loop Update POHeader...End")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("4")
            vsTextStream.WriteLine("Loop Generate QRCode...Start")
            Dim vnDRow As DataRow
            Dim vnPLDOID As String
            Dim vnData As String
            For vn = 0 To vnDtbPLD.Rows.Count - 1
                vnDRow = vnDtbPLD.Rows(vn)
                vnPLDOID = vnDRow.Item("OID")
                vnData = vnDRow.Item("BRGCODE") & Space(5) & Chr(10) & vnDRow.Item("BRGNAME") & Chr(10) & "No.PO :" & vnDRow.Item("PO_NO") & Chr(10) & "Tanggal PL :" & TxtPLDate.Text

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("4." & vn)
                vsTextStream.WriteLine("vnPLHOID = " & vnPLHOID)
                vsTextStream.WriteLine("vnPLDOID  = " & vnPLDOID)
                vsTextStream.WriteLine("vnData = " & vnData)
                vsTextStream.WriteLine("fsGenBrgQRCode_PL...Start")

                If fsGenBrgQRCode_PL(vnPLHOID, vnPLDOID, vnData, vnSQLConn, vnSQLTrans) = False Then
                    vsTextStream.WriteLine("fsGenBrgQRCode_PL...Error")
                    vsTextStream.WriteLine(pbMsgError)
                    vnSQLTrans.Rollback()
                    vnSQLTrans = Nothing

                    vsTextStream.WriteLine("")
                    vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                    vsTextStream.WriteLine("--------------------------------- EOF ---------------------------------")

                    vsTextStream.Close()
                    vsTextStream = Nothing
                    vsFso = Nothing
                    Exit Sub
                End If

                vsTextStream.WriteLine("fsGenBrgQRCode_PL...End")
            Next
            vsTextStream.WriteLine("Loop Generate QRCode...End")

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
        LblConfirmMessage.Text = "Anda Membatalkan Packing List No. " & TxtPLNo.Text & " ?<br />WARNING : Batal Packing List Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelPL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
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

    Private Sub psShowListPO(vriBo As Boolean)
        If vriBo Then
            DivListPO.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListPO.Focus()
        Else
            DivListPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnListCustomerClose_Click(sender As Object, e As EventArgs) Handles BtnListSupplierClose.Click
        psShowListSupplier(False)
    End Sub

    Protected Sub BtnListCustomerFind_Click(sender As Object, e As EventArgs) Handles BtnListSupplierFind.Click
        psFillGrvListSupplier()
    End Sub

    Private Sub GrvListSupplier_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListSupplier.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "SupplierCode" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvListSupplier.Rows(vnIdx)
            TxtPLSupplier.Text = DirectCast(vnRow.Cells(ensColListSupplier.SupplierCode).Controls(0), LinkButton).Text & " " & vnRow.Cells(ensColListSupplier.SupplierName).Text

            HdfPLSupplierCode.Value = DirectCast(vnRow.Cells(ensColListSupplier.SupplierCode).Controls(0), LinkButton).Text
            HdfPLSupplierName.Value = vnRow.Cells(ensColListSupplier.SupplierName).Text

            psShowListSupplier(False)
        End If
    End Sub

    Private Sub GrvListCustomer_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListSupplier.PageIndexChanging
        GrvListSupplier.PageIndex = e.NewPageIndex
        psFillGrvListSupplier()
    End Sub

    Protected Sub BtnListPOFind_Click(sender As Object, e As EventArgs) Handles BtnListPOFind.Click
        LblMsgListPO.Text = ""

        If Trim(TxtListPO.Text) = "" Then
            LblMsgListPO.Text = "Pilih Nomor PO"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListPO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvListPO(vriSQLConn As SqlConnection)
        LblMsgListPO.Text = ""

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String

        vnCriteria = "      Where CompanyCode='" & DstCompany.SelectedValue & "'" 'and QTY > QTY_PL"
        vnCriteria += vbCrLf & "            and SUB='" & TxtListPOSupplierCode.Text & "'"
        vnCriteria += vbCrLf & "            and PO_NO like '%" & fbuFormatString(Trim(TxtListPO.Text)) & "%'"

        vnQuery = "Select CompanyCode,pod.OID vPODOID,PO_NO,convert(varchar(11),PO_DATE,106)vPO_DATE,BRG,NAMA_BARANG,"
        vnQuery += vbCrLf & "            QTY,QTY_PL,(QTY-QTY_PL)vQTY_PL_Sisa"
        vnQuery += vbCrLf & "       From Sys_SsoPOHeader_TR poh"
        vnQuery += vbCrLf & "            inner join Sys_SsoPODetail_TR pod on pod.POHOID=poh.OID"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by PO_NO,BRG"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListPO.DataSource = vnDtb
        GrvListPO.DataBind()

        TxtListPO.Focus()
    End Sub

    Protected Sub BtnListDocClose_Click(sender As Object, e As EventArgs) Handles BtnListPOClose.Click
        psShowListPO(False)
    End Sub

    Private Sub GrvListDoc_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListPO.PageIndexChanging
        GrvListPO.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListPO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListDoc_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListPO.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRowList As GridViewRow = GrvListPO.Rows(vnIdx)

        Dim vnCompanyCode As String = vnGRowList.Cells(ensColListPO.CompanyCode).Text
        Dim vnPONo As String = vnGRowList.Cells(ensColListPO.PO_NO).Text

        If e.CommandName = "BRG" Then
            Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListPO.BRG).Controls(0), LinkButton).Text

            Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)
            Dim vnTxtPLDQty As TextBox = vnGRowDetail.FindControl("TxtPLDQty")

            vnGRowDetail.Cells(ensColDetail.CompanyCode).Text = vnCompanyCode
            vnGRowDetail.Cells(ensColDetail.PO_NO).Text = vnGRowList.Cells(ensColListPO.PO_NO).Text
            vnGRowDetail.Cells(ensColDetail.PODOID).Text = vnGRowList.Cells(ensColListPO.vPODOID).Text
            vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = vnKodeBarang
            vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListPO.NAMA_BARANG).Text

            If Val(vnGRowList.Cells(ensColListPO.vQTY_PL_Sisa).Text) < 0 Then
                vnTxtPLDQty.Text = "0"
            Else
                vnTxtPLDQty.Text = vnGRowList.Cells(ensColListPO.vQTY_PL_Sisa).Text
            End If

            psShowListPO(False)
        End If
    End Sub

    Private Sub BtnPLCust_Click(sender As Object, e As EventArgs) Handles BtnPLCust.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        psClearMessage()
        If HdfActionStatus.Value = cbuActionNorm Then Exit Sub

        psShowListSupplier(True)
    End Sub

    Private Function fsGenBrgQRCode_PL(vriPLHOID As String, vriPLDOID As String, vriQRData As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        Dim vnReturn As Boolean
        Try
            Dim vnQuery As String

            Dim vsIOFileStream As System.IO.FileStream
            Dim vsFileLength As Long
            Const csFileFormat = ".jpg"

            Dim vnCmd As SqlCommand
            Dim vnFileName As String
            Dim vnFileByte() As Byte

            vnFileName = "PLBarangQR_" & vriPLHOID & "_" & vriPLDOID & "_" & Format(Date.Now, "yyyyMMdd_HHmmss") & "~sm" & csFileFormat

            Dim vnQRDir As String = ""

            pbuGenerateQRCode(vnFileName, vriQRData, vnQRDir)

            vsIOFileStream = System.IO.File.OpenRead(vnQRDir & vnFileName)

            vsFileLength = vsIOFileStream.Length
            ReDim vnFileByte(vsFileLength)

            vsIOFileStream.Read(vnFileByte, 0, vsFileLength)

            vnQuery = "Insert into Sys_SsoPLBarangQRCode_TR"
            vnQuery += vbCrLf & "(PLHOID,PLDOID,BRGCODEQRCodeImg)"
            vnQuery += vbCrLf & "Values("
            vnQuery += vbCrLf & "'" & vriPLHOID & "','" & vriPLDOID & "',@vnBRGQRCodeImg"
            vnQuery += vbCrLf & ")"

            vnCmd = New SqlClient.SqlCommand(vnQuery, vriSQLConn, vriSQLTrans)
            vnCmd.Parameters.AddWithValue("@vnBRGQRCodeImg", vnFileByte)
            vnCmd.Transaction = vriSQLTrans
            vnCmd.ExecuteNonQuery()

            vnReturn = True

            Return vnReturn
        Catch ex As Exception
            pbMsgError = ex.Message
            Return False
        End Try
    End Function

    Private Sub psPreview_QRBarang(vriPLHOID As String, vriPLDOID As String, vriBarangCode As String, vriPrintCount As Integer, vriSQLConn As SqlConnection)
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        Dim vnCrpFileName As String = ""
        psGenerateCrp_QRBarang(vnCrpFileName, vriPLHOID, vriPLDOID, vriBarangCode, vriPrintCount, vriSQLConn)

        Dim vnRootURL As String = ConfigurationManager.AppSettings("WebRootFolder")
        Dim vnParam As String
        vnParam = "vqCrpPreviewType=" & stuCrpPreviewType.ByQueryPopwin
        vnParam += "&vqCrpFileName=" & vnCrpFileName
        vnParam += "&vqCrpSubReport1="
        vnParam += "&vqCrpSubReport2="
        vnParam += "&vqCrpSubReport3="
        vnParam += "&vqCrpSubReport4="
        vnParam += "&vqCrpQuery=" & vbuCrpQuery
        vnParam += "&vqCrpQuery1="
        vnParam += "&vqCrpQuery2="
        vnParam += "&vqCrpQuery3="
        vnParam += "&vqCrpQuery4="
        vnParam += "&vqCrpPreview=Pdf"

        vbuPreviewOnClose = "0"

        ifrPreview.Src = vnRootURL & "Preview/WbfCrpViewer.aspx?" & vnParam
        psShowPreview(True)
    End Sub

    Private Sub psGenerateCrp_QRBarang(ByRef vriCrpFileName As String, vriPLHOID As String, vriPLDOID As String, vriBarangCode As String, vriPrintCount As Integer, vriSQLConn As SqlConnection)
        'Barcode print 2x (printer Barcode)
        Dim vnUserOID As String = Session("UserOID")
        Dim vnQuery As String
        Dim vn As Integer
        Dim vnCount As Integer = vriPrintCount

        vnCount = Math.Ceiling(Val(vriPrintCount) / 2)

        Dim vnSQLTrans As SqlTransaction = Nothing

        Try
            vnSQLTrans = vriSQLConn.BeginTransaction("inp")
            vnQuery = "Delete Sys_SsoPrintQRBarang_Temp Where UserOID=" & vnUserOID
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)

            For vn = 0 To vnCount - 2
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriBarangCode & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       1,'" & vriBarangCode & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "  From Sys_SsoPLBarangQRCode_TR Where PLHOID='" & vriPLHOID & "' and PLDOID='" & vriPLDOID & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            Next

            vn = vn + 1
            If Val(vriPrintCount) Mod 2 = 0 Then
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriBarangCode & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       1,'" & vriBarangCode & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "  From Sys_SsoPLBarangQRCode_TR Where PLHOID='" & vriPLHOID & "' and PLDOID='" & vriPLDOID & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            Else
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriBarangCode & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       0,'" & vriBarangCode & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "  From Sys_SsoPLBarangQRCode_TR Where PLHOID='" & vriPLHOID & "' and PLDOID='" & vriPLDOID & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            End If

            vriCrpFileName = stuSsoCrp.CrpBnsrphBarcodeSelectionQR

            vbuCrpQuery = "Select * From Sys_SsoPrintQRBarang_Temp with(nolock) Where UserOID=" & vnUserOID

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing

        Catch ex As Exception

            vnSQLTrans.Rollback()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
        End Try
    End Sub

    Protected Sub BtnListPOSelect_Click(sender As Object, e As EventArgs) Handles BtnListPOSelect.Click
        If GrvListPO.Rows.Count > 0 Then
            Dim vnChkSelect As CheckBox
            Dim vnGRowList As GridViewRow
            Dim vnGRowDetail As GridViewRow
            Dim vnTxtPLDQty As TextBox

            Dim vnCurrentRowIdx As Integer = Val(HdfDetailRowIdx.Value)

            For vn = 0 To GrvListPO.Rows.Count - 1
                vnGRowList = GrvListPO.Rows(vn)
                vnChkSelect = vnGRowList.FindControl("ChkSelect")
                If vnChkSelect.Checked Then
                    vnGRowDetail = GrvDetail.Rows(vnCurrentRowIdx)

                    vnGRowDetail.Cells(ensColDetail.CompanyCode).Text = vnGRowList.Cells(ensColListPO.CompanyCode).Text
                    vnGRowDetail.Cells(ensColDetail.PO_NO).Text = vnGRowList.Cells(ensColListPO.PO_NO).Text
                    vnGRowDetail.Cells(ensColDetail.PODOID).Text = vnGRowList.Cells(ensColListPO.vPODOID).Text
                    vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = DirectCast(vnGRowList.Cells(ensColListPO.BRG).Controls(0), LinkButton).Text
                    vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListPO.NAMA_BARANG).Text

                    vnTxtPLDQty = vnGRowDetail.FindControl("TxtPLDQty")
                    If Val(vnGRowList.Cells(ensColListPO.vQTY_PL_Sisa).Text) < 0 Then
                        vnTxtPLDQty.Text = "0"
                    Else
                        vnTxtPLDQty.Text = vnGRowList.Cells(ensColListPO.vQTY_PL_Sisa).Text
                    End If
                    vnCurrentRowIdx = vnCurrentRowIdx + 1
                End If
            Next

            psShowListPO(False)
        End If
    End Sub
End Class