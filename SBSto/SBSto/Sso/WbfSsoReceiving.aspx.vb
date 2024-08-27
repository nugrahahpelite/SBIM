Imports System.Data.SqlClient
Public Class WbfSsoReceiving
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoReceiving"
    Const csTNoPrefix = "RCV"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String
    Enum ensColList
        OID = 0
    End Enum

    Enum ensColDetail
        OID = 0
        vAddItem = 1
        BRGCODE = 2
        BRGNAME = 3
        BRGUNIT = 4
        RcvDQty = 5
        TxtRcvDQty = 6
        vSumRcvScanQty = 7
        vSumRcvScanVarian = 8
        vRcvDNote = 9
        TxtvRcvDNote = 10
        vRcvDNoteBy = 11
        vRcvDNoteDatetime = 12
        vDelItem = 13
    End Enum

    Enum ensColListItem
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
    End Enum
    Enum ensColLsScan
        vRcvScanDeleted = 4
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtRcvDate.Text = ""
        TxtRcvNo.Text = ""
        TxtRcvNote.Text = ""

        HdfTransStatus.Value = enuTCSRCV.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListItem.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListItem.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsScan.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsScan.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrOption.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Sso"
        If Not IsPostBack Then
            psDefaultDisplay()
            psFillDstSOReport()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoReceiving, vnSQLConn)

            pbuFillDstRcvType(DstRcvRefType, False, vnSQLConn)

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

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_Closed.Checked = False And ChkSt_ScanClosed.Checked = False And ChkSt_ScanOpen.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_ScanOpen.Checked = True
            ChkSt_ScanClosed.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCSRCV.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCSRCV.Cancelled & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCSRCV.Closed & ","
        End If
        If ChkSt_ScanClosed.Checked = True Then
            vnCrStatus += enuTCSRCV.Scan_Closed & ","
        End If
        If ChkSt_ScanOpen.Checked = True Then
            vnCrStatus += enuTCSRCV.Scan_Open & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.RcvNo,convert(varchar(11),PM.RcvDate,106)vRcvDate,"
        vnQuery += vbCrLf & "     RT.RcvTypeName,PM.RcvRefNo,convert(varchar(11),PM.RcvRefDate,106)vRcvRefDate,"
        vnQuery += vbCrLf & "     PM.RcvCompanyCode,WM.WarehouseName,SW.SubWhsName,PM.RcvNote,PM.RcvCloseNote,PM.RcvCancelNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanOpenDatetime,106)+' '+convert(varchar(5),PM.ScanOpenDatetime,108)+' '+ PR.UserName vScanOpen,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanClosedDatetime,106)+' '+convert(varchar(5),PM.ScanClosedDatetime,108)+' '+ AP.UserName vScanClosed,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ CL.UserName vClosed"

        vnQuery += vbCrLf & "From Sys_SsoRcvHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.RcvWarehouseOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SW with(nolock) on SW.OID=PM.RcvSubWarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoRcvType_MA RT with(nolock) on RT.OID=PM.RcvRefTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.ScanOpenUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA AP with(nolock) on AP.OID=PM.ScanClosedUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CL with(nolock) on CL.OID=PM.ClosedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.RcvCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.RcvWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListRTNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.RcvNo like '%" & Trim(TxtListRTNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvDate <= '" & TxtListEnd.Text & "'"
        End If
        vnQuery += vbCrLf & "Order by PM.RcvNo"
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

        Dim vnDOID As String = "0"
        Dim vnvAddItem As String = "..."
        Dim vnBRGCODE As String = ""
        Dim vnBRGNAME As String = ""
        Dim vnBRGUNIT As String = ""
        Dim vnRcvDQty As String = "0"
        Dim vnvSumRcvScanQty As String = "0"
        Dim vnvSumRcvScanVarian As String = "0"
        Dim vnRcvDNote As String = ""
        Dim vnvRcvDNoteBy As String = ""
        Dim vnvRcvDNoteDatetime As String = ""
        Dim vnvDelItem As String = ""

        If vriHOID = 0 Then
            vnQuery = "Select 0 OID,''vAddItem,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,"
            vnQuery += vbCrLf & "       0 RcvDQty,0 vSumRcvScanQty,0 vSumRcvScanVarian,''vRcvDNote,''vRcvDNoteBy,''vRcvDNoteDatetime,''vDelItem Where 1=2"
        Else
            vnQuery = "Select d.OID,''vAddItem,"
            vnQuery += vbCrLf & "       d.BRGCODE,mb.BRGNAME,mb.BRGUNIT,"
            vnQuery += vbCrLf & "       d.RcvDQty,d.vSumRcvScanQty,d.vSumRcvScanVarian,d.vRcvDNote,d.vRcvDNoteBy,d.vRcvDNoteDatetime,'Hapus Item'vDelItem"
            vnQuery += vbCrLf & "       From fnTbl_SsoRcvDetailScan(" & vriHOID & ") d"
            vnQuery += vbCrLf & "            inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb on mb.BRGCODE=d.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "'"

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " Where d.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%'"
            End If
            If ChkFindVarian.Checked Then
                vnQuery += vbCrLf & " and d.vSumRcvScanVarian<>0"
            End If

            vnQuery += vbCrLf & "Order by d.OID"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        Dim vn As Integer
        If HdfActionStatus.Value = cbuActionNorm Then
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

            If HdfTransStatus.Value = enuTCSRCV.Baru Then
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = ""
            Else
                GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"
            End If

            GrvDetail.Columns(ensColDetail.RcvDQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.RcvDQty).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vRcvDNote).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vRcvDNote).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtRcvDQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtRcvDQty).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtvRcvDNote).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtvRcvDNote).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

            If HdfTransStatus.Value = enuTCSRCV.Baru Then
                GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.RcvDQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.RcvDQty).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.vRcvDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vRcvDNote).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtRcvDQty).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtRcvDQty).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.TxtvRcvDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtvRcvDNote).ItemStyle.CssClass = ""

                For vn = 0 To 9
                    vnDtb.Rows.Add(New Object() {vnDOID, vnvAddItem, vnBRGCODE, vnBRGNAME, vnBRGUNIT, vnRcvDQty, vnvSumRcvScanQty, vnvSumRcvScanVarian, vnRcvDNote, vnvRcvDNoteBy, vnvRcvDNoteDatetime, vnvDelItem})
                Next

            Else
                GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.RcvDQty).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.RcvDQty).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.vRcvDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vRcvDNote).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtRcvDQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.TxtRcvDQty).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtvRcvDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtvRcvDNote).ItemStyle.CssClass = ""

            End If
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()

        Dim vnGRow As GridViewRow
        If HdfActionStatus.Value = cbuActionEdit Then
            Dim vnTxtRcvDQty As TextBox
            Dim vnTxtvRcvDNote As TextBox

            For vn = 0 To GrvDetail.Rows.Count - 1
                vnGRow = GrvDetail.Rows(vn)
                vnTxtRcvDQty = vnGRow.FindControl("TxtRcvDQty")
                vnTxtvRcvDNote = vnGRow.FindControl("TxtvRcvDNote")

                vnTxtRcvDQty.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.RcvDQty).Text)
                vnTxtvRcvDNote.Text = Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vRcvDNote).Text), "<br />", Chr(10))
            Next
        End If
    End Sub

    Private Sub psFillGrvListItem()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable
        vnQuery = "select distinct BRGCODE,BRGNAME,BRGUNIT from " & fbuGetDBMaster() & "Sys_MstBarang_MA where CompanyCode='" & DstCompany.SelectedValue & "'"

        If Trim(TxtListItem.Text) <> "" Then
            vnQuery += vbCrLf & " and (BRGCODE Like '%" & fbuFormatString(Trim(TxtListItem.Text)) & "%' or BRGNAME Like '%" & fbuFormatString(Trim(TxtListItem.Text)) & "%')"
        End If

        vnQuery += vbCrLf & "Order by BRGNAME"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvListItem.DataSource = vnDtb
        GrvListItem.DataBind()
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        psFillGrvList()
    End Sub

    Protected Sub BtnListClose_Click(sender As Object, e As EventArgs) Handles BtnListClose.Click
        psShowList(False)
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

    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        If Not IsDate(TxtListStart.Text) Then
            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -1, Date.Now), "dd MMM yyyy")
        End If
        If Not IsDate(TxtListEnd.Text) Then
            TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")
        End If
        psShowList(True)
    End Sub

    Private Sub psSetTransNo(vriCompanyCode As String, vriSubWhsCode As String, vriSQLConn As SqlConnection)
        Dim vnTNoPrefix As String = csTNoPrefix & "/" & vriCompanyCode & "/" & vriSubWhsCode & "/'+substring(convert(varchar(10),getdate(),111),3,10)"
        Dim vnQuery As String
        vnQuery = "Select '" & vnTNoPrefix & "+'/'"
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(RcvNo,len(RcvNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(RcvNo,len(RcvNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoRcvHeader_TR with(nolock)"
        vnQuery += vbCrLf & "      Where RcvNo like '" & vnTNoPrefix & "+'/%'"
        TxtRcvNo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
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

        TxtRcvDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        If DstCompany.Items.Count > 0 Then
            DstCompany.SelectedIndex = 0
        End If

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
        LblMsgRcvDate.Text = ""
        LblMsgRcvNo.Text = ""
        LblMsgRcvRefNo.Text = ""
        LblMsgRcvRefType.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtRcvRefNo.ReadOnly = Not vriBo
        TxtRcvRefDate.ReadOnly = Not vriBo
        TxtRcvNote.ReadOnly = Not vriBo
        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelRcv.Visible = Not vriBo
        BtnScanOpen.Visible = Not vriBo
        BtnScanClosed.Visible = Not vriBo
        BtnCloseRcv.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "vAddItem" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx

            psShowListItem(True)

        ElseIf e.CommandName = "vDelItem" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx
            LblConfirmMessage.Text = "Anda Hapus Item " & vnGRow.Cells(ensColDetail.BRGCODE).Text & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text & " ?"
            HdfProcess.Value = "vDelItem"
            TxtConfirmNote.Visible = False
            LblConfirmNote.Visible = False
            psShowConfirm(True)

        ElseIf e.CommandName = "vSumRcvScanQty" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx
            HdfLsScanBrgCode.Value = vnGRow.Cells(ensColDetail.BRGCODE).Text
            psFillGrvLsScan(HdfLsScanBrgCode.Value)
            LblLsScanTitle.Text = "SCAN " & vnGRow.Cells(ensColDetail.BRGCODE).Text & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text

            psShowLsScan(True)
        End If
    End Sub

    Private Sub psShowConfirm(vriBo As Boolean)
        If vriBo Then
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psShowLsScan(vriBo As Boolean)
        If vriBo Then
            DivLsScan.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsScan.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psShowPrOption(vriBo As Boolean)
        If vriBo Then
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psShowListItem(vriBo As Boolean)
        If vriBo Then
            DivListItem.Style(HtmlTextWriterStyle.Visibility) = "visible"
            'psFillGrvListItem()
        Else
            DivListItem.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnListItemClose_Click(sender As Object, e As EventArgs) Handles BtnListItemClose.Click
        psShowListItem(False)
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoReceiving
            vnParam += "&vqTrNo=" & TxtRcvNo.Text

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

        psShowListItem(False)
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,convert(varchar(11),PM.RcvDate,106)vRcvDate,convert(varchar(11),PM.RcvRefDate,106)vRcvRefDate,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoRcvHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoReceiving & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtRcvDate.Text = vnDtb.Rows(0).Item("vRcvDate")
            TxtRcvNo.Text = vnDtb.Rows(0).Item("RcvNo")
            TxtRcvNote.Text = vnDtb.Rows(0).Item("RcvNote")

            DstRcvRefType.SelectedValue = vnDtb.Rows(0).Item("RcvRefTypeOID")
            TxtRcvRefNo.Text = vnDtb.Rows(0).Item("RcvRefNo")
            TxtRcvRefDate.Text = vnDtb.Rows(0).Item("vRcvRefDate")

            DstCompany.SelectedValue = vnDtb.Rows(0).Item("RcvCompanyCode")

            pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vriSQLConn)
            DstSubWhs.SelectedValue = vnDtb.Rows(0).Item("RcvSubWarehouseOID")
            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        GrvDetail.PageIndex = 0
        psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnCancelRcv.Enabled = False
        BtnScanOpen.Enabled = False
        BtnScanClosed.Enabled = False
        BtnCloseRcv.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelRcv.Visible = BtnCancelRcv.Enabled
        BtnScanOpen.Visible = BtnScanOpen.Enabled
        BtnScanClosed.Visible = BtnScanClosed.Enabled
        BtnCloseRcv.Visible = BtnCloseRcv.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelRcv.Enabled = False
        BtnCloseRcv.Enabled = False
        BtnScanOpen.Enabled = False
        BtnScanClosed.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCSRCV.Baru Or HdfTransStatus.Value = enuTCSRCV.Scan_Open Or HdfTransStatus.Value = enuTCSRCV.Scan_Closed)

            BtnCancelRcv.Enabled = (HdfTransStatus.Value = enuTCSRCV.Baru Or HdfTransStatus.Value = enuTCSRCV.Scan_Closed)

            BtnCancelRcv.Enabled = (HdfTransStatus.Value = enuTCSRCV.Baru)

            BtnScanOpen.Enabled = (HdfTransStatus.Value = enuTCSRCV.Baru Or HdfTransStatus.Value = enuTCSRCV.Scan_Closed)
            BtnScanClosed.Enabled = (HdfTransStatus.Value = enuTCSRCV.Scan_Open)
            BtnCloseRcv.Enabled = (HdfTransStatus.Value = enuTCSRCV.Scan_Closed)

            BtnPreview.Enabled = (HdfTransStatus.Value >= enuTCSRCV.Closed)

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
        If HdfTransStatus.Value = enuTCSRCV.Baru Then
            psSaveBaru()
        Else
            psSaveBeforeClose()
        End If
    End Sub

    Private Sub psSaveBeforeClose()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If
        psClearMessage()

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

            vnSQLTrans = vnSQLConn.BeginTransaction()

            psSaveDetailAfterKembali(vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            psEnableInput(False)
            psEnableSave(False)
            psButtonStatus()

            HdfActionStatus.Value = cbuActionNorm
            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psButtonStatus()
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

    Private Sub psSaveDetailAfterKembali(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnTxtvRcvDNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtvRcvDNote = vnGRow.FindControl("TxtvRcvDNote")

            vnQuery = "Update Sys_SsoRcvDetail_TR set"
            vnQuery += vbCrLf & "RcvDNote='" & fbuFormatString(vnTxtvRcvDNote.Text) & "',"

            vnQuery += vbCrLf & "RcvDNoteDatetime=getdate(),RcvDNoteUserOID=" & Session("UserOID")

            vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
        Next
    End Sub

    Private Sub psSaveBaru()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If
        Dim vnSave As Boolean = True
        psClearMessage()

        If Not IsDate(TxtRcvDate.Text) Then
            LblMsgRcvDate.Text = "Isi Tanggal Penerimaan"
            vnSave = False
        End If
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "" Then
            LblMsgSubWhs.Text = "Pilih Sub Warehouse"
            vnSave = False
        End If
        If Trim(TxtRcvRefNo.Text) = "" Then
            TxtRcvRefNo.Text = "Isi Nomor dan Tanggal Ref"
            vnSave = False
        End If
        If Not IsDate(TxtRcvRefDate.Text) Then
            TxtRcvRefDate.Text = "Isi Nomor dan Tanggal Ref"
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

            If HdfActionStatus.Value = cbuActionNew Then
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                Dim vnCompanyCode As String = DstCompany.SelectedValue
                Dim vnSubWhsOID As String = DstSubWhs.SelectedValue

                Dim vnWarehouseOID As String = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID, vnSQLConn)
                Dim vnSubWhsCode As String = fbuGetSubWhsCode_ByOID(vnSubWhsOID, vnSQLConn)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoRcvHeader_TR"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                psSetTransNo(vnCompanyCode, vnSubWhsCode, vnSQLConn)

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoRcvHeader_TR(OID,RcvNo,RcvDate,"
                vnQuery += vbCrLf & "RcvCompanyCode,RcvWarehouseOID,RcvSubWarehouseOID,"
                vnQuery += vbCrLf & "RcvNote,"
                vnQuery += vbCrLf & "RcvRefTypeOID,RcvRefNo,RcvRefDate,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & Trim(TxtRcvNo.Text) & "','" & TxtRcvDate.Text & "',"
                vnQuery += vbCrLf & "'" & vnCompanyCode & "','" & vnWarehouseOID & "','" & vnSubWhsOID & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtRcvNote.Text)) & "',"
                vnQuery += vbCrLf & DstRcvRefType.SelectedValue & ",'" & fbuFormatString(Trim(TxtRcvRefNo.Text)) & "','" & TxtRcvRefDate.Text & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoReceiving & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnOID, vnSQLConn, vnSQLTrans)

                pbuInsertStatusReceiving(vnOID, enuTCSRCV.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCSRCV.Baru
            Else
                Dim vnSubWhsOID As String = DstSubWhs.SelectedValue
                Dim vnWarehouseOID As String = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID, vnSQLConn)

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update Sys_SsoRcvHeader_TR set"
                vnQuery += vbCrLf & "RcvDate='" & TxtRcvDate.Text & "',"
                vnQuery += vbCrLf & "RcvNote='" & fbuFormatString(Trim(TxtRcvNote.Text)) & "',"

                vnQuery += vbCrLf & "RcvRefTypeOID=" & DstRcvRefType.SelectedValue & ","
                vnQuery += vbCrLf & "RcvRefNo='" & fbuFormatString(Trim(TxtRcvRefNo.Text)) & "',"
                vnQuery += vbCrLf & "RcvRefDate='" & TxtRcvRefDate.Text & "',"

                If (HdfTransStatus.Value = enuTCSRCV.Baru) Then
                    vnQuery += vbCrLf & "RcvWarehouseOID='" & vnWarehouseOID & "',"
                    vnQuery += vbCrLf & "RcvSubWarehouseOID='" & vnSubWhsOID & "',"
                End If
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

                pbuInsertStatusReceiving(TxtTransID.Text, enuTCSRCV.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psSaveDetail(vriOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnTxtRcvDQty As TextBox
        Dim vnTxtvRcvDNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtRcvDQty = vnGRow.FindControl("TxtRcvDQty")
            vnTxtvRcvDNote = vnGRow.FindControl("TxtvRcvDNote")
            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnQuery = "Insert into Sys_SsoRcvDetail_TR"
                    vnQuery += vbCrLf & "(RcvHOID,"
                    vnQuery += vbCrLf & "BRGCODE,BRGNAME,BRGUNIT,"
                    vnQuery += vbCrLf & "RcvDQty,"
                    vnQuery += vbCrLf & "RcvDNote,RcvDNoteUserOID,RcvDNoteDatetime)"
                    vnQuery += vbCrLf & "values(" & vriOID & ","
                    vnQuery += vbCrLf & "'" & vnGRow.Cells(ensColDetail.BRGCODE).Text & "','" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGNAME).Text) & "','" & fbuFormatString(vnGRow.Cells(ensColDetail.BRGUNIT).Text) & "',"
                    vnQuery += vbCrLf & Val(vnTxtRcvDQty.Text) & ","
                    vnQuery += vbCrLf & "'" & fbuFormatString(vnTxtvRcvDNote.Text) & "'," & Session("UserOID") & ",getdate())"
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                End If
            Else
                If Trim(vnTxtvRcvDNote.Text) <> Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vRcvDNote).Text), "<br />", vbLf) Then
                    vnQuery = "Update Sys_SsoRcvDetail_TR set "
                    vnQuery += vbCrLf & "RcvDQty=" & Val(vnTxtRcvDQty.Text) & ","
                    vnQuery += vbCrLf & "RcvDNote='" & fbuFormatString(vnTxtvRcvDNote.Text) & "',"
                    vnQuery += vbCrLf & "RcvDNoteUserOID='" & Session("UserOID") & "',RcvDNoteDatetime=getdate()"
                    vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
                End If
            End If
        Next
    End Sub


    Protected Sub GrvListItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvListItem.SelectedIndexChanged

    End Sub

    Private Sub GrvListItem_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListItem.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "BRGCODE" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowItem As GridViewRow = GrvListItem.Rows(vnIdx)
            Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

            vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = DirectCast(vnGRowItem.Cells(ensColListItem.BRGCODE).Controls(0), LinkButton).Text
            vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowItem.Cells(ensColListItem.BRGNAME).Text
            vnGRowDetail.Cells(ensColDetail.BRGUNIT).Text = vnGRowItem.Cells(ensColListItem.BRGUNIT).Text

            psShowListItem(False)
        End If
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
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

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If HdfProcess.Value = "CancelRCV" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelRCV()
        ElseIf HdfProcess.Value = "vDelItem" Then
            psDeleteItem()
        ElseIf HdfProcess.Value = "ScanOpen" Then
            psScanOpen()
        ElseIf HdfProcess.Value = "ScanClosed" Then
            psScanClosed()
        ElseIf HdfProcess.Value = "CloseRCV" Then
            psCloseRCV()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub BtnScanOpen_Click(sender As Object, e As EventArgs) Handles BtnScanOpen.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Scan_Open) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Receiving " & TxtRcvNo.Text & " Scan Open ?"
        HdfProcess.Value = "ScanOpen"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub BtnScanClosed_Click(sender As Object, e As EventArgs) Handles BtnScanClosed.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Scan_Close) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Receiving " & TxtRcvNo.Text & " Scan Close ?"
        HdfProcess.Value = "ScanClosed"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub
    Private Sub psCloseRCV()
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
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvHeader_TR set TransStatus=" & enuTCSRCV.Closed & ",RcvCloseNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "ClosedUserOID=" & Session("UserOID") & ",ClosedDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusReceiving(TxtTransID.Text, enuTCSRCV.Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psScanOpen()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psScanOpen", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0")

            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvHeader_TR set TransStatus=" & enuTCSRCV.Scan_Open & ",ScanOpenUserOID=" & Session("UserOID") & ",ScanOpenDatetime=getdate() Where OID=" & TxtTransID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusReceiving...Start")
            pbuInsertStatusReceiving(TxtTransID.Text, enuTCSRCV.Scan_Open, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusReceiving...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Scan Start Sukses")
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

    Private Sub psScanClosed()
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

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvHeader_TR Set TransStatus=" & enuTCSRCV.Scan_Closed & ",ScanClosedUserOID=" & Session("UserOID") & ",ScanClosedDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusReceiving(TxtTransID.Text, enuTCSRCV.Scan_Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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
            Dim vnQuery As String
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Delete Sys_SsoRcvDetail_TR Where RcvHOID=" & TxtTransID.Text & " and OID=" & GrvDetail.Rows(HdfDetailRowIdx.Value).Cells(ensColDetail.OID).Text
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

    Private Sub psCancelRCV()
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
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoRcvHeader_TR set TransStatus=" & enuTCSRCV.Cancelled & ",RcvCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusReceiving(TxtTransID.Text, enuTCSRCV.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Protected Sub BtnCancelRcv_Click(sender As Object, e As EventArgs) Handles BtnCancelRcv.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan Receiving No. " & TxtRcvNo.Text & " ?<br />WARNING : Batal Receiving Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelRCV"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        TxtConfirmNote.Visible = True
        LblConfirmNote.Visible = True
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psShowPrOption(True)
    End Sub

    Private Sub BtnListItemFind_Click(sender As Object, e As EventArgs) Handles BtnListItemFind.Click
        psFillGrvListItem()
    End Sub

    Private Sub GrvListItem_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListItem.PageIndexChanging
        GrvListItem.PageIndex = e.NewPageIndex
        psFillGrvListItem()
    End Sub

    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Private Sub BtnCloseRcv_Click(sender As Object, e As EventArgs) Handles BtnCloseRcv.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Close Receiving No. " & TxtRcvNo.Text & " ?<br />WARNING : Close Receiving Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CloseRCV"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnLsScanClose_Click(sender As Object, e As EventArgs) Handles BtnLsScanClose.Click
        psShowLsScan(False)
    End Sub

    Protected Sub BtnProCancel_Click(sender As Object, e As EventArgs) Handles BtnProCancel.Click
        psShowPrOption(False)
    End Sub
    Private Sub psFillDstSOReport()
        Dim vnDtb As New DataTable
        vnDtb.Columns.Add("RptCode")
        vnDtb.Columns.Add("RptName")
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOTallyRcv, "Tally Penerimaan"})

        DstProReport.DataSource = vnDtb
        DstProReport.DataValueField = "RptCode"
        DstProReport.DataTextField = "RptName"
        DstProReport.DataBind()
    End Sub

    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        Dim vnCrpFileName As String = ""

        If DstProReport.SelectedValue = stuSsoReportType.RptSOTallyRcv Then
            psGenerateCrpTallyRcv(vnCrpFileName)
        Else
            Exit Sub
        End If

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

    Private Sub psGenerateCrpTallyRcv(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyRcv

        vbuCrpQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
        vbuCrpQuery += "       From fnTbl_SsoTallyRcv(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.RcvCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.RcvCompanyCode and mb.BRGCODE=ta.BRGCODE"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.RcvSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.RcvWarehouseOID"
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub

    Private Sub psFillGrvLsScan(vriBrgCode As String)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If ChkLsScanSt_DelNo.Checked = False And ChkLsScanSt_DelYes.Checked = False Then
            ChkLsScanSt_DelNo.Checked = True
            ChkLsScanSt_DelYes.Checked = True
        End If

        Dim vnCriteria As String = fbuFormatString(TxtLsScanDataFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select sc.RcvScanQty,sc.RcvScanNote,"
        vnQuery += vbCrLf & "       mu.UserName vRcvScanUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.RcvScanDatetime,106) + ' ' + convert(varchar(5),sc.RcvScanDatetime,108)vRcvScanTime,"
        vnQuery += vbCrLf & "	    case when abs(sc.RcvScanDeleted)=1 then 'Y' else 'N' end vRcvScanDeleted,"
        vnQuery += vbCrLf & "	    sc.RcvScanDeletedNote,"
        vnQuery += vbCrLf & "       du.UserID vRcvScanDeletedUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.RcvScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.RcvScanDeletedDatetime,108)vRcvScanDeletedTime"
        vnQuery += vbCrLf & "  From Sys_SsoRcvScan_TR sc"
        vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu on mu.OID=sc.RcvScanUserOID"
        vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA du on du.OID=sc.RcvScanDeletedUserOID"
        vnQuery += vbCrLf & " Where sc.RcvHOID=" & TxtTransID.Text & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
        vnQuery += vbCrLf & "       and (sc.RcvScanNote like '%" & vnCriteria & "%')"

        If Not (ChkLsScanSt_DelNo.Checked = True And ChkLsScanSt_DelYes.Checked = True) Then
            If ChkLsScanSt_DelNo.Checked = True Then
                vnQuery += vbCrLf & "       and abs(RcvScanDeleted)=0"
            Else
                vnQuery += vbCrLf & "       and abs(RcvScanDeleted)=1"
            End If
        End If

        vnQuery += vbCrLf & " Order by sc.OID"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvLsScan.DataSource = vnDtb
        GrvLsScan.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        If ChkLsScanSt_DelYes.Checked = True Then
            Dim vn As Integer
            For vn = 0 To GrvLsScan.Rows.Count - 1
                If GrvLsScan.Rows(vn).Cells(ensColLsScan.vRcvScanDeleted).Text = "Y" Then
                    GrvLsScan.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Protected Sub BtnLsScanDataFind_Click(sender As Object, e As EventArgs) Handles BtnLsScanDataFind.Click
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
    End Sub
    Private Sub GrvLsScan_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsScan.PageIndexChanging
        GrvLsScan.PageIndex = e.NewPageIndex
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Val(TxtTransID.Text) = 0 Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvDetail(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub DstCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstCompany.SelectedIndexChanged
        If BtnBaru.Visible Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If Session("UserWarehouseCode") = "" Then
            pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vnSQLConn)
        Else
            pbuFillDstSubWarehouse_ByCompanyCode_ByUserOID(DstSubWhs, False, DstCompany.SelectedValue, Session("UserOID"), vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class