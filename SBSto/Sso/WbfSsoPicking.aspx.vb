Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoPicking
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoPicking"
    Const csTNoPrefix = "PC"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
    End Enum

    Enum ensColDetail
        vNo = 0
        OID = 1
        BRGCODE = 2
        BRGNAME = 3
        BRGUNIT = 4
        vIsSN = 5
        vPickDQtyTotal = 6
        vSumPickScanQty = 7
        vPickScanVarian = 8
        vPickDNote = 9
        TxtvPickDNote = 10
        vPickDNoteBy = 11
        vPickDNoteDatetime = 12
    End Enum
    Enum ensColListDoc
        CompanyCode = 0
        no_nota = 1
        vtanggal = 2
        kode_cust = 3
        CUSTOMER = 4
        ALAMAT = 5
        kota = 6
    End Enum

    Enum ensColListTRB
        OID = 0
        CompanyCode = 1
        NoBukti = 2
        vTanggal = 3
        GudangAsal = 4
        GudangTujuan = 5
    End Enum

    Enum ensColLsScan
        vStorageInfoHtml = 0
        PickScanQty = 1
        vPickScanNoteSN = 2
        vPickScanUser = 3
        vPickScanTime = 4
        vPickScanDeleted = 5
        PickScanDeletedNote = 6
        vPickScanDeletedUser = 7
        vPickScanDeletedTime = 8
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPickDate.Text = ""
        TxtPickNo.Text = ""
        TxtPickNote.Text = ""

        TxtPickDate.Text = ""

        HdfCompanyCode.Value = ""

        HdfCustCode.Value = ""
        HdfCustName.Value = ""

        TxtPickTujuan.Text = ""

        HdfSchDTypeOID.Value = "5"

        HdfPickRefOID.Value = "0"
        TxtPickRefNo.Text = ""
        HdfPickRefDate.Value = ""

        HdfTransStatus.Value = enuTCSPCK.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsScan.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsScan.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrOption.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListDoc.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListDoc.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListTRB.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPicking, vnSQLConn)

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
            vnCrStatus += enuTCSPCK.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCSPCK.Cancelled & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCSPCK.Closed & ","
        End If
        If ChkSt_ScanOpen.Checked = True Then
            vnCrStatus += enuTCSPCK.Scan_Open & ","
        End If
        If ChkSt_ScanClosed.Checked = True Then
            vnCrStatus += enuTCSPCK.Scan_Closed & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.PickNo,PM.PickRefNo,SM.SchDTypeName,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PickDate,106)vPickDate,"

        vnQuery += vbCrLf & "     PM.PickCompanyCode,"

        vnQuery += vbCrLf & "     WMA.WarehouseName vWarehouseNameAsal,SWA.SubWhsName vSubWhsNameAsal,"
        vnQuery += vbCrLf & "     WMD.WarehouseName vWarehouseNameTujuan,"

        vnQuery += vbCrLf & "     case when PM.SchDTypeOID = " & enuSchDType.TRB & " then SWD.SubWhsName else PM.PickCustName end vTujuan,"

        vnQuery += vbCrLf & "     PM.PickNote,PM.PickCloseNote,PM.PickCancelNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanOpenDatetime,106)+' '+convert(varchar(5),PM.ScanOpenDatetime,108)+' '+ PR.UserName vScanOpen,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanClosedDatetime,106)+' '+convert(varchar(5),PM.ScanClosedDatetime,108)+' '+ AP.UserName vScanClosed,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ CL.UserName vClosed"

        vnQuery += vbCrLf & "From Sys_SsoPickHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBDcm() & "Sys_DcmSchDType_MA SM with(nolock) on SM.OID=PM.SchDTypeOID"

        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WMA with(nolock) on WMA.OID=PM.PickWarehouseAsalOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SWA with(nolock) on SWA.OID=PM.PickSubWarehouseAsalOID"

        vnQuery += vbCrLf & "     left outer join " & fbuGetDBMaster() & "Sys_Warehouse_MA WMD with(nolock) on WMD.OID=PM.PickWarehouseTujuanOID"
        vnQuery += vbCrLf & "     left outer join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SWD with(nolock) on SWD.OID=PM.PickSubWarehouseTujuanOID"

        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.ScanOpenUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA AP with(nolock) on AP.OID=PM.ScanClosedUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CL with(nolock) on CL.OID=PM.ClosedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.PickCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.PickWarehouseAsalOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.PickNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.PickDate AS DATE) >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.PickDate AS DATE) <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.PickNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
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

        vnQuery = "Select sd.vStorageInfoHtml,sc.PickScanQty,"
        vnQuery += vbCrLf & "       case when isnull(sc.PickScanNote,'')='' then sc.PickScanSerialNo else sc.PickScanNote end vPickScanNoteSN,"
        vnQuery += vbCrLf & "       mu.UserName vPickScanUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.PickScanDatetime,106) + ' ' + convert(varchar(5),sc.PickScanDatetime,108)vPickScanTime,"
        vnQuery += vbCrLf & "	    case when abs(sc.PickScanDeleted)=1 then 'Y' else 'N' end vPickScanDeleted,"
        vnQuery += vbCrLf & "	    sc.PickScanDeletedNote,"
        vnQuery += vbCrLf & "       du.UserID vPickScanDeletedUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.PickScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.PickScanDeletedDatetime,108)vPickScanDeletedTime"
        vnQuery += vbCrLf & "  From Sys_SsoPickScan_TR sc"
        vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('')sd on sd.vStorageOID=sc.StorageOID"
        vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu on mu.OID=sc.PickScanUserOID"
        vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA du on du.OID=sc.PickScanDeletedUserOID"
        vnQuery += vbCrLf & " Where sc.PickHOID=" & TxtTransID.Text & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
        vnQuery += vbCrLf & "       and (sc.PickScanNote like '%" & vnCriteria & "%' OR sc.PickScanSerialNo like '%" & vnCriteria & "%')"

        If Not (ChkLsScanSt_DelNo.Checked = True And ChkLsScanSt_DelYes.Checked = True) Then
            If ChkLsScanSt_DelNo.Checked = True Then
                vnQuery += vbCrLf & "       and abs(PickScanDeleted)=0"
            Else
                vnQuery += vbCrLf & "       and abs(PickScanDeleted)=1"
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
                If GrvLsScan.Rows(vn).Cells(ensColLsScan.vPickScanDeleted).Text = "Y" Then
                    GrvLsScan.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Private Sub psFillGrvDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,''vIsSN,0 vPickDQtyTotal,0 vSumPickScanQty,0 vSOStockScanVarian,"
            vnQuery += vbCrLf & "       ''vPickDNote,''vPickDNoteBy,Null vPickDNoteDatetime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()
        Else
            vnQuery = "Select Row_Number()over(order by mb.BRGNAME)vNo,d.OID,"
            vnQuery += vbCrLf & "       d.BRGCODE,mb.BRGNAME,mb.BRGUNIT,"
            vnQuery += vbCrLf & "       case when abs(mb.IsSN)=0 then 'N' else 'Y' end vIsSN,"
            vnQuery += vbCrLf & "       d.vPickDQtyTotal,d.vSumPickScanQty,d.vPickScanVarian,"
            vnQuery += vbCrLf & "       d.vPickDNote,d.vPickDNoteBy,d.vPickDNoteDatetime"
            vnQuery += vbCrLf & "  From fnTbl_SsoPickScan(" & vriHOID & ")d"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb on mb.BRGCODE=d.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "'"

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " Where d.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%'"
            End If
            If ChkFindVarian.Checked Then
                vnQuery += vbCrLf & " and d.vPickScanVarian<>0"
            End If

            vnQuery += vbCrLf & "Order by mb.BRGNAME"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            Dim vn As Integer
            If HdfActionStatus.Value = cbuActionNorm Then
                GrvDetail.Columns(ensColDetail.vPickDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vPickDNote).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.TxtvPickDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.TxtvPickDNote).ItemStyle.CssClass = "myDisplayNone"
            Else
                GrvDetail.Columns(ensColDetail.vPickDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vPickDNote).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtvPickDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtvPickDNote).ItemStyle.CssClass = ""
            End If

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()

            Dim vnGRow As GridViewRow
            If HdfActionStatus.Value = cbuActionEdit Then
                Dim vnTxtvPickDNote As TextBox

                For vn = 0 To GrvDetail.Rows.Count - 1
                    vnGRow = GrvDetail.Rows(vn)
                    vnTxtvPickDNote = vnGRow.FindControl("TxtvPickDNote")

                    vnTxtvPickDNote.Text = Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vPickDNote).Text), "<br />", Chr(10))
                Next
            End If

            If HdfTransStatus.Value = enuTCSPCK.Scan_Open Or HdfTransStatus.Value = enuTCSPCK.Scan_Closed Then

            End If
        End If
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
            psButtonShowList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
            psButtonStatus()
        End If
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnCancelPick.Enabled = False
        BtnScanOpen.Enabled = False
        BtnScanClosed.Enabled = False
        BtnClosePick.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True
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
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(PickNo,len(PickNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(PickNo,len(PickNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoPickHeader_TR with(nolock)"
        vnQuery += vbCrLf & "      Where PickNo like '" & vnTNoPrefix & "+'/%'"
        TxtPickNo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
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
        psFillGrvDetail(0, vnSQLConn)

        TxtPickDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgPickNo.Text = ""
        LblMsgPickDate.Text = ""
        LblMsgCompany.Text = ""
        LblMsgSubWhs.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPickDate.Enabled = vriBo
        TxtPickNote.ReadOnly = Not vriBo

        If HdfActionStatus.Value = cbuActionNew Then
            DstCompany.Enabled = vriBo
            DstSubWhs.Enabled = vriBo
        Else
            If HdfActionStatus.Value = cbuActionEdit Then
                DstCompany.Enabled = False
                DstSubWhs.Enabled = False
            Else
                DstCompany.Enabled = True
                DstSubWhs.Enabled = True
            End If
        End If

        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelPick.Visible = Not vriBo
        BtnScanOpen.Visible = Not vriBo
        BtnScanClosed.Visible = Not vriBo
        BtnClosePick.Visible = Not vriBo

        BtnPickRefNo.Visible = vriBo

        BtnPreview.Visible = Not vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "vSumPickScanQty" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx
            HdfLsScanBrgCode.Value = vnGRow.Cells(ensColDetail.BRGCODE).Text
            psFillGrvLsScan(HdfLsScanBrgCode.Value)
            LblLsScanTitle.Text = "SCAN " & vnGRow.Cells(ensColDetail.BRGCODE).Text & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text

            If vnGRow.Cells(ensColDetail.vIsSN).Text = "Y" Then
                GrvLsScan.Columns(ensColLsScan.vPickScanNoteSN).HeaderText = "Serial No"
            Else
                GrvLsScan.Columns(ensColLsScan.vPickScanNoteSN).HeaderText = "Note"
            End If
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

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoPicking
            vnParam += "&vqTrNo=" & TxtPickNo.Text

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

        HdfActionStatus.Value = cbuActionNorm

        psEnableInput(False)
        psEnableSave(False)
        psButtonVisible()

        If TxtTransID.Text = "" Then
            psClearData()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            GrvDetail.PagerSettings.Visible = True

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

            GrvDetail.PagerSettings.Visible = True
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

        vnQuery = "Select PM.*,convert(varchar(11),PM.PickDate,106)vPickDate,convert(varchar(11),PM.PickRefDate,106)vPickRefDate,"
        vnQuery += vbCrLf & "     case when PM.SchDTypeOID = " & enuSchDType.TRB & " then SWD.SubWhsName else PM.PickCustName end vTujuan,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoPickHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     left outer join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SWD with(nolock) on SWD.OID=PM.PickSubWarehouseTujuanOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoPicking & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtPickDate.Text = vnDtb.Rows(0).Item("vPickDate")
            TxtPickNo.Text = vnDtb.Rows(0).Item("PickNo")
            TxtPickNote.Text = vnDtb.Rows(0).Item("PickNote")

            HdfCompanyCode.Value = Trim(vnDtb.Rows(0).Item("PickCompanyCode"))
            DstCompany.SelectedValue = HdfCompanyCode.Value

            pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vriSQLConn)
            DstSubWhs.SelectedValue = vnDtb.Rows(0).Item("PickSubWarehouseAsalOID")

            HdfCustCode.Value = vnDtb.Rows(0).Item("PickCustCode")
            HdfCustName.Value = vnDtb.Rows(0).Item("PickCustName")

            TxtPickTujuan.Text = vnDtb.Rows(0).Item("vTujuan")

            HdfSchDTypeOID.Value = vnDtb.Rows(0).Item("SchDTypeOID")

            HdfPickRefOID.Value = vnDtb.Rows(0).Item("PickRefOID")
            TxtPickRefNo.Text = vnDtb.Rows(0).Item("PickRefNo")
            HdfPickRefDate.Value = vnDtb.Rows(0).Item("vPickRefDate")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        GrvDetail.PageIndex = 0
        psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelPick.Visible = BtnCancelPick.Enabled
        BtnScanOpen.Visible = BtnScanOpen.Enabled
        BtnScanClosed.Visible = BtnScanClosed.Enabled
        BtnClosePick.Visible = BtnClosePick.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelPick.Enabled = False
        BtnClosePick.Enabled = False
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
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCSPCK.Baru Or HdfTransStatus.Value = enuTCSPCK.Scan_Open Or HdfTransStatus.Value = enuTCSPCK.Scan_Closed)

            BtnCancelPick.Enabled = (HdfTransStatus.Value = enuTCSPCK.Baru)

            BtnScanOpen.Enabled = (HdfTransStatus.Value = enuTCSPCK.Baru Or HdfTransStatus.Value = enuTCSPCK.Scan_Closed)
            BtnScanClosed.Enabled = (HdfTransStatus.Value = enuTCSPCK.Scan_Open)
            BtnClosePick.Enabled = (HdfTransStatus.Value = enuTCSPCK.Scan_Closed)

            BtnPreview.Enabled = (HdfTransStatus.Value >= enuTCSPCK.Closed)

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
        GrvDetail.PagerSettings.Visible = False
        psFillGrvDetail(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
        BtnPickRefNo.Visible = False
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

        If Not IsDate(TxtPickDate.Text) Then
            LblMsgPickDate.Text = "Isi Tanggal Picking dengan benar"
            vnSave = False
        End If
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "" Then
            LblMsgSubWhs.Text = "Pilih Gudang"
            vnSave = False
        End If
        If HdfSchDTypeOID.Value = enuSchDType.TRB Then
            If DstSubWhs.SelectedValue <> HdfSubWhsAsalOID.Value Then
                LblMsgSubWhs.Text = "Warehouse TRB<>Warehouse Picking...Pilih TRB Kembali"
                vnSave = False
            End If
        End If
        If Not vnSave Then Exit Sub

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psSaveBaru", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True

                vsTextStream.WriteLine("Error Open Koneksi SQLServer :")
                vsTextStream.WriteLine(pbMsgError)
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("------------------------EOF------------------------")

                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            Dim vnCompanyCode As String = DstCompany.SelectedValue

            Dim vnQuery As String

            vnQuery = "Select count(1) From Sys_SsoPickHeader_TR Where PickNo='" & Trim(TxtPickNo.Text) & "'"
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgPickNo.Text = "No.Picking " & Trim(TxtPickNo.Text) & " Sudah pernah dipakai."

                vsTextStream.WriteLine(LblMsgPickNo.Text)
                vsTextStream.WriteLine(pbMsgError)
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("------------------------EOF------------------------")

                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                vnSave = False
            End If

            If Not vnSave Then
                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Exit Sub
            End If

            Dim vnSubWhsOID_Asal As String = DstSubWhs.SelectedValue
            Dim vnWarehouseOID_Asal As String = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID_Asal, vnSQLConn)

            Dim vnSubWhsOID_Tujuan As String
            Dim vnWarehouseOID_Tujuan As String

            If HdfSchDTypeOID.Value = enuSchDType.TRB Then
                vnSubWhsOID_Tujuan = HdfSubWhsTujuanOID.Value
                vnWarehouseOID_Tujuan = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID_Asal, vnSQLConn)
            Else
                vnSubWhsOID_Tujuan = "Null"
                vnWarehouseOID_Tujuan = "Null"
            End If

            Dim vnRefNo As String = fbuFormatString(Trim(TxtPickRefNo.Text))

            Dim vnOID As Integer
            vnQuery = "Select max(OID) from Sys_SsoPickHeader_TR"
            vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            psSetTransNo(Trim(HdfCompanyCode.Value), DstSubWhs.SelectedValue, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoPickHeader_TR(OID,"
            vnQuery += vbCrLf & "SchDTypeOID,PickRefOID,"
            vnQuery += vbCrLf & "PickRefNo,PickRefDate,"
            vnQuery += vbCrLf & "PickNo,PickDate,"
            vnQuery += vbCrLf & "PickCompanyCode,"

            vnQuery += vbCrLf & "PickWarehouseAsalOID,PickSubWarehouseAsalOID,"
            vnQuery += vbCrLf & "PickWarehouseTujuanOID,PickSubWarehouseTujuanOID,"

            vnQuery += vbCrLf & "PickCustCode,PickCustName,"
            vnQuery += vbCrLf & "PickNote,"
            vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
            vnQuery += vbCrLf & "values(" & vnOID & ","
            vnQuery += vbCrLf & "" & HdfSchDTypeOID.Value & "," & HdfPickRefOID.Value & ","
            vnQuery += vbCrLf & "'" & vnRefNo & "','" & HdfPickRefDate.Value & "',"
            vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPickNo.Text)) & "','" & TxtPickDate.Text & "',"
            vnQuery += vbCrLf & "'" & HdfCompanyCode.Value & "',"
            vnQuery += vbCrLf & vnWarehouseOID_Asal & "," & vnSubWhsOID_Asal & ","
            vnQuery += vbCrLf & vnWarehouseOID_Tujuan & "," & vnSubWhsOID_Tujuan & ","

            vnQuery += vbCrLf & "'" & fbuFormatString(HdfCustCode.Value) & "','" & fbuFormatString(HdfCustName.Value) & "',"

            vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPickNote.Text)) & "',"
            vnQuery += vbCrLf & "'" & stuTransCode.SsoPicking & "'," & Session("UserOID") & ",getdate())"
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vnQuery = "Insert into Sys_SsoPickDetail_TR"
            vnQuery += vbCrLf & "(PickHOID,BRGCODE,BRGNAME,BRGUNIT,PickDQty,PickDQtyBonus)"
            If HdfSchDTypeOID.Value = enuSchDType.Invoice Then
                vnQuery += vbCrLf & "Select " & vnOID & ",KODE_BARANG,NAMA_BARANG,SATUAN,sum(QTY),sum(QTYBONUS)"
                vnQuery += vbCrLf & "  From " & fbuGetDBDcm() & "Sys_DcmJUAL Where CompanyCode='" & HdfCompanyCode.Value & "' and NO_NOTA='" & vnRefNo & "'"
                vnQuery += vbCrLf & " Group by KODE_BARANG,NAMA_BARANG,SATUAN"
            ElseIf HdfSchDTypeOID.Value = enuSchDType.TRB Then
                vnQuery += vbCrLf & "Select " & vnOID & ",KodeBrg,NamaBrg,Satuan,sum(Qty),0"
                vnQuery += vbCrLf & "  From " & fbuGetDBDcm() & "Sys_DcmTRBDetail_TR Where TRBHOID=" & HdfPickRefOID.Value
                vnQuery += vbCrLf & " Group by KodeBrg,NamaBrg,Satuan"
            End If
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusPicking...Start")
            pbuInsertStatusPicking(vnOID, enuTCSPCK.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPicking...End")
            vsTextStream.WriteLine("")

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            TxtTransID.Text = vnOID

            HdfTransStatus.Value = enuTCSPCK.Baru

            psEnableInput(False)
            psEnableSave(False)

            HdfActionStatus.Value = cbuActionNorm
            GrvDetail.PagerSettings.Visible = True
            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("Create Picking Sukses")
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("PROCESS TERMINATED...ERROR :")
            vsTextStream.WriteLine(ex.Message)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------ERROR-----------------------")
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("Create Picking Error")
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

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

        If Trim(TxtPickNo.Text) = "" Then
            LblMsgPickNo.Text = "Isi Nomor SO"
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

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            If (HdfTransStatus.Value = enuTCSPCK.Baru) Then
                vnQuery = "Update Sys_SsoPickHeader_TR set "
                vnQuery += vbCrLf & "PickDate='" & TxtPickDate.Text & "',"
                vnQuery += vbCrLf & "PickNote='" & fbuFormatString(Trim(TxtPickNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

            pbuInsertStatusPicking(TxtTransID.Text, enuTCSPCK.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psSaveDetail(vriOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnTxtvPickDNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtvPickDNote = vnGRow.FindControl("TxtvPickDNote")
            If Trim(vnTxtvPickDNote.Text) <> Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vPickDNote).Text), "<br />", vbLf) Then
                vnQuery = "Update Sys_SsoPickDetail_TR set "
                vnQuery += vbCrLf & "PickDNote='" & fbuFormatString(vnTxtvPickDNote.Text) & "',"
                vnQuery += vbCrLf & "PickDNoteUserOID='" & Session("UserOID") & "',PickDNoteDatetime=getdate()"
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
        If HdfProcess.Value = "CancelPick" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelPick()
        ElseIf HdfProcess.Value = "ScanOpen" Then
            psScanOpen()
        ElseIf HdfProcess.Value = "ScanClosed" Then
            psScanClosed()
        ElseIf HdfProcess.Value = "ClosePick" Then
            psClosePick()
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
        LblConfirmMessage.Text = "SO " & TxtPickNo.Text & " Scan Open ?"
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
        LblConfirmMessage.Text = "SO " & TxtPickNo.Text & " Scan Close ?"
        HdfProcess.Value = "ScanClosed"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub
    Private Sub psCancelPick()
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

            vnQuery = "Update Sys_SsoPickHeader_TR set TransStatus=" & enuTCSPCK.Cancelled & ",PickCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusPicking(TxtTransID.Text, enuTCSPCK.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psClosePick()
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

            vnQuery = "Update Sys_SsoPickHeader_TR set TransStatus=" & enuTCSPCK.Closed & ",PickCloseNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "ClosedUserOID=" & Session("UserOID") & ",ClosedDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusPicking(TxtTransID.Text, enuTCSPCK.Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

            vnQuery = "Update Sys_SsoPickHeader_TR set TransStatus=" & enuTCSPCK.Scan_Open & ",ScanOpenUserOID=" & Session("UserOID") & ",ScanOpenDatetime=getdate() Where OID=" & TxtTransID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPicking...Start")
            pbuInsertStatusPicking(TxtTransID.Text, enuTCSPCK.Scan_Open, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPicking...End")

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

            vnQuery = "Update Sys_SsoPickHeader_TR Set TransStatus=" & enuTCSPCK.Scan_Closed & ",ScanClosedUserOID=" & Session("UserOID") & ",ScanClosedDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusPicking(TxtTransID.Text, enuTCSPCK.Scan_Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Protected Sub BtnCancelPick_Click(sender As Object, e As EventArgs) Handles BtnCancelPick.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan Picking No. " & TxtPickNo.Text & " ?<br />WARNING : Batal Picking Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelPick"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
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

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Private Sub BtnClosePick_Click(sender As Object, e As EventArgs) Handles BtnClosePick.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Close Picking No. " & TxtPickNo.Text & " ?<br />WARNING : Close Picking Tidak Dapat Dibatalkan"
        HdfProcess.Value = "ClosePick"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub GrvDetail_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvDetail.SelectedIndexChanged

    End Sub

    Private Sub GrvDetail_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvDetail.PageIndexChanging
        GrvDetail.PageIndex = e.NewPageIndex
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
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOTallyPick, "Tally Picking"})

        DstProReport.DataSource = vnDtb
        DstProReport.DataValueField = "RptCode"
        DstProReport.DataTextField = "RptName"
        DstProReport.DataBind()
    End Sub

    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
        Dim vnCrpFileName As String = ""

        If DstProReport.SelectedValue = stuSsoReportType.RptSOTallyPick Then
            psGenerateCrpTallyPick(vnCrpFileName)
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

    Private Sub psGenerateCrpTallyPick(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyPick

        vbuCrpQuery = "Select ta.*,mc.CompanyName,"
        vbuCrpQuery += "      wha.WarehouseName vWarehouseNameAsal,swa.SubWhsCode vSubWhsCodeAsal,swa.SubWhsName vSubWhsNameAsal,"
        vbuCrpQuery += "      whd.WarehouseName vWarehouseNameTujuan,swd.SubWhsCode vSubWhsCodeTujuan,swd.SubWhsName vSubWhsNameTujuan,"
        vbuCrpQuery += "      row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
        vbuCrpQuery += " From fnTbl_SsoTallyPick(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "      inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.PickCompanyCode"
        vbuCrpQuery += "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.PickCompanyCode and mb.BRGCODE=ta.BRGCODE"

        vbuCrpQuery += "      inner join " & vnDBMaster & "Sys_SubWarehouse_MA swa with(nolock) on swa.OID=ta.PickSubWarehouseAsalOID"
        vbuCrpQuery += "      inner join " & vnDBMaster & "Sys_Warehouse_MA wha with(nolock) on wha.OID=ta.PickWarehouseAsalOID"

        vbuCrpQuery += "      left outer join " & vnDBMaster & "Sys_SubWarehouse_MA swd with(nolock) on swd.OID=ta.PickSubWarehouseAsalOID"
        vbuCrpQuery += "      left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=ta.PickWarehouseAsalOID"
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub

    Private Sub psShowListDoc(vriBo As Boolean)
        If vriBo Then
            DivListDoc.Style(HtmlTextWriterStyle.Visibility) = "visible"
            HdfOnList.Value = enuSchDType.Invoice

            psShowListTRB(False)
        Else
            DivListDoc.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub psShowListTRB(vriBo As Boolean)
        If vriBo Then
            DivListTRB.Style(HtmlTextWriterStyle.Visibility) = "visible"
            HdfOnList.Value = enuSchDType.TRB

            psShowListDoc(False)
        Else
            DivListTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnListDocClose_Click(sender As Object, e As EventArgs) Handles BtnListDocClose.Click
        psShowListDoc(False)
    End Sub

    Private Sub BtnListTRBClose_Click(sender As Object, e As EventArgs) Handles BtnListTRBClose.Click
        psShowListTRB(False)
    End Sub

    Protected Sub BtnListTRBDoc_Click(sender As Object, e As EventArgs) Handles BtnListTRBDoc.Click
        psShowListDoc(True)
    End Sub

    Protected Sub BtnListDocTRB_Click(sender As Object, e As EventArgs) Handles BtnListDocTRB.Click
        psShowListTRB(True)
    End Sub

    Private Sub psFillGrvListDoc(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String = ""

        vnCriteria = "      Where 1=1"
        If DstCompany.SelectedIndex > 0 Then
            vnCriteria += vbCrLf & "            and CompanyCode='" & DstCompany.SelectedValue & "'"
        End If
        If Trim(TxtListDocCustomer.Text) <> "" Then
            vnCriteria += vbCrLf & "            and CUSTOMER like '%" & fbuFormatString(Trim(TxtListDocCustomer.Text)) & "%'"
        End If
        If Trim(TxtListDocNota.Text) <> "" Then
            vnCriteria += vbCrLf & "            and no_nota like '%" & fbuFormatString(Trim(TxtListDocNota.Text)) & "%'"
        End If
        If IsDate(TxtListDocStart.Text) Then
            vnCriteria += vbCrLf & "            and tanggal >= '" & TxtListDocStart.Text & "'"
        End If
        If IsDate(TxtListDocEnd.Text) Then
            vnCriteria += vbCrLf & "            and tanggal <= '" & TxtListDocEnd.Text & "'"
        End If

        vnQuery = "Select Distinct CompanyCode,no_nota,convert(varchar(11),tanggal,106)vtanggal,kode_cust,CUSTOMER,ALAMAT,kota"
        vnQuery += vbCrLf & "       From " & fbuGetDBDcm() & "Sys_DcmJUAL"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by CompanyCode,no_nota"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListDoc.DataSource = vnDtb
        GrvListDoc.DataBind()
    End Sub

    Private Sub psFillGrvListTRB(vriSQLConn As SqlConnection)
        HdfSubWhsAsalOID.Value = DstSubWhs.SelectedValue
        Dim vnWarehouseOID As String
        vnWarehouseOID = fbuGetWarehouseOID_BySubWhsOID(HdfSubWhsAsalOID.Value, vriSQLConn)

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.OID,PM.CompanyCode,PM.NoBukti,convert(varchar(11),PM.Tanggal,106)vTanggal,PM.GudangAsal,PM.GudangTujuan"
        vnQuery += vbCrLf & "       From " & fbuGetDBDcm() & "Sys_DcmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "Where PM.GudangAsalOID=" & vnWarehouseOID

        If DstCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.CompanyCode='" & DstCompany.SelectedValue & "'"
        End If
        If Trim(TxtListTRBNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and PM.NoBukti like '%" & fbuFormatString(Trim(TxtListTRBNo.Text)) & "%'"
        End If

        vnQuery += vbCrLf & "Order by PM.CompanyCode,PM.NoBukti"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListTRB.DataSource = vnDtb
        GrvListTRB.DataBind()
    End Sub

    Private Sub GrvListTRB_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListTRB.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "NoBukti" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowList As GridViewRow = GrvListTRB.Rows(vnIdx)

            Dim vnCompanyCode As String = vnGRowList.Cells(ensColListTRB.CompanyCode).Text
            Dim vnGudangAsal As String = vnGRowList.Cells(ensColListTRB.GudangAsal).Text

            HdfSchDTypeOID.Value = enuSchDType.TRB

            HdfCompanyCode.Value = vnCompanyCode

            Dim vnSubWhsTujuan As String = vnGRowList.Cells(ensColListTRB.GudangTujuan).Text
            HdfSubWhsTujuanOID.Value = fbuGetSubWhOID_BySubWhsName(vnSubWhsTujuan, vnSQLConn)

            TxtPickTujuan.Text = vnSubWhsTujuan

            HdfPickRefOID.Value = vnGRowList.Cells(ensColListTRB.OID).Text
            TxtPickRefNo.Text = DirectCast(vnGRowList.Cells(ensColListTRB.NoBukti).Controls(0), LinkButton).Text

            HdfPickRefDate.Value = vnGRowList.Cells(ensColListTRB.vTanggal).Text

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowListTRB(False)
        End If
    End Sub

    Private Sub GrvListDoc_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListDoc.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "no_nota" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowList As GridViewRow = GrvListDoc.Rows(vnIdx)

            Dim vnCompanyCode As String = vnGRowList.Cells(ensColListDoc.CompanyCode).Text
            Dim vnNotaNo As String = DirectCast(vnGRowList.Cells(ensColListDoc.no_nota).Controls(0), LinkButton).Text
            Dim vnRetHOID As String = "0"

            HdfSchDTypeOID.Value = enuSchDType.Invoice

            HdfCompanyCode.Value = vnCompanyCode
            HdfCustCode.Value = vnGRowList.Cells(ensColListDoc.kode_cust).Text
            HdfCustName.Value = vnGRowList.Cells(ensColListDoc.CUSTOMER).Text
            TxtPickTujuan.Text = HdfCustName.Value

            HdfPickRefOID.Value = vnRetHOID
            TxtPickRefNo.Text = vnNotaNo

            HdfPickRefDate.Value = vnGRowList.Cells(ensColListDoc.vtanggal).Text

            psShowListDoc(False)
        End If
    End Sub

    Private Sub GrvListDoc_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListDoc.PageIndexChanging
        GrvListDoc.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListDoc(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListTRB_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListTRB.PageIndexChanging
        GrvListTRB.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListTRB(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListDocFind_Click(sender As Object, e As EventArgs) Handles BtnListDocFind.Click
        If Trim(TxtListDocCustomer.Text) = "" And Trim(TxtListDocNota.Text) = "" And IsDate(TxtListDocStart.Text) = False And IsDate(TxtListDocEnd.Text) = False Then
            LblMsgListDoc.Text = "Pilih Nomor Nota, Customer atau Periode Tanggal Nota"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListDoc(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListTRBFind_Click(sender As Object, e As EventArgs) Handles BtnListTRBFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListTRB(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvListTRB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvListTRB.SelectedIndexChanged

    End Sub

    Protected Sub BtnPickRefNo_Click(sender As Object, e As EventArgs) Handles BtnPickRefNo.Click
        If HdfOnList.Value = enuSchDType.Invoice Then
            psShowListDoc(True)
        ElseIf HdfOnList.Value = enuSchDType.TRB Then
            psShowListTRB(True)
        End If
    End Sub

    Protected Sub BtnLsScanDataFind_Click(sender As Object, e As EventArgs) Handles BtnLsScanDataFind.Click
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
    End Sub
    Private Sub GrvLsScan_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsScan.PageIndexChanging
        GrvLsScan.PageIndex = e.NewPageIndex
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
    End Sub

    Private Sub DstCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstCompany.SelectedIndexChanged
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