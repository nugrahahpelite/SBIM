Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoSO
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSO"
    Const csTNoPrefix = "SO"

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

    Enum ensColDetail
        vNo = 0
        OID = 1
        BRGCODE = 2
        BRGNAME = 3
        BRGUNIT = 4
        SOStockQty = 5
        vSumSOScanQty = 6
        vSOStockScanVarian = 7
        vSOStockNote = 8
        TxtvSOStockNote = 9
        vSOStockNoteBy = 10
        vSOStockNoteDatetime = 11
    End Enum

    Enum ensColTaDetail
        vSOScanDeleted = 11
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtSODate.Text = ""
        TxtSONo.Text = ""
        TxtSONote.Text = ""

        HdfTransStatus.Value = enuTCSSOH.Baru
    End Sub
    Enum ensColLsScan
        vSOScanDeleted = 5
    End Enum
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoSSOH, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            pbuFillDstHour(DstCutOffHour)
            pbuFillDstMinute(DstCutOffMin)
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
            vnCrStatus += enuTCSSOH.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCSSOH.Cancelled & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCSSOH.Closed & ","
        End If
        If ChkSt_ScanOpen.Checked = True Then
            vnCrStatus += enuTCSSOH.Scan_Open & ","
        End If
        If ChkSt_ScanClosed.Checked = True Then
            vnCrStatus += enuTCSSOH.Scan_Closed & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.SONo,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.SOCutOff,106) + ' '+ convert(varchar(5),PM.SOCutOff,108)vSOCutOff,"
        vnQuery += vbCrLf & "     PM.SOCompanyCode,WM.WarehouseName,SW.SubWhsName,PM.SONote,PM.SOCloseNote,PM.SOCancelNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanOpenDatetime,106)+' '+convert(varchar(5),PM.ScanOpenDatetime,108)+' '+ PR.UserName vScanOpen,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanClosedDatetime,106)+' '+convert(varchar(5),PM.ScanClosedDatetime,108)+' '+ AP.UserName vScanClosed,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ CL.UserName vClosed"

        vnQuery += vbCrLf & "From Sys_SsoSOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.SOWarehouseOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SW with(nolock) on SW.OID=PM.SOSubWarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.ScanOpenUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA AP with(nolock) on AP.OID=PM.ScanClosedUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CL with(nolock) on CL.OID=PM.ClosedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.SOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.SOWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where PM.SOTypeOID=" & enuSOType.WinAcc

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.SONo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.SOCutOff AS DATE) >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.SOCutOff AS DATE) <= '" & TxtListEnd.Text & "'"
        End If
        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.SOCompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.SOWarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "Order by PM.SONo"
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

        vnQuery = "Select sd.vStorageInfoHtml,sc.SOScanQty,sc.SOScanNote,"
        vnQuery += vbCrLf & "       mu.UserName vSOScanUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.SOScanDatetime,106) + ' ' + convert(varchar(5),sc.SOScanDatetime,108)vSOScanTime,"
        vnQuery += vbCrLf & "	    case when abs(sc.SOScanDeleted)=1 then 'Y' else 'N' end vSOScanDeleted,"
        vnQuery += vbCrLf & "	    sc.SOScanDeletedNote,"
        vnQuery += vbCrLf & "       du.UserID vSOScanDeletedUser,"
        vnQuery += vbCrLf & "	    convert(varchar(11),sc.SOScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.SOScanDeletedDatetime,108)vSOScanDeletedTime"
        vnQuery += vbCrLf & "  From Sys_SsoSOScan_TR sc"
        vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('')sd on sd.vStorageOID=sc.StorageOID"
        vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu on mu.OID=sc.SOScanUserOID"
        vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA du on du.OID=sc.SOScanDeletedUserOID"
        vnQuery += vbCrLf & " Where sc.SOHOID=" & TxtTransID.Text & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
        vnQuery += vbCrLf & "       and (sc.SOScanNote like '%" & vnCriteria & "%')"

        If Not (ChkLsScanSt_DelNo.Checked = True And ChkLsScanSt_DelYes.Checked = True) Then
            If ChkLsScanSt_DelNo.Checked = True Then
                vnQuery += vbCrLf & "       and abs(SOScanDeleted)=0"
            Else
                vnQuery += vbCrLf & "       and abs(SOScanDeleted)=1"
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
                If GrvLsScan.Rows(vn).Cells(ensColLsScan.vSOScanDeleted).Text = "Y" Then
                    GrvLsScan.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Private Sub psFillGrvDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnCompanyCode As String = DstCompany.SelectedValue

        If vriHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,0 SOStockQty,0 vSumSOScanQty,0 vSOStockScanVarian,"
            vnQuery += vbCrLf & "       ''vSOStockNote,''vSOStockNoteBy,Null vSOStockNoteDatetime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()
        Else
            vnQuery = "Select Row_Number()over(order by mb.BRGNAME)vNo,d.OID,"
            vnQuery += vbCrLf & "       d.BRGCODE,mb.BRGNAME,mb.BRGUNIT,d.SOStockQty,d.vSumSOScanQty,d.vSOStockScanVarian,"
            vnQuery += vbCrLf & "       d.vSOStockNote,d.vSOStockNoteBy,d.vSOStockNoteDatetime"
            vnQuery += vbCrLf & "  From fnTbl_SsoSOStockScan(" & vriHOID & ")d"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb on mb.BRGCODE=d.BRGCODE and mb.CompanyCode='" & vnCompanyCode & "'"

            vnQuery += vbCrLf & " Where 1=1"

            If ChkFindNotActive.Checked Then
                vnQuery += vbCrLf & " and abs(mb.IsActive)=0"
            Else
                vnQuery += vbCrLf & " and abs(mb.IsActive)=1"
            End If

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " and (d.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%')"
            End If
            If ChkFindVarian.Checked Then
                vnQuery += vbCrLf & " and d.vSOStockScanVarian<>0"
            End If
            If ChkFindScan.Checked Then
                vnQuery += vbCrLf & " and d.vSumSOScanQty>0"
            End If

            vnQuery += vbCrLf & "Order by mb.BRGNAME"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            Dim vn As Integer
            If HdfActionStatus.Value = cbuActionNorm Then
                GrvDetail.Columns(ensColDetail.vSOStockNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.vSOStockNote).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.TxtvSOStockNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.TxtvSOStockNote).ItemStyle.CssClass = "myDisplayNone"
            Else
                GrvDetail.Columns(ensColDetail.vSOStockNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.vSOStockNote).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtvSOStockNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtvSOStockNote).ItemStyle.CssClass = ""
            End If

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()

            Dim vnGRow As GridViewRow
            If HdfActionStatus.Value = cbuActionEdit Then
                Dim vnTxtvSOStockNote As TextBox

                For vn = 0 To GrvDetail.Rows.Count - 1
                    vnGRow = GrvDetail.Rows(vn)
                    vnTxtvSOStockNote = vnGRow.FindControl("TxtvSOStockNote")

                    vnTxtvSOStockNote.Text = Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vSOStockNote).Text), "<br />", Chr(10))
                Next
            End If

            If HdfTransStatus.Value = enuTCSSOH.Scan_Open Or HdfTransStatus.Value = enuTCSSOH.Scan_Closed Then

            End If
        End If
    End Sub

    Private Sub psFillGrvTaDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,''BRGCODE,''BRGNAME,''BRGUNIT,0 StorageOID,''vStorageInfoHtml,0 SOScanQty,''SOScanNote,"
            vnQuery += vbCrLf & "       ''vSOScanUser,"
            vnQuery += vbCrLf & "	    ''vSOScanDeleted,"
            vnQuery += vbCrLf & "	    ''SOScanDeletedNote,''vSOScanDeletedUser,"
            vnQuery += vbCrLf & "	    ''vSOScanDeletedTime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvTaDetail.DataSource = vnDtb
            GrvTaDetail.DataBind()
        Else
            vnQuery = "Select Row_Number()over(order by mb.BRGNAME)vNo,sc.OID,sc.BRGCODE,mb.BRGNAME,mb.BRGUNIT,sc.StorageOID,sd.vStorageInfoHtml,sc.SOScanQty,sc.SOScanNote,"
            vnQuery += vbCrLf & "       mu.UserName vSOScanUser,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.SOScanDatetime,106) + ' ' + convert(varchar(5),sc.SOScanDatetime,108)vSOScanTime,"
            vnQuery += vbCrLf & "	    case when abs(sc.SOScanDeleted)=1 then 'Y' else 'N' end vSOScanDeleted,"
            vnQuery += vbCrLf & "	    sc.SOScanDeletedNote,du.UserID vSOScanDeletedUser,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.SOScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.SOScanDeletedDatetime,108)vSOScanDeletedTime"
            vnQuery += vbCrLf & "  From Sys_SsoSOScan_TR sc"
            vnQuery += vbCrLf & "       inner join Sys_SsoSOHeader_TR sh on sh.OID=sc.SOHOID"
            vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('')sd on sd.vStorageOID=sc.StorageOID"
            vnQuery += vbCrLf & "		inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb on mb.CompanyCode=sh.SOCompanyCode and mb.BRGCODE=sc.BRGCODE"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu on mu.OID=sc.SOScanUserOID"
            vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA du on du.OID=sc.SOScanDeletedUserOID"

            If ChkFindVarian.Checked Then
                vnQuery += vbCrLf & "       inner join fnTbl_SsoSOStockScan(" & vriHOID & ")ss on ss.BRGCODE=sc.BRGCODE and ss.vSOStockScanVarian<>0"
            End If

            vnQuery += vbCrLf & " Where sc.SOHOID=" & vriHOID

            If ChkFindNotActive.Checked Then
                vnQuery += vbCrLf & " and abs(mb.IsActive)=0"
            Else
                vnQuery += vbCrLf & " and abs(mb.IsActive)=1"
            End If

            If ChkFindIncludeDihapus.Checked = False Then
                vnQuery += vbCrLf & " and abs(sc.SOScanDeleted)=0"
            End If

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " and (mb.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%')"
            End If

            vnQuery += vbCrLf & "Order by mb.BRGNAME"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvTaDetail.DataSource = vnDtb
            GrvTaDetail.DataBind()

            If ChkFindIncludeDihapus.Checked = True Then
                Dim vn As Integer
                For vn = 0 To GrvTaDetail.Rows.Count - 1
                    If GrvTaDetail.Rows(vn).Cells(ensColTaDetail.vSOScanDeleted).Text = "Y" Then
                        GrvTaDetail.Rows(vn).ForeColor = Drawing.Color.Red
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub psSsoInsertSoStockNoQty(vriSQLConn As SqlConnection)

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

        BtnCancelSO.Enabled = False
        BtnScanOpen.Enabled = False
        BtnScanClosed.Enabled = False
        BtnCloseSO.Enabled = False

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
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(SONo,len(SONo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(SONo,len(SONo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoSOHeader_TR with(nolock)"
        vnQuery += vbCrLf & "      Where SONo like '" & vnTNoPrefix & "+'/%'"
        TxtSONo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session("UserLocationOID") = "0" Then
            LblMsgError.Text = "Anda Tidak Memiliki Akses Create Stock Opname"
            LblMsgError.Visible = True
            Exit Sub
        End If
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

        TxtSODate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)
        psFillGrvTaDetail(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgSONo.Text = ""
        LblMsgSODate.Text = ""
        LblMsgCompany.Text = ""
        LblMsgSubWhs.Text = ""
        LblMsgError.Text = ""
        LblXlsProses.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSODate.Enabled = vriBo
        TxtSONote.ReadOnly = Not vriBo

        If HdfActionStatus.Value = cbuActionNew Then
            TxtXlsWorksheet.ReadOnly = Not vriBo
            DstCompany.Enabled = vriBo
            DstSubWhs.Enabled = vriBo

            DstCutOffHour.Enabled = vriBo
            DstCutOffMin.Enabled = vriBo

            FupXls.Enabled = vriBo
        Else
            If HdfActionStatus.Value = cbuActionEdit Then
                TxtXlsWorksheet.ReadOnly = False
                DstCompany.Enabled = False
                DstSubWhs.Enabled = False

                FupXls.Enabled = False
            Else
                TxtXlsWorksheet.ReadOnly = False
                DstCompany.Enabled = True
                DstSubWhs.Enabled = True

                FupXls.Enabled = True
            End If

            DstCutOffHour.Enabled = True
            DstCutOffMin.Enabled = True
        End If

        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelSO.Visible = Not vriBo
        BtnScanOpen.Visible = Not vriBo
        BtnScanClosed.Visible = Not vriBo
        BtnCloseSO.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "vSumSOScanQty" Then
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

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoSSOH
            vnParam += "&vqTrNo=" & TxtSONo.Text

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
            psFillGrvTaDetail(0, vnSQLConn)

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

        vnQuery = "Select PM.*,convert(varchar(11),PM.SOCutOff,106)vSOCutOff,convert(varchar(5),PM.SOCutOff,108)vSOCutOff_HM,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoSOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoSSOH & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtSODate.Text = vnDtb.Rows(0).Item("vSOCutOff")
            TxtSONo.Text = vnDtb.Rows(0).Item("SONo")
            TxtSONote.Text = vnDtb.Rows(0).Item("SONote")

            DstCompany.SelectedValue = vnDtb.Rows(0).Item("SOCompanyCode")

            pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vriSQLConn)
            DstSubWhs.SelectedValue = vnDtb.Rows(0).Item("SOSubWarehouseOID")

            DstCutOffHour.SelectedValue = Mid(vnDtb.Rows(0).Item("vSOCutOff_HM"), 1, 2)
            DstCutOffMin.SelectedValue = Mid(vnDtb.Rows(0).Item("vSOCutOff_HM"), 4, 2)

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        RdbDetailType.SelectedValue = "Det"
        psChkDetFindVisible(True)

        GrvDetail.PageIndex = 0
        GrvTaDetail.PageIndex = 0

        psFillGrvTaDetail(Val(TxtTransID.Text), vriSQLConn)
        psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelSO.Visible = BtnCancelSO.Enabled
        BtnScanOpen.Visible = BtnScanOpen.Enabled
        BtnScanClosed.Visible = BtnScanClosed.Enabled
        BtnCloseSO.Visible = BtnCloseSO.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelSO.Enabled = False
        BtnCloseSO.Enabled = False
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
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCSSOH.Baru Or HdfTransStatus.Value = enuTCSSOH.Scan_Open Or HdfTransStatus.Value = enuTCSSOH.Scan_Closed)

            BtnCancelSO.Enabled = (HdfTransStatus.Value = enuTCSSOH.Baru Or HdfTransStatus.Value = enuTCSSOH.Scan_Open)

            BtnScanOpen.Enabled = (HdfTransStatus.Value = enuTCSSOH.Baru Or HdfTransStatus.Value = enuTCSSOH.Scan_Closed)
            BtnScanClosed.Enabled = (HdfTransStatus.Value = enuTCSSOH.Scan_Open)
            BtnCloseSO.Enabled = (HdfTransStatus.Value = enuTCSSOH.Scan_Closed)

            BtnPreview.Enabled = (HdfTransStatus.Value > enuTCSSOH.Baru)

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

        If Not IsDate(TxtSODate.Text) Then
            LblMsgSODate.Text = "Isi Tanggal Cut Off dengan benar"
            vnSave = False
        End If
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "0" Then
            LblMsgSubWhs.Text = "Pilih Sub Warehouse"
            vnSave = False
        End If
        If TxtXlsWorksheet.Text = "" Then
            LblXlsWorksheet.Text = "Isi Nama Worksheet"
            LblXlsWorksheet.Visible = True
            vnSave = False
        End If

        If Not vnSave Then Exit Sub

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psSaveBaru", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

        vsXlsFolder = Server.MapPath("~") & "\XlsFolder\"
        vsXlsFileName = vsXlsFolder & "SO_" & Format(Date.Now, "yyyyMMdd_HHmmss ") & FupXls.FileName
        vsTextStream.WriteLine("vnFileName : " & vsXlsFileName)

        vsSheetName = Trim(TxtXlsWorksheet.Text)
        vsTextStream.WriteLine("vnSheetName : " & vsSheetName)

        vsTextStream.WriteLine("FupXls.SaveAs(" & vsXlsFileName & ")...Start")

        FupXls.SaveAs(vsXlsFileName)

        vsTextStream.WriteLine("FupXls.SaveAs(" & vsXlsFileName & ")...End")

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            vsTextStream.WriteLine("Open SQL Connection....Start")
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgXlsProsesError.Text = pbMsgError
                LblMsgXlsProsesError.Visible = True

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
            Dim vnSubWhsOID As String = DstSubWhs.SelectedValue

            Dim vnQuery As String

            vnQuery = "Select count(1) From Sys_SsoSOHeader_TR Where SONo='" & Trim(TxtSONo.Text) & "'"
            If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                LblMsgSONo.Text = "No.SO " & Trim(TxtSONo.Text) & " Sudah pernah dipakai."

                vsTextStream.WriteLine(LblMsgSONo.Text)
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

            Dim vnWarehouseOID As String = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID, vnSQLConn)
            Dim vnSubWhsCode As String = fbuGetSubWhsCode(vnSubWhsOID, vnSQLConn)

            Dim vnOID As Integer
            vnQuery = "Select max(OID) from Sys_SsoSOHeader_TR"
            vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            psSetTransNo(vnCompanyCode, vnSubWhsCode, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoSOHeader_TR(OID,SONo,"
            vnQuery += vbCrLf & "SOTypeOID,SOCutOff,"
            vnQuery += vbCrLf & "SOCompanyCode,SOWarehouseOID,SOSubWarehouseOID,"
            vnQuery += vbCrLf & "SOXlsFileName,SOXlsSheetName,"
            vnQuery += vbCrLf & "SONote,"
            vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime,"
            vnQuery += vbCrLf & "SOStockDownload)"
            vnQuery += vbCrLf & "values(" & vnOID & ",'" & Trim(TxtSONo.Text) & "',"
            vnQuery += vbCrLf & enuSOType.WinAcc & ",'" & TxtSODate.Text & " " & DstCutOffHour.SelectedValue & ":" & DstCutOffMin.SelectedValue & "',"
            vnQuery += vbCrLf & "'" & vnCompanyCode & "'," & vnWarehouseOID & "," & vnSubWhsOID & ","
            vnQuery += vbCrLf & "'" & vsXlsFileName & "','" & vsSheetName & "',"
            vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtSONote.Text)) & "',"
            vnQuery += vbCrLf & "'" & stuTransCode.SsoSSOH & "'," & Session("UserOID") & ",getdate(),"
            vnQuery += vbCrLf & "1)"
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSSOH...Start")
            pbuInsertStatusSSOH(vnOID, enuTCSSOH.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSSOH...End")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------")
            vsTextStream.WriteLine("fsXlsImportData...Start")
            If fsXlsImportData(vnOID, vnCompanyCode, vnSubWhsCode, vnSQLConn, vnSQLTrans) Then
                vsTextStream.WriteLine("fsXlsImportData...End")
                vsTextStream.WriteLine("=======================")
                vsTextStream.WriteLine("")

                vnBeginTrans = False
                vnSQLTrans.Commit()

                Session(csModuleName & stuSession.Simpan) = "Done"
            Else
                vsTextStream.WriteLine("fsXlsImportData...Gagal")
                vsTextStream.WriteLine("=======================")
                vsTextStream.WriteLine("")

                vnBeginTrans = False
                vnSQLTrans.Rollback()
            End If

            vnSQLTrans = Nothing

            TxtTransID.Text = vnOID

            HdfTransStatus.Value = enuTCSSOH.Baru

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
            vsTextStream.WriteLine("Create Stock Opname Sukses")
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
            vsTextStream.WriteLine("Create Stock Opname Error")
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

    Private Function fsXlsImportData(vriHOID As String, vriCompanyCode As String, vriSubWhsCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        '24 Jun 2023 diperbaiki...ditambahi validasi gudang sesuai gudang di headernya
        fsXlsImportData = False
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        Const cnBRG = 0
        Const cnNAMABRG = 1
        Const cnGDG = 2
        Const cnQTYAKHIR = 5

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGUNIT As String
        Dim vnGDG As String

        Dim vnSOStockQty As String

        Dim vnQuery As String

        Dim vnPath As String = System.IO.Path.GetFullPath(vsXlsFileName)
        vsTextStream.WriteLine("vnPath : " & vnPath)

        Dim vnXConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vsXlsFileName & ";" & "Extended Properties=Excel 12.0 Xml;" ';IMEX=1;HDR=YES
        vsTextStream.WriteLine("vnXConnStr : " & vnXConnStr)

        Dim vnXConn As New OleDbConnection(vnXConnStr)
        vnXConn.Open()

        Dim vnXCommand As OleDbCommand
        vnXCommand = vnXConn.CreateCommand()
        vnXCommand.CommandText = "Select * From [" & vsSheetName & "$]"
        vsTextStream.WriteLine("")

        Dim vnXReader As OleDbDataReader
        vnXReader = vnXCommand.ExecuteReader

        Dim vnsSubWhs As String = DstSubWhs.SelectedItem.Text

        Dim vnNo As Integer
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnBRGCODE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRG))))
            vsTextStream.WriteLine("vnBRGCODE " & vnBRGCODE)

            vnSOStockQty = Val(fbuValStr(vnXReader.Item(cnQTYAKHIR)))
            vsTextStream.WriteLine("vnSOStockQty " & vnSOStockQty)

            vnGDG = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnGDG))))
            vsTextStream.WriteLine("vnGDG " & vnGDG)

            If vnBRGCODE <> "" Then
                If vnGDG = Mid(vnsSubWhs, 1, Len(vnGDG)) Then
                    vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnNAMABRG))))
                    vnBRGUNIT = fbuGetBarangUnitTrans(vnCompanyCode, vnBRGCODE, vriSQLConn, vriSQLTrans)

                    vnQuery = "Insert into Sys_SsoSOStock_TMP("
                    vnQuery += vbCrLf & "SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
                    vnQuery += vbCrLf & "Select '" & vriHOID & "' SOHOID,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnSOStockQty & " vnSOStockQty"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("Tidak Diupload --> Gudang Tidak Sesuai")
                    vsTextStream.WriteLine("vnGDG = " & vnGDG)
                    vsTextStream.WriteLine("Mid(vnsSubWhs, 1, Len(vnGDG)) = " & Mid(vnsSubWhs, 1, Len(vnGDG)))
                End If
            Else
                vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
                Exit While
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("<--------------------Insert Summary...Start")
        vnQuery = "Insert into Sys_SsoSOStock_TR("
        vnQuery += vbCrLf & "SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
        vnQuery += vbCrLf & "Select " & vriHOID & " SOHOID,BRGCODE,BRGNAME,BRGUNIT,isnull(sum(SOStockQty),0) From Sys_SsoSOStock_TMP With(nolock) Where SOHOID=" & vriHOID & " Group by BRGCODE,BRGNAME,BRGUNIT"
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        vsTextStream.WriteLine("<--------------------Insert Summary...End")
        vsTextStream.WriteLine("")

        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("<--------------------Delete Temporary...Start")
        vnQuery = "Delete Sys_SsoSOStock_TMP Where SOHOID=" & vriHOID
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vriSQLTrans)
        vsTextStream.WriteLine("<--------------------Delete Temporary...End")
        vsTextStream.WriteLine("")

        '<---11 Jul 2023 Ga Jadi Pake
        'vsTextStream.WriteLine("")
        'vsTextStream.WriteLine("<--------------------Insert master Sys_SsoSOBrg_TR...Start")
        'vnQuery = "Insert into Sys_SsoSOBrg_TR(SOHOID,BRGCODE,isActive)"
        'vnQuery += vbCrLf & "Select " & vriHOID & ",mb.BRGCODE,mb.isActive"
        'vnQuery += vbCrLf & "  From Sys_SsoSOStock_TR st with(nolock)"
        'vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=st.BRGCODE and mb.CompanyCode='" & vriCompanyCode & "' and mb.isActive=0"
        'vnQuery += vbCrLf & " Where st.SOHOID=" & vriHOID
        'vsTextStream.WriteLine(vnQuery)
        'pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
        'vsTextStream.WriteLine("<--------------------Insert master Sys_SsoSOBrg_TR...End")
        'vsTextStream.WriteLine("")
        '<<==11 Jul 2023 Ga Jadi Pake

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData = True
    End Function

    Private Sub psRefreshSsoSOBrg_20230711_Ga_Jadi_Pake(vriHOID As String, vriCompanyCode As String, vriSQLConn As SqlConnection)
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psRefreshSsoSOBrg", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean
        Dim vnQuery As String

        Try
            vnSQLTrans = vriSQLConn.BeginTransaction("inp")
            vnBeginTrans = True

            vsTextStream.WriteLine("")
            vnQuery = "Insert into Sys_SsoSOBrg_TR(SOHOID,BRGCODE,isActive)"
            vnQuery += vbCrLf & "Select " & vriHOID & ",mb.BRGCODE,mb.isActive"
            vnQuery += vbCrLf & "  From Sys_SsoSOScan_TR st with(nolock)"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=st.BRGCODE and mb.CompanyCode='" & vriCompanyCode & "' and mb.isActive=0"
            vnQuery += vbCrLf & " Where st.SOHOID=" & vriHOID & " and"
            vnQuery += vbCrLf & "       not st.BRGCODE in(Select b.BRGCODE From Sys_SsoSOBrg_TR b with(nolock) Where b.SOHOID=" & vriHOID & ")"

            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("")

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("psRefreshSsoSOBrg Done...")
        Catch ex As Exception
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("psRefreshSsoSOBrg ERROR...")

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
            End If
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
        End Try

        vsTextStream.WriteLine("")
        vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vsTextStream.WriteLine("------------------------EOF------------------------")

        vsTextStream.Close()
        vsTextStream = Nothing
        vsFso = Nothing
    End Sub

    Private Function fsXlsImportData_20230628_Error_Gudang_Belom_Divalidasi(vriHOID As String, vriSubWhsCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData_20230628_Error_Gudang_Belom_Divalidasi = False
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        Const cnBRG = 0
        Const cnNAMABRG = 1
        Const cnQTYAKHIR = 2

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGUNIT As String

        Dim vnSOStockQty As String

        Dim vnQuery As String

        Dim vnPath As String = System.IO.Path.GetFullPath(vsXlsFileName)
        vsTextStream.WriteLine("vnPath : " & vnPath)

        Dim vnXConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vsXlsFileName & ";" & "Extended Properties=Excel 12.0 Xml;" ';IMEX=1;HDR=YES
        vsTextStream.WriteLine("vnXConnStr : " & vnXConnStr)

        Dim vnXConn As New OleDbConnection(vnXConnStr)
        vnXConn.Open()

        Dim vnXCommand As OleDbCommand
        vnXCommand = vnXConn.CreateCommand()
        vnXCommand.CommandText = "Select * From [" & vsSheetName & "$]"
        vsTextStream.WriteLine("")

        Dim vnXReader As OleDbDataReader
        vnXReader = vnXCommand.ExecuteReader

        Dim vnNo As Integer
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnBRGCODE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRG))))
            vsTextStream.WriteLine("vnBRGCODE " & vnBRGCODE)

            vnSOStockQty = Val(fbuValStr(vnXReader.Item(cnQTYAKHIR)))

            If vnBRGCODE <> "" Then
                vsTextStream.WriteLine("BRGCODE NOT EXIST")
                vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnNAMABRG))))
                vnBRGUNIT = fbuGetBarangUnitTrans(vnCompanyCode, vnBRGCODE, vriSQLConn, vriSQLTrans)

                vnQuery = "Insert into Sys_SsoSOStock_TR("
                vnQuery += vbCrLf & "SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
                vnQuery += vbCrLf & "Select '" & vriHOID & "' SOHOID,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnSOStockQty & " vnSOStockQty"

                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
            Else
                vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
                Exit While
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData_20230628_Error_Gudang_Belom_Divalidasi = True
    End Function

    Private Function fsXlsImportData_20230125_Orig_FormatXls_Asli(vriHOID As String, vriSubWhsCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData_20230125_Orig_FormatXls_Asli = False
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        Const cnBRG = 0
        Const cnNAMABRG = 1
        Const cnGDG = 2
        Const cnNAMAGDG = 3
        Const cnUNIT = 4
        Const cnQTYAKHIR = 5

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGUNIT As String

        Dim vnGDG As String

        Dim vnSOStockQty As String

        Dim vnQuery As String

        Dim vnPath As String = System.IO.Path.GetFullPath(vsXlsFileName)
        vsTextStream.WriteLine("vnPath : " & vnPath)

        Dim vnXConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vsXlsFileName & ";" & "Extended Properties=Excel 12.0 Xml;" ';IMEX=1;HDR=YES
        vsTextStream.WriteLine("vnXConnStr : " & vnXConnStr)

        Dim vnXConn As New OleDbConnection(vnXConnStr)
        vnXConn.Open()

        Dim vnXCommand As OleDbCommand
        vnXCommand = vnXConn.CreateCommand()
        vnXCommand.CommandText = "Select * From [" & vsSheetName & "$]"
        vsTextStream.WriteLine("")

        Dim vnXReader As OleDbDataReader
        vnXReader = vnXCommand.ExecuteReader

        Dim vnNo As Integer
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnGDG = fbuValStr(vnXReader.Item(cnGDG))
            vsTextStream.WriteLine("vnGDG " & vnGDG)

            If vnGDG = vriSubWhsCode Then
                vnBRGCODE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRG))))
                vsTextStream.WriteLine("vnBRGCODE " & vnBRGCODE)

                vnSOStockQty = Val(fbuValStr(vnXReader.Item(cnQTYAKHIR)))

                If vnBRGCODE <> "" Then
                    vsTextStream.WriteLine("BRGCODE NOT EXIST")
                    vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnNAMABRG))))
                    vnBRGUNIT = Trim(fbuValStr(vnXReader.Item(cnUNIT)))

                    vnQuery = "Insert into Sys_SsoSOStock_TR("
                    vnQuery += vbCrLf & "SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
                    vnQuery += vbCrLf & "Select '" & vriHOID & "' SOHOID,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnSOStockQty & " vnSOStockQty"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
                    Exit While
                End If
            Else
                vsTextStream.WriteLine("vnGDG <> vnGudangCode --> " & vnGDG & "<>" & vriSubWhsCode)
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData_20230125_Orig_FormatXls_Asli = True
    End Function

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
                vnQuery = "Update Sys_SsoSOHeader_TR set "
                vnQuery += vbCrLf & "SOCutOff='" & TxtSODate.Text & " " & DstCutOffHour.SelectedValue & ":" & DstCutOffMin.SelectedValue & "',"
                vnQuery += vbCrLf & "SONote='" & fbuFormatString(Trim(TxtSONote.Text)) & "',"
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

    Private Sub psSaveDetail(vriOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnTxtvSOStockNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtvSOStockNote = vnGRow.FindControl("TxtvSOStockNote")
            'Trim(vnTxtvRcvDNote.Text) <> replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vRcvDNote).Text),"<br />",vbLf)
            If Trim(vnTxtvSOStockNote.Text) <> Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vSOStockNote).Text), "<br />", vbLf) Then
                vnQuery = "Update Sys_SsoSOStock_TR set "
                vnQuery += vbCrLf & "SOStockNote='" & fbuFormatString(vnTxtvSOStockNote.Text) & "',"
                vnQuery += vbCrLf & "SOStockNoteUserOID='" & Session("UserOID") & "',SOStockNoteDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                vnQuery = "Insert into Sys_SsoSOStock_HS "
                vnQuery += vbCrLf & "(SOSOID,SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty,SOStockNote,SOStockNoteUserOID,SOStockNoteDatetime)"
                vnQuery += vbCrLf & "Select OID,SOHOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty,SOStockNote,SOStockNoteUserOID,SOStockNoteDatetime"
                vnQuery += vbCrLf & "From Sys_SsoSOStock_TR Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
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

            If BtnEdit.Enabled Then
                BtnEdit.Enabled = False
                BtnEdit.Visible = BtnEdit.Enabled
            End If
        End If
    End Sub

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If HdfProcess.Value = "CancelSO" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelSO()
        ElseIf HdfProcess.Value = "ScanOpen" Then
            psScanOpen()
        ElseIf HdfProcess.Value = "ScanClosed" Then
            psScanClosed()
        ElseIf HdfProcess.Value = "CloseSO" Then
            psCloseSO()
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
        LblConfirmMessage.Text = "SO " & TxtSONo.Text & " Scan Open ?"
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
        LblConfirmMessage.Text = "SO " & TxtSONo.Text & " Scan Close ?"
        HdfProcess.Value = "ScanClosed"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub
    Private Sub psCancelSO()
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

            vnQuery = "Update Sys_SsoSOHeader_TR set TransStatus=" & enuTCSSOH.Cancelled & ",SOCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSSOH(TxtTransID.Text, enuTCSSOH.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psCloseSO()
        Dim vnSOHOID As String = TxtTransID.Text
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psCloseSO", vnSOHOID, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

        Dim vnSQLConn As New SqlConnection

        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("Open SQLConnection...Start")
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("ERROR")
            vsTextStream.WriteLine(pbMsgError)
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            Exit Sub
        End If
        vsTextStream.WriteLine("Open SQLConnection...Sukses")

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnQuery As String
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoSOBrg_TR(SOHOID,BRGCODE,isActive)"
            vnQuery += vbCrLf & "Select " & vnSOHOID & ",mb.BRGCODE,mb.isActive"
            vnQuery += vbCrLf & "  From Sys_SsoSOStock_TR st with(nolock)"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=st.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "' and mb.isActive=0"
            vnQuery += vbCrLf & " Where st.SOHOID=" & vnSOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vnQuery = "Insert into Sys_SsoSOBrg_TR(SOHOID,BRGCODE,isActive)"
            vnQuery += vbCrLf & "Select " & vnSOHOID & ",mb.BRGCODE,mb.isActive"
            vnQuery += vbCrLf & "  From Sys_SsoSOScan_TR st with(nolock)"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=st.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "' and mb.isActive=0"
            vnQuery += vbCrLf & " Where st.SOHOID=" & vnSOHOID & " and"
            vnQuery += vbCrLf & "       not st.BRGCODE in(Select b.BRGCODE From Sys_SsoSOBrg_TR b with(nolock) Where b.SOHOID=" & vnSOHOID & ")"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoSOHeader_TR set TransStatus=" & enuTCSSOH.Closed & ",SOCloseNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "ClosedUserOID=" & Session("UserOID") & ",ClosedDatetime=getdate() Where OID=" & vnSOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("4")
            pbuInsertStatusSSOH(vnSOHOID, enuTCSSOH.Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("5")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            psDisplayData(vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("ERROR")
            vsTextStream.WriteLine(ex.Message)
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

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

            vnQuery = "Update Sys_SsoSOHeader_TR set TransStatus=" & enuTCSSOH.Scan_Open & ",ScanOpenUserOID=" & Session("UserOID") & ",ScanOpenDatetime=getdate() Where OID=" & TxtTransID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSSOH...Start")
            pbuInsertStatusSSOH(TxtTransID.Text, enuTCSSOH.Scan_Open, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSSOH...End")

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

            vnQuery = "Update Sys_SsoSOHeader_TR Set TransStatus=" & enuTCSSOH.Scan_Closed & ",ScanClosedUserOID=" & Session("UserOID") & ",ScanClosedDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSSOH(TxtTransID.Text, enuTCSSOH.Scan_Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Protected Sub BtnCancelSO_Click(sender As Object, e As EventArgs) Handles BtnCancelSO.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan SO No. " & TxtSONo.Text & " ?<br />WARNING : Batal SO Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelSO"
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

    Private Sub BtnCloseSO_Click(sender As Object, e As EventArgs) Handles BtnCloseSO.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Close SO No. " & TxtSONo.Text & " ?<br />WARNING : Close SO Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CloseSO"
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
        If Val(TxtTransID.Text) = 0 Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If RdbDetailType.SelectedValue = "Det" Then
            psFillGrvTaDetail(TxtTransID.Text, vnSQLConn)
        Else
            psFillGrvDetail(TxtTransID.Text, vnSQLConn)
        End If

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
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSODetail, "SO Detail"})
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOTally, "Tally SO"})
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOStatus, "SO Status"})

        DstProReport.DataSource = vnDtb
        DstProReport.DataValueField = "RptCode"
        DstProReport.DataTextField = "RptName"
        DstProReport.DataBind()
    End Sub

    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
        psCrpXls()
    End Sub

    Private Sub psCrpXls()
        If LCase(RdbProXls.SelectedValue) = "pdf" Then
            Dim vnCrpFileName As String = ""

            If DstProReport.SelectedValue = stuSsoReportType.RptSOTally Then
                psGenerateCrpTally(vnCrpFileName)
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptSODetail Then
                psGenerateCrpTallyDetail(vnCrpFileName)
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptSOStatus Then
                psGenerateCrpSOStatus(vnCrpFileName)
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
        Else
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            If DstProReport.SelectedValue = stuSsoReportType.RptSOTally Then
                pbuCreateXlsx_SOTally(stuSsoReportType.RptSOTally, TxtTransID.Text, IIf(ChkProVarianOnly.Checked, 1, 0), vnSQLConn)
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptSODetail Then
                pbuCreateXlsx_SOTallyDetail(stuSsoReportType.RptSODetail, TxtTransID.Text, IIf(ChkProVarianOnly.Checked, 1, 0), vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psGenerateCrpTally(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTally

        vbuCrpQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += vbCrLf & "            " & IIf(ChkProVarianOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += vbCrLf & "       From fnTbl_SsoTally(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        If ChkProVarianOnly.Checked Then
            vbuCrpQuery += vbCrLf & "       Where ta.vSOStockScanVarian!=0"
        End If
        vbuCrpQuery += vbCrLf & " order by mb.BRGNAME"
    End Sub

    Private Sub psGenerateCrpTallyDetail(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyDetail

        vbuCrpQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += "       From fnTbl_SsoTallyDetail(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('') stg on stg.vStorageOID=ta.StorageOID"
        If ChkProVarianOnly.Checked Then
            vbuCrpQuery += "       Where ta.BRGCODE in(Select b.BRGCODE From fnTbl_SsoTally(" & TxtTransID.Text & ",'" & Session("UserID") & "') b Where b.vSOStockScanVarian!=0)"
        End If
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub
    Private Sub psGenerateCrpSOStatus(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOStatus

        vbuCrpQuery = "Select so.*,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName"
        vbuCrpQuery += vbCrLf & "       From fnTbl_SsoSOStatus('" & Session("UserID") & "')so"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=so.SOSubWarehouseOID"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=so.SOWarehouseOID"
        vbuCrpQuery += vbCrLf & "       Where so.SOHOID=" & TxtTransID.Text
        vbuCrpQuery += vbCrLf & " order by so.SOHOID"
    End Sub

    Protected Sub BtnLsScanDataFind_Click(sender As Object, e As EventArgs) Handles BtnLsScanDataFind.Click
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
    End Sub

    Protected Sub GrvLsScan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsScan.SelectedIndexChanged

    End Sub

    Private Sub GrvLsScan_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsScan.PageIndexChanging
        GrvLsScan.PageIndex = e.NewPageIndex
        psFillGrvLsScan(HdfLsScanBrgCode.Value)
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

    Protected Sub RdbDetailType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdbDetailType.SelectedIndexChanged
        If RdbDetailType.SelectedValue = "Det" Then
            GrvTaDetail.Visible = True
            GrvDetail.Visible = False

            If BtnEdit.Enabled Then
                BtnEdit.Enabled = False
                BtnEdit.Visible = BtnEdit.Enabled
            End If

            psChkDetFindVisible(True)
        Else
            GrvTaDetail.Visible = False
            GrvDetail.Visible = True

            psButtonStatus()

            psChkDetFindVisible(False)
        End If
    End Sub

    Private Sub psChkDetFindVisible(vriBo As Boolean)
        If vriBo Then
            ChkFindVarian.Visible = vriBo
            ChkFindScan.Visible = False
            ChkFindIncludeDihapus.Visible = vriBo
        Else
            ChkFindVarian.Visible = Not vriBo
            ChkFindScan.Visible = Not vriBo
            ChkFindIncludeDihapus.Visible = False
        End If
    End Sub

    Private Sub GrvTaDetail_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvTaDetail.PageIndexChanging
        GrvTaDetail.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvTaDetail(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class