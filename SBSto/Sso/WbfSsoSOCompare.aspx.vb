Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoSOCompare
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSOCompare"
    Const csTNoPrefix = "SOC"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
    End Enum

    Enum ensColLsSO
        OID = 0
        SONo = 1
        vSOCutOff = 2
        SOCompanyCode = 3
        SubWhsName = 4
        WarehouseName = 5
        SONote = 6
        TransStatusDescr = 7
    End Enum

    Enum ensColDetail
        vNo = 0
        OID = 1
        BRGCODE = 2
        BRGNAME = 3
        BRGUNIT = 4
        SOStockQty1 = 5
        SOStockQty2 = 6
        vSOStockQtyVarian = 7
        SOScanQty1 = 8
        SOScanQty2 = 9
        vSOScanQtyVarian = 10
        SOCompareDNote = 11
        TxtSOCompareDNote = 12
        vSOCompareDNoteBy = 13
        vSOCompareDNoteDatetime = 14
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        HdfTransID.Value = ""

        TxtTransStatus.Text = ""
        TxtSOHOID1.Text = ""
        HdfSOHOID1.Value = "0"
        TxtSODate1.Text = ""
        TxtSONo1.Text = ""
        TxtSONote1.Text = ""

        TxtSOHOID2.Text = ""
        HdfSOHOID2.Value = "0"
        TxtSODate2.Text = ""
        TxtSONo2.Text = ""
        TxtSONote2.Text = ""

        HdfTransStatus.Value = enuTCSSOC.Baru
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

        DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsSO.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoSSOC, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            End If

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

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_Closed.Checked = False Then
            ChkSt_Baru.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCSSOC.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCSSOC.Cancelled & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCSSOC.Closed & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.SOHOID1,PM.SOHOID2,"
        vnQuery += vbCrLf & "     convert(varchar(11),SO1.SOCutOff,106) + ' '+ convert(varchar(5),SO1.SOCutOff,108)vSOCutOff1,"
        vnQuery += vbCrLf & "     convert(varchar(11),SO2.SOCutOff,106) + ' '+ convert(varchar(5),SO2.SOCutOff,108)vSOCutOff2,"
        vnQuery += vbCrLf & "     PM.SOCompanyCode,GM.SubWhsName,LM.WarehouseName,PM.SOCompareNote,PM.SOCompareCloseNote,PM.SOCompareCancelNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.LastCompareDatetime,106)+' '+convert(varchar(5),PM.LastCompareDatetime,108)+' '+ LC.UserName vLastCompare,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ CL.UserName vClosed"

        vnQuery += vbCrLf & "From Sys_SsoSOCompareH_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoSOHeader_TR SO1 with(nolock) on SO1.OID=PM.SOHOID1"
        vnQuery += vbCrLf & "     inner join Sys_SsoSOHeader_TR SO2 with(nolock) on SO2.OID=PM.SOHOID2"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA LM with(nolock) on LM.OID=PM.SOWarehouseOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA GM with(nolock) on GM.OID=PM.SOSubWarehouseOID"

        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA LC with(nolock) on LC.OID=PM.LastCompareUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CL with(nolock) on CL.OID=PM.ClosedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.SOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.SOWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and SO1.SONo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and CAST(SO1.SOCutOff AS DATE) >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and CAST(SO1.SOCutOff AS DATE) <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.SOWarehouseOID = " & DstListWhs.SelectedValue
        End If
        vnQuery += vbCrLf & "Order by SO1.SONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvLsScan(vriBrgCode As String, vriSOHOID As String)
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
        vnQuery += vbCrLf & "	    inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageData('" & HttpContext.Current.Session("UserID") & "')sd on sd.vStorageOID=sc.StorageOID"
        vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu on mu.OID=sc.SOScanUserOID"
        vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA du on du.OID=sc.SOScanDeletedUserOID"
        vnQuery += vbCrLf & " Where sc.SOHOID=" & vriSOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"
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

        If vriHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,"
            vnQuery += vbCrLf & "       0 SOStockQty1,0 SOStockQty2,0 vSOStockQtyVarian,"
            vnQuery += vbCrLf & "       0 vSOStockScanVarian1,"
            vnQuery += vbCrLf & "       0 vSOScanQtyVarian,"
            vnQuery += vbCrLf & "       0 vSOStockScanVarian2,"
            vnQuery += vbCrLf & "       0 VstockScanVarianAll,"
            vnQuery += vbCrLf & "	 '' vSOCompareDNoteBy,	"
            vnQuery += vbCrLf & "       ''SOCompareDNote,''vSOCompareDNoteBy,Null vSOCompareDNoteDatetime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()
        Else
            vnQuery = "	Select Row_Number()over(order by mb.BRGNAME)vNo,d.OID,	"
            vnQuery += vbCrLf & "	 d.BRGCODE,mb.BRGNAME,mb.BRGUNIT,	"
            vnQuery += vbCrLf & "	 d.SOStockQty1,d.SOStockQty2,abs(d.SOStockQty1 - d.SOStockQty2)vSOStockQtyVarian,	"
            vnQuery += vbCrLf & "	abs(d.SOStockQty1 - d.SOScanQty1)vSOStockScanVarian1,	"
            vnQuery += vbCrLf & "	 d.SOScanQty1,d.SOScanQty2,abs(d.SOScanQty1 - d.SOScanQty2)vSOScanQtyVarian,	"
            vnQuery += vbCrLf & "	abs(d.SOStockQty2 - d.SOScanQty2)vSOStockScanVarian2,	"
            vnQuery += vbCrLf & "	abs((d.SOStockQty1 - d.SOScanQty1) - (d.SOStockQty2 - d.SOScanQty2)) VstockScanVarianAll ,	"
            vnQuery += vbCrLf & "	 Replace(d.SOCompareDNote,char(10),'<br />')SOCompareDNote,	"
            vnQuery += vbCrLf & "	 ud.UserName vSOCompareDNoteBy,	"
            vnQuery += vbCrLf & "	 convert(varchar(11),d.SOCompareDNoteDatetime,106)+' '+convert(varchar(5),d.SOCompareDNoteDatetime,108) vSOCompareDNoteDatetime	"
            vnQuery += vbCrLf & "	               From Sys_SsoSOCompareD_TR d with(nolock)	"
            vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=d.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "'"
            vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA ud with(nolock) on ud.OID=d.SOCompareDNoteUserOID"
            vnQuery += vbCrLf & " Where d.SOCHOID=" & vriHOID
            vnQuery += vbCrLf & "       and abs(mb.IsActive)=1"

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " and (d.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%')"
            End If

            If ChkFindVarianStock.Checked Then
                vnQuery += vbCrLf & " and (d.SOStockQty1 - d.SOStockQty2)<>0"
            End If
            If ChkFindVarianScan.Checked Then
                vnQuery += vbCrLf & " and (d.SOScanQty1 - d.SOScanQty2)<>0"
            End If
            If ChkFindScan.Checked Then
                vnQuery += vbCrLf & " and (d.SOScanQty1>0 or d.SOScanQty2>0)"
            End If

            vnQuery += vbCrLf & "Order by mb.BRGNAME"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            Dim vn As Integer
            If HdfActionStatus.Value = cbuActionNorm Then
                GrvDetail.Columns(ensColDetail.SOCompareDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.SOCompareDNote).ItemStyle.CssClass = ""

                GrvDetail.Columns(ensColDetail.TxtSOCompareDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.TxtSOCompareDNote).ItemStyle.CssClass = "myDisplayNone"
            Else
                GrvDetail.Columns(ensColDetail.SOCompareDNote).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail.Columns(ensColDetail.SOCompareDNote).ItemStyle.CssClass = "myDisplayNone"

                GrvDetail.Columns(ensColDetail.TxtSOCompareDNote).HeaderStyle.CssClass = ""
                GrvDetail.Columns(ensColDetail.TxtSOCompareDNote).ItemStyle.CssClass = ""
            End If

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()

            Dim vnGRow As GridViewRow
            If HdfActionStatus.Value = cbuActionEdit Then
                Dim vnTxtvSOStockNote As TextBox

                For vn = 0 To GrvDetail.Rows.Count - 1
                    vnGRow = GrvDetail.Rows(vn)
                    vnTxtvSOStockNote = vnGRow.FindControl("TxtSOCompareDNote")

                    vnTxtvSOStockNote.Text = Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.SOCompareDNote).Text), "<br />", Chr(10))
                Next
            End If
        End If
    End Sub

    Private Sub psFillGrvTaDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = 0 Then
            vnQuery = "Select 0 vDSeqNo,''BRGNAME,''BRGCODE,''BRGUNIT,''vStorageInfo,"
            vnQuery += vbCrLf & "       0 vStorageOID,0 vSOScanQty1,0 vSOScanQty2,0 vSOScanVarian"
            vnQuery += vbCrLf & "	    ,''vScanByName1,''vScanByName2"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvTaDetail.DataSource = vnDtb
            GrvTaDetail.DataBind()
        Else
            Dim vnDBMaster As String = fbuGetDBMaster()

            vnQuery = "Select row_number()over(order by mb.BRGNAME)vDSeqNo,ta.BRGCODE,mb.BRGNAME,mb.BRGUNIT,stg.vStorageInfo,"
            vnQuery += vbCrLf & "       ta.vStorageOID,ta.vSOScanQty1,ta.vSOScanQty2,ta.vSOScanVarian"
            vnQuery += vbCrLf & "	    ,ta.vScanByName1,ta.vScanByName2"
            vnQuery += vbCrLf & "  From fnTbl_SsoTallyCompareDetail(" & vriHOID & "," & HdfSOHOID1.Value & "," & HdfSOHOID2.Value & ",'bnsrph')ta"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageData('" & HttpContext.Current.Session("UserID") & "') stg on stg.vStorageOID=ta.vStorageOID"

            If Trim(TxtFind.Text) <> "" Then
                Dim vnCr As String = fbuFormatString(Trim(TxtFind.Text))
                vnQuery += vbCrLf & " and (ta.BRGCODE like '%" & vnCr & "%' or mb.BRGNAME like '%" & vnCr & "%')"
            End If
            If ChkFindVarianScan.Checked Then
                vnQuery += vbCrLf & " and ta.vSOScanVarian<>0"
            End If

            vnQuery += vbCrLf & "Order by mb.BRGNAME"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvTaDetail.DataSource = vnDtb
            GrvTaDetail.DataBind()
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

        BtnCancelSO.Enabled = False
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
        LblMsgCompany.Text = ""
        LblMsgSubWhs.Text = ""
        LblMsgError.Text = ""
        LblMsgSOH1.Text = ""
        LblMsgSOH2.Text = ""
        LblProMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        If HdfActionStatus.Value = cbuActionNew Then
            DstCompany.Enabled = vriBo
            DstSubWhs.Enabled = vriBo
            BtnSOH1.Enabled = vriBo
            BtnSOH2.Enabled = vriBo
        Else
            If HdfActionStatus.Value = cbuActionEdit Then
                DstCompany.Enabled = False
                DstSubWhs.Enabled = False
            Else
                DstCompany.Enabled = True
                DstSubWhs.Enabled = True
            End If
            BtnSOH1.Enabled = False
            BtnSOH2.Enabled = False
        End If

        BtnSOH1.Visible = BtnSOH1.Enabled
        BtnSOH2.Visible = BtnSOH2.Enabled

        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelSO.Visible = Not vriBo
        BtnCloseSO.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "SOScanQty1" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx
            HdfLsScanBrgCode.Value = vnGRow.Cells(ensColDetail.BRGCODE).Text
            HdfLsScanHOID.Value = HdfSOHOID1.Value
            psFillGrvLsScan(HdfLsScanBrgCode.Value, HdfLsScanHOID.Value)
            LblLsScanTitle.Text = "SCAN OID " & HdfLsScanHOID.Value & " Barang = " & vnGRow.Cells(ensColDetail.BRGCODE).Text & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text

            psShowLsScan(True)

        ElseIf e.CommandName = "SOScanQty2" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
            HdfDetailRowIdx.Value = vnIdx
            HdfLsScanBrgCode.Value = vnGRow.Cells(ensColDetail.BRGCODE).Text
            HdfLsScanHOID.Value = HdfSOHOID2.Value
            psFillGrvLsScan(HdfLsScanBrgCode.Value, HdfLsScanHOID.Value)
            LblLsScanTitle.Text = "SCAN OID " & HdfLsScanHOID.Value & " - Barang = " & vnGRow.Cells(ensColDetail.BRGCODE).Text & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text

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

    Private Sub psShowLsSO(vriBo As Boolean)
        If vriBo Then
            DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
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
        If Not IsNumeric(HdfTransID.Value) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & HdfTransID.Value
            vnParam += "&vqTrCode=" & stuTransCode.SsoSSOC
            vnParam += "&vqTrNo=" & TxtSONo1.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)

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

        If Val(HdfTransID.Value) = 0 Then
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
        If Val(HdfTransID.Value) = 0 Then
            psClearData()
            Exit Sub
        End If

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,"
        vnQuery += vbCrLf & "     SO1.SONo vSONo1,SO2.SONo vSONo2,"
        vnQuery += vbCrLf & "     SO1.SONote vSONote1,SO2.SONote vSONote2,"
        vnQuery += vbCrLf & "     convert(varchar(11),SO1.SOCutOff,106) + ' '+ convert(varchar(5),SO1.SOCutOff,108)vSOCutOff1,"
        vnQuery += vbCrLf & "     convert(varchar(11),SO2.SOCutOff,106) + ' '+ convert(varchar(5),SO2.SOCutOff,108)vSOCutOff2,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoSOCompareH_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoSOHeader_TR SO1 with(nolock) on SO1.OID=PM.SOHOID1"
        vnQuery += vbCrLf & "     inner join Sys_SsoSOHeader_TR SO2 with(nolock) on SO2.OID=PM.SOHOID2"

        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoSSOC & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & HdfTransID.Value
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtSODate1.Text = vnDtb.Rows(0).Item("vSOCutOff1")
            TxtSODate2.Text = vnDtb.Rows(0).Item("vSOCutOff2")
            TxtSONo1.Text = vnDtb.Rows(0).Item("vSONo1")
            TxtSONo2.Text = vnDtb.Rows(0).Item("vSONo2")
            TxtSONote1.Text = vnDtb.Rows(0).Item("vSONote1")
            TxtSONote2.Text = vnDtb.Rows(0).Item("vSONote2")

            TxtSOHOID1.Text = vnDtb.Rows(0).Item("SOHOID1")
            TxtSOHOID2.Text = vnDtb.Rows(0).Item("SOHOID2")

            HdfSOHOID1.Value = TxtSOHOID1.Text
            HdfSOHOID2.Value = TxtSOHOID2.Text

            DstCompany.SelectedValue = vnDtb.Rows(0).Item("SOCompanyCode")



            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            'pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vriSQLConn).

            'If () Then



        End If

        DstSubWhs.SelectedValue = vnDtb.Rows(0).Item("SOSubWarehouseOID")
        psButtonStatus()
        GrvDetail.PageIndex = 0
        GrvTaDetail.PageIndex = 0

        If HdfActionStatus.Value = cbuActionNorm Then
            RdbDetailType.SelectedValue = "Det"
            psChkDetFindVisible(True)

            psFillGrvTaDetail(HdfTransID.Value, vriSQLConn)
            psFillGrvDetail(HdfTransID.Value, vriSQLConn)
        Else
            psFillGrvTaDetail(HdfTransID.Value, vriSQLConn)
            psFillGrvDetail(HdfTransID.Value, vriSQLConn)
        End If

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelSO.Visible = BtnCancelSO.Enabled
        BtnCloseSO.Visible = BtnCloseSO.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelSO.Enabled = False
        BtnCloseSO.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If Val(HdfTransID.Value) = 0 Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCSSOC.Baru)

            BtnCancelSO.Enabled = (HdfTransStatus.Value = enuTCSSOC.Baru)

            BtnCloseSO.Enabled = (HdfTransStatus.Value = enuTCSSOC.Baru)

            BtnPreview.Enabled = (HdfTransStatus.Value <> enuTCSSOC.Cancelled)

            psButtonVisible()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Val(HdfTransID.Value) = 0 Then Exit Sub
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

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "0" Then
            LblMsgSubWhs.Text = "Pilih Sub Warehouse"
            vnSave = False
        End If
        If HdfSOHOID1.Value = HdfSOHOID2.Value Then
            LblMsgSOH1.Text = "SO 1 = SO 2"
            vnSave = False
        End If
        If Val(HdfSOHOID1.Value) = 0 Then
            LblMsgSOH1.Text = "Pilih SO 1"
            vnSave = False
        End If
        If Val(HdfSOHOID2.Value) = 0 Then
            LblMsgSOH2.Text = "Pilih SO 2"
            vnSave = False
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

            Dim vnWarehouseOID As String = fbuGetWarehouseOID_BySubWhsOID(vnSubWhsOID, vnSQLConn)

            Dim vnQuery As String

            Dim vnSOCHOID As Integer
            vnQuery = "Select max(OID) from Sys_SsoSOCompareH_TR"
            vnSOCHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoSOCompareH_TR(OID,"
            vnQuery += vbCrLf & "SOHOID1,SOHOID2,"
            vnQuery += vbCrLf & "SOCompanyCode,SOWarehouseOID,SOSubWarehouseOID,"
            vnQuery += vbCrLf & "SOCompareNote,"
            vnQuery += vbCrLf & "LastCompareUserOID,LastCompareDatetime,"
            vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
            vnQuery += vbCrLf & "values(" & vnSOCHOID & ","
            vnQuery += vbCrLf & "'" & HdfSOHOID1.Value & "','" & HdfSOHOID2.Value & "',"
            vnQuery += vbCrLf & "'" & vnCompanyCode & "'," & vnWarehouseOID & "," & vnSubWhsOID & ","
            vnQuery += vbCrLf & "'',"
            vnQuery += vbCrLf & Session("UserOID") & ",getdate(),"
            vnQuery += vbCrLf & "'" & stuTransCode.SsoSSOC & "'," & Session("UserOID") & ",getdate())"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------")
            vsTextStream.WriteLine("Copy Detail Barang From SO1-SO2...Start")

            vnQuery = "Insert into Sys_SsoSOCompareD_TR"
            vnQuery += vbCrLf & "(SOCHOID,BRGCODE,SOStockQty1)"
            vnQuery += vbCrLf & "Select " & vnSOCHOID & " SOCHOID,BRGCODE,SOStockQty From Sys_SsoSOStock_TR Where SOHOID=" & HdfSOHOID1.Value
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2.1")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vnQuery = "Insert into Sys_SsoSOCompareD_TR"
            vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
            vnQuery += vbCrLf & "Select " & vnSOCHOID & " SOCHOID,BRGCODE From Sys_SsoSOStock_TR Where SOHOID=" & HdfSOHOID2.Value
            vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vnSOCHOID & ")"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3.1")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vnQuery = "Update scd set SOStockQty2=so2.SOStockQty"
            vnQuery += vbCrLf & "From Sys_SsoSOCompareD_TR scd"
            vnQuery += vbCrLf & "     inner join Sys_SsoSOStock_TR so2 on so2.BRGCODE=scd.BRGCODE and SOHOID=" & HdfSOHOID2.Value
            vnQuery += vbCrLf & "Where scd.SOCHOID=" & vnSOCHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3.1")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            psRecompareSO(vnSOCHOID, HdfSOHOID1.Value, HdfSOHOID2.Value, vnSQLConn, vnSQLTrans)

            'vnQuery = "Insert into Sys_SsoSOCompareD_TR"
            'vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
            'vnQuery += vbCrLf & "Select distinct " & vnOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID1.Value
            'vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vnOID & ")"
            'vsTextStream.WriteLine("")
            'vsTextStream.WriteLine("4")
            'vsTextStream.WriteLine(vnQuery)
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            'vnQuery = "Insert into Sys_SsoSOCompareD_TR"
            'vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
            'vnQuery += vbCrLf & "Select distinct " & vnOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID2.Value
            'vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vnOID & ")"
            'vsTextStream.WriteLine("")
            'vsTextStream.WriteLine("5")
            'vsTextStream.WriteLine(vnQuery)
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            'vnQuery = "Update scd set SOScanQty1=vSumSOScanQty"
            'vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
            'vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty From Sys_SsoSOScan_TR where SOHOID=" & HdfSOHOID1.Value & " group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
            'vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vnOID
            'vsTextStream.WriteLine("")
            'vsTextStream.WriteLine("6")
            'vsTextStream.WriteLine(vnQuery)
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            'vnQuery = "Update scd set SOScanQty2=vSumSOScanQty"
            'vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
            'vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty From Sys_SsoSOScan_TR where SOHOID=" & HdfSOHOID2.Value & " group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
            'vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vnOID
            'vsTextStream.WriteLine("")
            'vsTextStream.WriteLine("7")
            'vsTextStream.WriteLine(vnQuery)
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            'vnQuery = "Update Sys_SsoSOCompareD_TR set SOScanQtyVarian=SOScanQty1-SOScanQty2 Where SOCHOID=" & vnOID
            'vsTextStream.WriteLine("")
            'vsTextStream.WriteLine("8")
            'vsTextStream.WriteLine(vnQuery)
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("Copy Detail Barang From SO1-SO2...End")
            vsTextStream.WriteLine("=======================")
            vsTextStream.WriteLine("")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSOCompareH...Start")
            pbuInsertStatusSOCompareH(vnSOCHOID, enuTCSSOC.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSOCompareH...End")

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            TxtTransID.Text = vnSOCHOID
            HdfTransID.Value = vnSOCHOID

            HdfTransStatus.Value = enuTCSSOC.Baru

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
            vsTextStream.WriteLine("Create Stock Opname Sukses")
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

            Dim vnUserNIP As String = Session("EmpNIP")

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            psSaveDetail(HdfTransID.Value, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSOCompareH(HdfTransID.Value, enuTCSSOC.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            Session(csModuleName & stuSession.Simpan) = "Done"

            psEnableInput(False)
            psEnableSave(False)

            GrvDetail.PagerSettings.Visible = True
            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            HdfActionStatus.Value = cbuActionNorm
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
        Dim vnTxtSOCompareDNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtSOCompareDNote = vnGRow.FindControl("TxtSOCompareDNote")
            If Trim(vnTxtSOCompareDNote.Text) <> Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.TxtSOCompareDNote).Text), "<br />", vbLf) Then
                vnQuery = "Update Sys_SsoSOCompareD_TR set "
                vnQuery += vbCrLf & "SOCompareDNote='" & fbuFormatString(vnTxtSOCompareDNote.Text) & "',"
                vnQuery += vbCrLf & "SOCompareDNoteUserOID='" & Session("UserOID") & "',SOCompareDNoteDatetime=getdate()"
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
        If e.CommandName = "OID" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)
            TxtTransID.Text = DirectCast(vnRow.Cells(ensColList.OID).Controls(0), LinkButton).Text
            HdfTransID.Value = DirectCast(vnRow.Cells(ensColList.OID).Controls(0), LinkButton).Text

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
        If HdfProcess.Value = "CancelSOCompare" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelSOCompare()
        ElseIf HdfProcess.Value = "CloseSO" Then
            psCloseSOCompare()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub
    Private Sub psCancelSOCompare()
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

            vnQuery = "Update Sys_SsoSOCompareH_TR set TransStatus=" & enuTCSSOC.Cancelled & ",SOCompareCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & HdfTransID.Value
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSOCompareH(HdfTransID.Value, enuTCSSOC.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

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

    Private Sub psCloseSOCompare()
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

            vnQuery = "Update Sys_SsoSOCompareH_TR set TransStatus=" & enuTCSSOC.Closed & ",SOCompareCloseNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "ClosedUserOID=" & Session("UserOID") & ",ClosedDatetime=getdate() Where OID=" & HdfTransID.Value
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSOCompareH(HdfTransID.Value, enuTCSSOC.Closed, Session("UserOID"), vnSQLConn, vnSQLTrans)

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
        LblConfirmMessage.Text = "Anda Membatalkan SO Compare " & TxtSOHOID1.Text & " - " & TxtSOHOID2.Text & " ?<br />WARNING : Batal SO Compare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelSOCompare"
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
        LblConfirmMessage.Text = "Anda Close SO No. " & TxtSONo1.Text & " ?<br />WARNING : Close SO Tidak Dapat Dibatalkan"
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

        psFillGrvDetail(HdfTransID.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If RdbDetailType.SelectedValue = "Det" Then
            psFillGrvTaDetail(HdfTransID.Value, vnSQLConn)
        Else
            psFillGrvDetail(HdfTransID.Value, vnSQLConn)
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
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOCompareDetail, "SO Compare Detail"})
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptSOTallyCompare, "Tally SO Compare"})

        DstProReport.DataSource = vnDtb
        DstProReport.DataValueField = "RptCode"
        DstProReport.DataTextField = "RptName"
        DstProReport.DataBind()
    End Sub

    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

        If HdfTransStatus.Value < enuTCSSOC.Closed Then
            'psReCompareSOStart()
        End If

        psCrpXls()
    End Sub

    Private Sub psCrpXls()
        If LCase(RdbProXls.SelectedValue) = "pdf" Then
            Dim vnCrpFileName As String = ""

            If DstProReport.SelectedValue = stuSsoReportType.RptSOTallyCompare Then
                psGenerateCrpTallyCompare(vnCrpFileName)
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptSOCompareDetail Then
                psGenerateCrpTallyCompareDetail(vnCrpFileName)
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

            If DstProReport.SelectedValue = stuSsoReportType.RptSOTallyCompare Then
                pbuCreateXlsx_SOTallyCompare(stuSsoReportType.RptSOTallyCompare2, TxtTransID.Text, IIf(ChkProVarianScanOnly.Checked, 1, 0), vnSQLConn)
                'pbuCreateXlsx_SOTallyCompare2(stuSsoReportType.RptSOTallyCompare2, TxtTransID.Text, IIf(ChkProVarianScanOnly.Checked, 1, 0), vnSQLConn)
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptSOCompareDetail Then
                pbuCreateXlsx_SOTallyCompareDetail(stuSsoReportType.RptSOCompareDetail2, TxtTransID.Text, HdfSOHOID1.Value, HdfSOHOID2.Value, IIf(ChkProVarianScanOnly.Checked, 1, 0), vnSQLConn)
                'pbuCreateXlsx_SOTallyCompareDetail2(stuSsoReportType.RptSOCompareDetail2, TxtTransID.Text, HdfSOHOID1.Value, HdfSOHOID2.Value, IIf(ChkProVarianScanOnly.Checked, 1, 0), vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            LblProMsgError.Text = pbMsgError
        End If
    End Sub
    Private Sub psGenerateCrpTallyCompare(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyCompare

        vbuCrpQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianScanOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += "       From fnTbl_SsoTallyCompare(" & HdfTransID.Value & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        If ChkProVarianScanOnly.Checked Then
            vbuCrpQuery += "       Where ta.vSOScanQtyVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub

    Private Sub psGenerateCrpTallyCompare2(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyCompareDetail2


        vbuCrpQuery = "	DECLARE @vriUser AS VARCHAR(25)	"
        vbuCrpQuery += vbCrLf & "	DECLARE @vriHOID AS INTEGER 	"
        vbuCrpQuery += vbCrLf & "	SET @vriUser = '" & HttpContext.Current.Session("UserID") & "'	"
        vbuCrpQuery += vbCrLf & "	SET @vriHOID = " & TxtTransID.Text & "	"
        vbuCrpQuery += vbCrLf & "	Select sh.OID,	"
        vbuCrpQuery += vbCrLf & "	       sh.SOCompanyCode,sh.SOWarehouseOID,sh.SOSubWarehouseOID,mc.CompanyName,sh	"
        vbuCrpQuery += vbCrLf & "	       sh.SOCompareCloseNote,sh.SOCompareCancelNote,	"
        vbuCrpQuery += vbCrLf & "	       sh.SOHOID1,so1.SONo vSONo1,so1.SOCutOff vSOCutOff1,so1.SONote vSONote1,so1.SOCloseNote vSOCloseNote1,so1.SOCancelNote vSOCancelNote1,	"
        vbuCrpQuery += vbCrLf & "	       sh.SOHOID2,so2.SONo vSONo2,so2.SOCutOff vSOCutOff2,so2.SONote vSONote2,so2.SOCloseNote vSOCloseNote2,so2.SOCancelNote vSOCancelNote2,	"
        vbuCrpQuery += vbCrLf & "	       sh.TransCode,sh.TransStatus,	"
        vbuCrpQuery += vbCrLf & "	       st.TransStatusDescr,	"
        vbuCrpQuery += vbCrLf & "	       sh.CreationUserOID,sh.CreationDatetime,	"
        vbuCrpQuery += vbCrLf & "	       sh.LastCompareUserOID,sh.LastCompareDatetime,	"
        vbuCrpQuery += vbCrLf & "	       sh.ClosedUserOID,sh.ClosedDatetime,	"
        vbuCrpQuery += vbCrLf & "	       sh.CancelledUserOID,sh.CancelledDatetime,	"
        vbuCrpQuery += vbCrLf & "	       sd.OID vDOID,	"
        vbuCrpQuery += vbCrLf & "	       sd.BRGCODE,	"
        vbuCrpQuery += vbCrLf & "	       sd.SOStockQty1,sd.SOStockQty2,(sd.SOStockQty1-sd.SOStockQty2)vSOStockQtyVarian,	"
        vbuCrpQuery += vbCrLf & "	       (sd.SOStockQty1-sd.SOScanQty1)vSOStockScan1,	"
        vbuCrpQuery += vbCrLf & "	       sd.SOScanQty1,sd.SOScanQty2,(sd.SOScanQty1-sd.SOScanQty2)vSOScanQtyVarian,	"
        vbuCrpQuery += vbCrLf & "	       (sd.SOStockQty2-sd.SOScanQty2)vSOStockScan2,	"
        vbuCrpQuery += vbCrLf & "	       (sd.SOStockQty1-sd.SOScanQty1) - (sd.SOStockQty2-sd.SOScanQty2) vSOStockScanAll,	"
        vbuCrpQuery += vbCrLf & "	       sd.SOCompareDNote,	"
        vbuCrpQuery += vbCrLf & "	       sd.SOCompareDNoteUserOID,su.UserName vSOCompareDNoteBy,	"
        vbuCrpQuery += vbCrLf & "	       sd.SOCompareDNoteDatetime,	"
        vbuCrpQuery += vbCrLf & "	       Convert(varchar(11),sd.SOCompareDNoteDatetime,106)+' '+Convert(varchar(5),sd.SOCompareDNoteDatetime,108) vSOCompareDNoteDatetime,	"
        vbuCrpQuery += vbCrLf & "	       Convert(varchar(11),getdate(),106)+' '+Convert(varchar(5),getdate(),108) vPrintDate	"
        vbuCrpQuery += vbCrLf & "	       ,@vriUser vPrintUser	"
        vbuCrpQuery += vbCrLf & "            " & IIf(ChkProVarianScanOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += vbCrLf & "	  From Sys_SsoSOCompareH_TR sh with(nolock)	"
        vbuCrpQuery += vbCrLf & "	       inner join Sys_SsoSOHeader_TR so1 with(nolock) on so1.OID=sh.SOHOID1	"
        vbuCrpQuery += vbCrLf & "	       inner join Sys_SsoSOHeader_TR so2 with(nolock) on so2.OID=sh.SOHOID2	"
        vbuCrpQuery += vbCrLf & "	       inner join Sys_SsoTransStatus_MA st with(nolock) on st.TransCode=sh.TransCode and st.TransStatus=sh.TransStatus	"
        vbuCrpQuery += vbCrLf & "	       inner join Sys_SsoSOCompareD_TR sd with(nolock) on sd.SOCHOID=sh.OID	"
        vbuCrpQuery += vbCrLf & "	       left outer join Sys_SsoUser_MA su with(nolock) on su.OID=sd.SOCompareDNoteUserOID	"
        vbuCrpQuery += vbCrLf & "	       inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=sh.SOCompanyCode	"
        vbuCrpQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sh.SOCompanyCode and mb.BRGCODE=sd.BRGCODE	"
        vbuCrpQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=sh.SOWarehouseOID	"
        vbuCrpQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=sh.SOWarehouseOID	"
        vbuCrpQuery += vbCrLf & "	       Where sh.OID=@vriHOID"

        If ChkProVarianScanOnly.Checked Then
            vbuCrpQuery += "       And ta.vSOScanQtyVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub

    Private Sub psGenerateCrpTallyCompareDetail(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyCompareDetail

        vbuCrpQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianScanOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += "       From fnTbl_SsoTallyCompareDetail(" & HdfTransID.Value & "," & HdfSOHOID1.Value & "," & HdfSOHOID2.Value & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('" & HttpContext.Current.Session("UserID") & "') stg on stg.vStorageOID=ta.vStorageOID"
        If ChkProVarianScanOnly.Checked Then
            vbuCrpQuery += "       Where ta.vSOScanVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME,stg.vStorageInfo"
    End Sub

    Private Sub psGenerateCrpTallyCompareDetail2(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOTallyCompareDetail2

        vbuCrpQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianScanOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += "       From fnTbl_SsoTallyCompareDetail(" & HdfTransID.Value & "," & HdfSOHOID1.Value & "," & HdfSOHOID2.Value & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('" & HttpContext.Current.Session("UserID") & "') stg on stg.vStorageOID=ta.vStorageOID"
        If ChkProVarianScanOnly.Checked Then
            vbuCrpQuery += "       Where ta.vSOScanVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME,stg.vStorageInfo"
    End Sub

    Protected Sub BtnLsScanDataFind_Click(sender As Object, e As EventArgs) Handles BtnLsScanDataFind.Click
        psFillGrvLsScan(HdfLsScanBrgCode.Value, HdfLsScanHOID.Value)
    End Sub

    Protected Sub GrvLsScan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsScan.SelectedIndexChanged

    End Sub

    Private Sub GrvLsScan_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsScan.PageIndexChanging
        GrvLsScan.PageIndex = e.NewPageIndex
        psFillGrvLsScan(HdfLsScanBrgCode.Value, HdfLsScanHOID.Value)
    End Sub

    Protected Sub BtnLsSOClose_Click(sender As Object, e As EventArgs) Handles BtnLsSOClose.Click
        psShowLsSO(False)
    End Sub

    Protected Sub BtnSOH1_Click(sender As Object, e As EventArgs) Handles BtnSOH1.Click
        psClearMessage()
        Dim vnSave As Boolean = True
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "0" Then
            LblMsgSubWhs.Text = "Pilih Sub Warehouse"
            vnSave = False
        End If
        If vnSave = False Then
            Exit Sub
        End If

        If Not IsDate(TxtLsSOStart.Text) Then
            TxtLsSOStart.Text = Format(DateAdd(DateInterval.Day, -1, Date.Now), "dd MMM yyyy")
        End If
        If Not IsDate(TxtLsSOEnd.Text) Then
            TxtLsSOEnd.Text = Format(Date.Now, "dd MMM yyyy")
        End If

        HdfLsSO.Value = 1
        psShowLsSO(True)
    End Sub

    Protected Sub BtnSOH2_Click(sender As Object, e As EventArgs) Handles BtnSOH2.Click
        psClearMessage()
        Dim vnSave As Boolean = True
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstSubWhs.SelectedValue = "0" Then
            LblMsgSubWhs.Text = "Pilih Sub Warehouse"
            vnSave = False
        End If
        If vnSave = False Then
            Exit Sub
        End If

        HdfLsSO.Value = 2
        psShowLsSO(True)
    End Sub

    Private Sub psFillGrvLsSO()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.SONo,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.SOCutOff,106) + ' '+ convert(varchar(5),PM.SOCutOff,108)vSOCutOff,"
        vnQuery += vbCrLf & "     PM.SOCompanyCode,GM.SubWhsName,LM.WarehouseName,PM.SONote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr"

        vnQuery += vbCrLf & "From Sys_SsoSOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA LM with(nolock) on LM.OID=PM.SOWarehouseOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA GM with(nolock) on GM.OID=PM.SOSubWarehouseOID"
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
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uc with(nolock) on uc.WarehouseOID=PM.SOWarehouseOID and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where PM.TransStatus > " & enuTCSSOH.Baru

        vnQuery += vbCrLf & "      and PM.SOTypeOID=" & enuSOType.WinAcc
        vnQuery += vbCrLf & "      and PM.SOCompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and PM.SOSubWarehouseOID='" & DstSubWhs.SelectedValue & "'"

        If Trim(TxtLsSO.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.SONo like '%" & Trim(TxtLsSO.Text) & "%'"
        End If

        If IsDate(TxtLsSOStart.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.SOCutOff AS DATE) >= '" & TxtLsSOStart.Text & "'"
        End If
        If IsDate(TxtLsSOEnd.Text) Then
            vnQuery += vbCrLf & "            and CAST(PM.SOCutOff AS DATE) <= '" & TxtLsSOEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.SONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvLsSO.DataSource = vnDtb
        GrvLsSO.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnLsSOFind_Click(sender As Object, e As EventArgs) Handles BtnLsSOFind.Click
        psFillGrvLsSO()
    End Sub

    Protected Sub GrvLsSO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsSO.SelectedIndexChanged

    End Sub

    Private Sub GrvLsSO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsSO.PageIndexChanging
        GrvLsSO.PageIndex = e.NewPageIndex
        psFillGrvLsSO()
    End Sub

    Private Sub GrvLsSO_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsSO.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "SONo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsSO.Rows(vnIdx)

            If HdfLsSO.Value = "1" Then
                HdfSOHOID1.Value = vnRow.Cells(ensColLsSO.OID).Text
                TxtSOHOID1.Text = HdfSOHOID1.Value
                TxtSONo1.Text = DirectCast(vnRow.Cells(ensColLsSO.SONo).Controls(0), LinkButton).Text
                TxtSONote1.Text = vnRow.Cells(ensColLsSO.SONote).Text
                TxtSODate1.Text = vnRow.Cells(ensColLsSO.vSOCutOff).Text
            Else
                HdfSOHOID2.Value = vnRow.Cells(ensColLsSO.OID).Text
                TxtSOHOID2.Text = HdfSOHOID2.Value
                TxtSONo2.Text = DirectCast(vnRow.Cells(ensColLsSO.SONo).Controls(0), LinkButton).Text
                TxtSONote2.Text = vnRow.Cells(ensColLsSO.SONote).Text
                TxtSODate2.Text = vnRow.Cells(ensColLsSO.vSOCutOff).Text
            End If
            psShowLsSO(False)
        End If
    End Sub


    Private Sub psReCompareSOStart()
        Dim vnSave As Boolean = True
        psClearMessage()

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psReCompareSOStart", HdfTransID.Value, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
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

            Dim vnSOCHOID As String = HdfTransID.Value
            Dim vnCompanyCode As String = DstCompany.SelectedValue

            Dim vnQuery As String

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            psRecompareSO(vnSOCHOID, HdfSOHOID1.Value, HdfSOHOID2.Value, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoSOCompareH_TR set LastCompareUserOID=" & Session("UserOID") & ",LastCompareDatetime=getdate() Where OID=" & vnSOCHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("Refresh Scan Qty 1-2...Start")
            pbuInsertStatusSOCompareH(vnSOCHOID, enuTCSSOC.Recompare, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("Refresh Scan Qty 1-2...End")

            vnBeginTrans = False
            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            If RdbDetailType.SelectedValue = "Det" Then
                psFillGrvTaDetail(vnSOCHOID, vnSQLConn)
            Else
                psFillGrvDetail(vnSOCHOID, vnSQLConn)
            End If

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
            vsTextStream.WriteLine("Create Stock Opname Sukses")
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

    Protected Sub BtnReCompareSO_Click(sender As Object, e As EventArgs) Handles BtnReCompareSO.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Val(HdfTransID.Value) = 0 Then Exit Sub
        psReCompareSOStart()
    End Sub

    Private Sub psRecompareSO(vriSOCHOID As String, vriSOH1 As String, vriSOH2 As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("<----psRecompareSO......Start")
        Dim vnQuery As String
        vnQuery = "Update Sys_SsoSOCompareD_TR set SOScanQty1=0,SOScanQty2=0 Where SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("0")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        'sc.StorageOID not in(Select sti.OID From " & vnDBMaster & "Sys_Storage_MA sti Where sti.StorageTypeOID= case when @vriSOHOID in(58,59) then 0 else 1007 end)

        vnQuery = "Insert into Sys_SsoSOCompareD_TR"
        vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
        vnQuery += vbCrLf & "Select distinct " & vriSOCHOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID1.Value & " and SOScanDeleted=0"
        vnQuery += vbCrLf & " and not StorageOID in(Select sti.OID From " & vnDBMaster & "Sys_Storage_MA sti Where sti.StorageTypeOID=" & enuStorageType.DO_Titip & ")"
        vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vriSOCHOID & ")"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("1")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSOCompareD_TR"
        vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
        vnQuery += vbCrLf & "Select distinct " & vriSOCHOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID2.Value & " and SOScanDeleted=0"
        vnQuery += vbCrLf & " and not StorageOID in(Select sti.OID From " & vnDBMaster & "Sys_Storage_MA sti Where sti.StorageTypeOID=" & enuStorageType.DO_Titip & ")"
        vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vriSOCHOID & ")"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("2")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Update scd set SOScanQty1=vSumSOScanQty"
        vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
        vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty"
        vnQuery += vbCrLf & "			              From Sys_SsoSOScan_TR with(nolock) where SOHOID=" & HdfSOHOID1.Value & " and abs(SOScanDeleted)=0"
        vnQuery += vbCrLf & "                              and not StorageOID in(Select sti.OID From " & vnDBMaster & "Sys_Storage_MA sti Where sti.StorageTypeOID=" & enuStorageType.DO_Titip & ")"
        vnQuery += vbCrLf & "			             group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
        vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("3")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Update scd set SOScanQty2=vSumSOScanQty"
        vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
        vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty"
        vnQuery += vbCrLf & "			              From Sys_SsoSOScan_TR where SOHOID=" & HdfSOHOID2.Value & " and abs(SOScanDeleted)=0"
        vnQuery += vbCrLf & "                              and not StorageOID in(Select sti.OID From " & vnDBMaster & "Sys_Storage_MA sti Where sti.StorageTypeOID=" & enuStorageType.DO_Titip & ")"
        vnQuery += vbCrLf & "			             group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
        vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("4")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vsTextStream.WriteLine("<----psRecompareSO......End")
    End Sub

    Private Sub psRecompareSO_20230904_Orig_Bef_DOTitip_Dikeluarkan(vriSOCHOID As String, vriSOH1 As String, vriSOH2 As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("<----psRecompareSO......Start")
        Dim vnQuery As String
        vnQuery = "Update Sys_SsoSOCompareD_TR set SOScanQty1=0,SOScanQty2=0 Where SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("0")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSOCompareD_TR"
        vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
        vnQuery += vbCrLf & "Select distinct " & vriSOCHOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID1.Value
        vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vriSOCHOID & ")"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("1")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSOCompareD_TR"
        vnQuery += vbCrLf & "(SOCHOID,BRGCODE)"
        vnQuery += vbCrLf & "Select distinct " & vriSOCHOID & " SOCHOID,BRGCODE From Sys_SsoSOScan_TR Where SOHOID=" & HdfSOHOID2.Value
        vnQuery += vbCrLf & " and not BRGCODE in(Select b.BRGCODE From Sys_SsoSOCompareD_TR b Where b.SOCHOID=" & vriSOCHOID & ")"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("2")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Update scd set SOScanQty1=vSumSOScanQty"
        vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
        vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty From Sys_SsoSOScan_TR where SOHOID=" & HdfSOHOID1.Value & " and abs(SOScanDeleted)=0 group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
        vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("3")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Update scd set SOScanQty2=vSumSOScanQty"
        vnQuery += vbCrLf & "       From Sys_SsoSOCompareD_TR scd"
        vnQuery += vbCrLf & "			inner join (Select BRGCODE,sum(SOScanQty)vSumSOScanQty From Sys_SsoSOScan_TR where SOHOID=" & HdfSOHOID2.Value & " and abs(SOScanDeleted)=0 group by BRGCODE)scn on scn.BRGCODE=scd.BRGCODE"
        vnQuery += vbCrLf & "	  Where scd.SOCHOID=" & vriSOCHOID
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("4")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vsTextStream.WriteLine("<----psRecompareSO......End")
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
            ChkFindVarianStock.Visible = Not vriBo
            ChkFindScan.Visible = Not vriBo
            ChkFindVarianScan.Visible = vriBo
        Else
            ChkFindVarianStock.Visible = Not vriBo
            ChkFindScan.Visible = Not vriBo
            ChkFindVarianScan.Visible = Not vriBo
        End If
    End Sub

    Protected Sub GrvTaDetail_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvTaDetail.SelectedIndexChanged

    End Sub

    Private Sub GrvTaDetail_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvTaDetail.PageIndexChanging
        GrvTaDetail.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvTaDetail(HdfTransID.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class