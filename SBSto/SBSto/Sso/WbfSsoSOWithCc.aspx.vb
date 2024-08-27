Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoSOWithCc
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSOWithCc"
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

    Enum ensColLsBrg
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
        vSelect = 3
        vRemove = 4
    End Enum

    Enum ensColLsSto
        WarehouseName = 0
        BuildingName = 1
        LantaiDescription = 2
        ZonaName = 3
        StorageTypeName = 4
        StorageSequenceNumber = 5
        StorageColumn = 6
        StorageLevel = 7
        StorageNumber = 8
        vStorageStagIO = 9
        vStorageOID = 10
        vStorageInfoHtml = 11
        vSelect = 12
        vRemove = 13
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtSODate.Text = ""
        TxtSONo.Text = ""
        TxtSONote.Text = ""

        TxtStockDownload.Text = ""
        HdfStockDownload.Value = "0"

        HdfTransStatus.Value = enuTCSSOH.Baru

        BtnTp02_Loc.Visible = False
        BtnTp03_Brg.Visible = False
        BtnTp03_Loc.Visible = False
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

        DivLsBrg.Style(HtmlTextWriterStyle.MarginLeft) = "250px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsBStg.Style(HtmlTextWriterStyle.MarginLeft) = "250px"
        DivLsBStg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBStg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsSto.Style(HtmlTextWriterStyle.MarginLeft) = "50px"
        DivLsSto.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsSto.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            psFillGrvLsBrg(Val(TxtTransID.Text))
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowLsBStg(vriBo As Boolean)
        If vriBo Then
            psFillGrvLsBStg(Val(TxtTransID.Text))
            DivLsBStg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBStg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowLsSto(vriBo As Boolean)
        If vriBo Then
            psSetUpLsSto()
            psFillGrvLsSto(Val(TxtTransID.Text))
            DivLsSto.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsSto.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
        If HdfActionStatus.Value = cbuActionNew Then Exit Sub
        psClearMessage()

        psFillGrvLsBrg(Val(TxtTransID.Text))
    End Sub
    Private Sub psFillGrvLsBrg(vriSOHOID As String)
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        LblMsgLsBrg.Text = ""
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgLsBrg.Text = pbMsgError
            Exit Sub
        End If

        Dim vnLsBrg As String = fbuFormatString(Trim(TxtLsBrg.Text))
        If ChkLsBrgSelectedNot.Checked = True Then
            If vnLsBrg = "" Then
                LblMsgLsBrg.Text = "Isi Nama Barang"
                Exit Sub
            End If
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnCriteria As String
        vnCriteria = vbCrLf & "Where CompanyCode='" & DstCompany.SelectedValue & "' and (PM.BRGCODE like '%" & vnLsBrg & "%' OR PM.BRGNAME like '%" & vnLsBrg & "%')"

        If ChkLsBrgSelectedNot.Checked Then
            vnQuery = "Select * From ("
            vnQuery += vbCrLf & "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,''vSelect,'Remove'vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_MstBarang_MA PM with(nolock)"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and PM.BrgCode in(Select b.BrgCode From Sys_SsoSOCcBarang_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"

            vnQuery += vbCrLf & "UNION"
            vnQuery += vbCrLf & "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,'Select'vSelect,''vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_MstBarang_MA PM with(nolock)"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and not PM.BrgCode in(Select b.BrgCode From Sys_SsoSOCcBarang_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"
            vnQuery += vbCrLf & ")b"
            vnQuery += vbCrLf & " Order by b.BRGCODE"
            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        Else
            vnQuery = "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,''vSelect,'Remove'vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_MstBarang_MA PM with(nolock)"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and PM.BrgCode in(Select b.BrgCode From Sys_SsoSOCcBarang_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"
            vnQuery += vbCrLf & " Order by PM.BRGCODE"
            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        End If

        GrvLsBrg.DataSource = vnDtb
        GrvLsBrg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psSetEditable_LsBrg()
    End Sub
    Private Sub psFillGrvLsBStg(vriSOHOID As String)
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        LblMsgLsBStg.Text = ""
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgLsBStg.Text = pbMsgError
            Exit Sub
        End If

        Dim vnLsBStg As String = fbuFormatString(Trim(TxtLsBStg.Text))

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.BRGCODE,mb.BRGNAME,mb.BRGUNIT,sto.vStorageOID,sto.vStorageInfoHtml"
        vnQuery += vbCrLf & " From Sys_SsoSOCcBarangStorage_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=PM.BRGCODE and CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') sto on sto.vStorageOID=PM.StorageOID"
        vnQuery += vbCrLf & "Where PM.SOHOID=" & vriSOHOID & " and (PM.BRGCODE like '%" & vnLsBStg & "%' OR mb.BRGNAME like '%" & vnLsBStg & "%')"
        vnQuery += vbCrLf & "Order by PM.BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsBStg.DataSource = vnDtb
        GrvLsBStg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psSetEditable_LsBrg()
        If HdfStockDownload.Value = "1" Then
            GrvLsBrg.Columns(ensColLsBrg.vSelect).HeaderStyle.CssClass = "myDisplayNone"
            GrvLsBrg.Columns(ensColLsBrg.vSelect).ItemStyle.CssClass = "myDisplayNone"

            GrvLsBrg.Columns(ensColLsBrg.vRemove).HeaderStyle.CssClass = "myDisplayNone"
            GrvLsBrg.Columns(ensColLsBrg.vRemove).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvLsBrg.Columns(ensColLsBrg.vSelect).HeaderStyle.CssClass = ""
            GrvLsBrg.Columns(ensColLsBrg.vSelect).ItemStyle.CssClass = ""

            GrvLsBrg.Columns(ensColLsBrg.vRemove).HeaderStyle.CssClass = ""
            GrvLsBrg.Columns(ensColLsBrg.vRemove).ItemStyle.CssClass = ""
        End If
    End Sub

    Protected Sub BtnLsBrgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBrgClose.Click
        psShowLsBrg(False)
    End Sub

    Private Sub psFillGrvLsSto(vriSOHOID As String)
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        LblMsgLsSto.Text = ""
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgLsSto.Text = pbMsgError
            Exit Sub
        End If

        If ChkLsStoSelectedNot.Checked = True Then
            If Val(DstLsStoBuilding.SelectedValue) = 0 And Val(DstLsStoLantai.SelectedValue) = 0 And Val(DstLsStoZona.SelectedValue) = 0 And Val(DstLsStoStorageType.SelectedValue) = 0 Then
                LblMsgLsSto.Text = "Pilih Building, Lantai, Zona atau Storage Type"
                Exit Sub
            End If
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnCriteria As String = ""

        vnCriteria = vbCrLf & "            and pm.WarehouseOID=" & DstLsStoWhs.SelectedValue
        If Val(DstLsStoBuilding.SelectedValue) > 0 Then
            vnCriteria += vbCrLf & "            and pm.BuildingOID=" & DstLsStoBuilding.SelectedValue
        End If
        If Val(DstLsStoLantai.SelectedValue) > 0 Then
            vnCriteria += vbCrLf & "            and pm.LantaiOID=" & DstLsStoLantai.SelectedValue
        End If
        If Val(DstLsStoZona.SelectedValue) > 0 Then
            vnCriteria += vbCrLf & "            and pm.ZonaOID=" & DstLsStoZona.SelectedValue
        End If
        If Val(DstLsStoStorageType.SelectedValue) > 0 Then
            vnCriteria += vbCrLf & "            and pm.StorageTypeOID=" & DstLsStoStorageType.SelectedValue
        End If

        If DstLsStoStorageType.SelectedValue = enuStorageType.Rack Then
            If Trim(TxtLsStoRackY_SeqNo.Text) <> "" Then
                vnCriteria += vbCrLf & "            and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(TxtLsStoRackY_SeqNo.Text)) & "'"
            End If
            If Trim(TxtLsStoRackY_Column.Text) <> "" Then
                vnCriteria += vbCrLf & "            and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(TxtLsStoRackY_Column.Text)) & "'"
            End If
            If Trim(TxtLsStoRackY_Level.Text) <> "" Then
                vnCriteria += vbCrLf & "            and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(TxtLsStoRackY_Level.Text)) & "'"
            End If

        ElseIf DstLsStoStorageType.SelectedValue = enuStorageType.Floor Then
            If Trim(TxtLsStoRackN_Start.Text) <> "" Then
                vnCriteria += vbCrLf & "            and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(TxtLsStoRackN_Start.Text)) & "'"
            End If
            If Trim(TxtLsStoRackN_End.Text) <> "" Then
                vnCriteria += vbCrLf & "            and isnull(pm.StorageNumber,'')<='" & fbuFormatString(Trim(TxtLsStoRackN_End.Text)) & "'"
            End If

        ElseIf DstLsStoStorageType.SelectedValue = enuStorageType.Staging Then
            vnCriteria += vbCrLf & "            and pm.StorageStagIO=" & RdbLsStoStagging.SelectedValue
        End If

        If ChkLsStoSelectedNot.Checked Then
            vnQuery = "Select * From ("
            vnQuery += vbCrLf & "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
            vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
            vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
            vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
            vnQuery += vbCrLf & "     pm.vStorageOID,pm.vStorageInfoHtml,''vSelect,'Remove'vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
            vnQuery += vbCrLf & "Where 1=1"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and pm.vStorageOID in(Select b.StorageOID From Sys_SsoSOCcStorage_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"

            vnQuery += vbCrLf & "UNION"
            vnQuery += vbCrLf & "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
            vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
            vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
            vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
            vnQuery += vbCrLf & "     pm.vStorageOID,pm.vStorageInfoHtml,'Select'vSelect,''vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
            vnQuery += vbCrLf & "Where 1=1"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and not pm.vStorageOID in(Select b.StorageOID From Sys_SsoSOCcStorage_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"
            vnQuery += vbCrLf & ")b"
            vnQuery += vbCrLf & " Order by b.vStorageInfoHtml"
            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        Else
            vnQuery = "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
            vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
            vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
            vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
            vnQuery += vbCrLf & "     pm.vStorageOID,pm.vStorageInfoHtml,''vSelect,'Remove'vRemove"
            vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
            vnQuery += vbCrLf & "Where 1=1"
            vnQuery += vnCriteria
            vnQuery += vbCrLf & "      and pm.vStorageOID in(Select b.StorageOID From Sys_SsoSOCcStorage_TR b with(nolock) Where b.SOHOID=" & vriSOHOID & ")"
            vnQuery += vbCrLf & " Order by pm.vStorageInfoHtml"
            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        End If

        GrvLsSto.DataSource = vnDtb
        GrvLsSto.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psSetEditable_LsSto()
    End Sub

    Private Sub psSetEditable_LsSto()
        If HdfStockDownload.Value = "1" Then
            GrvLsSto.Columns(ensColLsSto.vSelect).HeaderStyle.CssClass = "myDisplayNone"
            GrvLsSto.Columns(ensColLsSto.vSelect).ItemStyle.CssClass = "myDisplayNone"

            GrvLsSto.Columns(ensColLsSto.vRemove).HeaderStyle.CssClass = "myDisplayNone"
            GrvLsSto.Columns(ensColLsSto.vRemove).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvLsSto.Columns(ensColLsSto.vSelect).HeaderStyle.CssClass = ""
            GrvLsSto.Columns(ensColLsSto.vSelect).ItemStyle.CssClass = ""

            GrvLsSto.Columns(ensColLsSto.vRemove).HeaderStyle.CssClass = ""
            GrvLsSto.Columns(ensColLsSto.vRemove).ItemStyle.CssClass = ""
        End If
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

        Dim vnCrSOType As String = ""
        If ChkListType_02CcLoc.Checked = False And ChkListType_03CcBrg.Checked = False Then
            ChkListType_02CcLoc.Checked = True
            ChkListType_03CcBrg.Checked = True
        End If
        If ChkListType_02CcLoc.Checked = True Then
            vnCrSOType += enuSOType.CcLocation & ","
        End If
        If ChkListType_03CcBrg.Checked = True Then
            vnCrSOType += enuSOType.CcBarang & ","
        End If
        If vnCrSOType <> "" Then
            vnCrSOType = vbCrLf & "      and PM.SOTypeOID in(" & Mid(vnCrSOType, 1, Len(vnCrSOType) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.SONo,sot.SOTypeName,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.SOCutOff,106) + ' '+ convert(varchar(5),PM.SOCutOff,108)vSOCutOff,"
        vnQuery += vbCrLf & "     PM.SOCompanyCode,WM.WarehouseName,SW.SubWhsName,PM.SONote,PM.SOCloseNote,PM.SOCancelNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanOpenDatetime,106)+' '+convert(varchar(5),PM.ScanOpenDatetime,108)+' '+ PR.UserName vScanOpen,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ScanClosedDatetime,106)+' '+convert(varchar(5),PM.ScanClosedDatetime,108)+' '+ AP.UserName vScanClosed,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ CL.UserName vClosed,"
        vnQuery += vbCrLf & "     case when abs(SOStockDownload)=1 then 'Y' else 'N' end vSOStockDownload"

        vnQuery += vbCrLf & "From Sys_SsoSOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.SOWarehouseOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_SubWarehouse_MA SW with(nolock) on SW.OID=PM.SOSubWarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoSOType_MA sot with(nolock) on sot.OID=PM.SOTypeOID and PM.SOTypeOID in(" & enuSOType.CcBarang & "," & enuSOType.CcLocation & ")"
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

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus
        vnQuery += vbCrLf & vnCrSOType

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

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        If vriHOID = 0 Then
            vnQuery = "Select 0 vNo,0 OID,0 vStorageOID,''vStorageInfoHtml,"
            vnQuery += vbCrLf & "       ''BRGCODE,''BRGNAME,''BRGUNIT,0 SOStockQty,0 vSumSOScanQty,0 vSOStockScanVarian,"
            vnQuery += vbCrLf & "       ''vSOStockNote,''vSOStockNoteBy,Null vSOStockNoteDatetime"
            vnQuery += vbCrLf & "Where 1=2"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            GrvDetail.DataSource = vnDtb
            GrvDetail.DataBind()
        Else
            vnQuery = "Select Row_Number()over(order by mb.BRGNAME)vNo,pm.vStorageOID,pm.vStorageInfoHtml,d.OID,"
            vnQuery += vbCrLf & "       d.BRGCODE,mb.BRGNAME,mb.BRGUNIT,d.SOStockQty,d.vSumSOScanQty,d.vSOStockScanVarian,"
            vnQuery += vbCrLf & "       d.vSOStockNote,d.vSOStockNoteBy,d.vSOStockNoteDatetime"
            vnQuery += vbCrLf & "  From fnTbl_SsoSOStockScan_Storage(" & vriHOID & ")d"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm on pm.vStorageOID=d.StorageOID"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.BRGCODE=d.BRGCODE and mb.CompanyCode='" & vnCompanyCode & "'"

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

        LblMsgLsBrg.Text = ""
        LblMsgLsSto.Text = ""
        LblXlsProses.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSODate.Enabled = vriBo
        TxtSONote.ReadOnly = Not vriBo

        If HdfActionStatus.Value = cbuActionNew Then
            DstCompany.Enabled = vriBo
            DstSubWhs.Enabled = vriBo

            DstCutOffHour.Enabled = vriBo
            DstCutOffMin.Enabled = vriBo
        Else
            If HdfActionStatus.Value = cbuActionEdit Then
                DstCompany.Enabled = False
                DstSubWhs.Enabled = False
            Else
                DstCompany.Enabled = True
                DstSubWhs.Enabled = True
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
        vnQuery += vbCrLf & "abs(SOStockDownload)vSOStockDownload,case when abs(SOStockDownload)=1 then 'Y' else 'N' end vSOStockDownload_YN,"
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

            HdfSOCompanyCode.Value = vnDtb.Rows(0).Item("SOCompanyCode")

            HdfSOWarehouseOID.Value = vnDtb.Rows(0).Item("SOWarehouseOID")
            HdfSOWarehouseName.Value = fbuGetWarehouseName(HdfSOWarehouseOID.Value, vriSQLConn)

            pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhs, False, DstCompany.SelectedValue, vriSQLConn)
            DstSubWhs.SelectedValue = vnDtb.Rows(0).Item("SOSubWarehouseOID")

            DstCutOffHour.SelectedValue = Mid(vnDtb.Rows(0).Item("vSOCutOff_HM"), 1, 2)
            DstCutOffMin.SelectedValue = Mid(vnDtb.Rows(0).Item("vSOCutOff_HM"), 4, 2)

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            TxtStockDownload.Text = vnDtb.Rows(0).Item("vSOStockDownload_YN")
            HdfStockDownload.Value = vnDtb.Rows(0).Item("vSOStockDownload")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            RdbSOType.SelectedValue = vnDtb.Rows(0).Item("SOTypeOID")
            HdfSOType.Value = vnDtb.Rows(0).Item("SOTypeOID")

            psSetPanTp_Visibile(RdbSOType.SelectedValue)

            BtnTp02_Loc.Visible = True
            BtnTp03_Brg.Visible = True

            If HdfStockDownload.Value = "1" Then
                BtnTp03_Loc.Visible = True
            Else
                BtnTp03_Loc.Visible = False
            End If

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
        BtnDownloadBrg.Visible = BtnDownloadBrg.Enabled

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

            BtnDownloadBrg.Enabled = (HdfStockDownload.Value = 0)

            ChkLsBrgSelectedNot.Checked = False
            ChkLsBrgSelectedNot.Enabled = (HdfStockDownload.Value = 0)
            ChkLsBrgSelectedNot.Visible = ChkLsBrgSelectedNot.Enabled

            ChkLsStoSelectedNot.Checked = False
            ChkLsStoSelectedNot.Enabled = (HdfStockDownload.Value = 0)
            ChkLsStoSelectedNot.Visible = ChkLsStoSelectedNot.Enabled

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

        If Not vnSave Then Exit Sub

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psSaveBaru", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

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

            Dim vnSOTypeOID As Byte = RdbSOType.SelectedValue
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
            vnQuery += vbCrLf & vnSOTypeOID & ",'" & TxtSODate.Text & " " & DstCutOffHour.SelectedValue & ":" & DstCutOffMin.SelectedValue & "',"
            vnQuery += vbCrLf & "'" & vnCompanyCode & "'," & vnWarehouseOID & "," & vnSubWhsOID & ","
            vnQuery += vbCrLf & "'" & vsXlsFileName & "','" & vsSheetName & "',"
            vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtSONote.Text)) & "',"
            vnQuery += vbCrLf & "'" & stuTransCode.SsoSSOH & "'," & Session("UserOID") & ",getdate(),"
            vnQuery += vbCrLf & "0)"

            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSSOH...Start")
            pbuInsertStatusSSOH(vnOID, enuTCSSOH.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSSOH...End")

            vsTextStream.WriteLine("")
            If vnSOTypeOID = enuSOType.CcLocation Then
                vsTextStream.WriteLine("vnSOTypeOID = enuSOType.CcLocation = " & enuSOType.CcLocation)
            Else
                vsTextStream.WriteLine("vnSOTypeOID = enuSOType.CcBarang = " & enuSOType.CcBarang)
            End If

            vnBeginTrans = False
            vnSQLTrans.Commit()

            Session(csModuleName & stuSession.Simpan) = "Done"

            vnSQLTrans = Nothing

            TxtTransID.Text = vnOID

            TxtStockDownload.Text = "N"
            HdfStockDownload.Value = "0"

            HdfSOCompanyCode.Value = vnCompanyCode

            HdfSOWarehouseOID.Value = vnWarehouseOID
            HdfSOWarehouseName.Value = fbuGetWarehouseName(vnWarehouseOID, vnSQLConn)
            HdfSOType.Value = vnSOTypeOID

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

            psSaveDetail(vnSQLConn, vnSQLTrans)

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

    Private Sub psSaveDetail(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnTxtvSOStockNote As TextBox
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtvSOStockNote = vnGRow.FindControl("TxtvSOStockNote")
            If Trim(vnTxtvSOStockNote.Text) <> Replace(fbuValStrHtml(vnGRow.Cells(ensColDetail.vSOStockNote).Text), "<br />", vbLf) Then
                vnQuery = "Update Sys_SsoSOStock_TR set "
                vnQuery += vbCrLf & "SOStockNote='" & fbuFormatString(vnTxtvSOStockNote.Text) & "',"
                vnQuery += vbCrLf & "SOStockNoteUserOID='" & Session("UserOID") & "',SOStockNoteDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                vnQuery = "Insert into Sys_SsoSOStock_HS "
                vnQuery += vbCrLf & "(SOSOID,SOHOID,StorageOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty,SOStockNote,SOStockNoteUserOID,SOStockNoteDatetime)"
                vnQuery += vbCrLf & "Select OID,SOHOID,StorageOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty,SOStockNote,SOStockNoteUserOID,SOStockNoteDatetime"
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
        ElseIf HdfProcess.Value = "DownloadStock" Then
            psDownloadStock()
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
        vnDtb.Rows.Add(New Object() {stuSsoReportType.RptCycleCount, "Cycle Count Storage - Barang"})

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
            ElseIf DstProReport.SelectedValue = stuSsoReportType.RptCycleCount Then
                psGenerateCrpTally_CycleCount(vnCrpFileName)
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
        vriCrpFileName = stuSsoCrp.CrpSsoSOTally_Storage

        vbuCrpQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,pm.vStorageOID,pm.vStorageInfo,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianOnly.Checked, 1, 0) & " vVarianOnly"
        vbuCrpQuery += "       From fnTbl_SsoTally_Storage(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm on pm.vStorageOID=ta.StorageOID"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        If ChkProVarianOnly.Checked Then
            vbuCrpQuery += "       Where ta.vSOStockScanVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME"
    End Sub
    Private Sub psGenerateCrpTally_CycleCount(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOCycleCount

        vbuCrpQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,pm.vStorageOID,pm.vStorageInfo,"
        vbuCrpQuery += "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
        vbuCrpQuery += "            " & IIf(ChkProVarianOnly.Checked, 1, 0) & " vVarianOnly"

        vbuCrpQuery += "       From fnTbl_SsoTally_Storage(" & TxtTransID.Text & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm on pm.vStorageOID=ta.StorageOID"
        vbuCrpQuery += "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE and abs(mb.IsActive)=1"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
        vbuCrpQuery += "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
        If ChkProVarianOnly.Checked Then
            vbuCrpQuery += "       Where ta.vSOStockScanVarian!=0"
        End If
        vbuCrpQuery += " order by mb.BRGNAME"
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

    Protected Sub RdbSOType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdbSOType.SelectedIndexChanged
        If HdfActionStatus.Value = cbuActionNorm Then Exit Sub
        psSetPanTp_Visibile(RdbSOType.SelectedValue)
    End Sub

    Private Sub psSetPanTp_Visibile(vriSOType As Byte)
        If vriSOType = 1 Then
            PanTp02_CcLoc.Visible = False
            PanTp03_CcBrg.Visible = False
        ElseIf vriSOType = 2 Then
            PanTp02_CcLoc.Visible = True
            PanTp03_CcBrg.Visible = False
        ElseIf vriSOType = 3 Then
            PanTp02_CcLoc.Visible = False
            PanTp03_CcBrg.Visible = True
        End If
    End Sub

    Protected Sub BtnTp03_Brg_Click(sender As Object, e As EventArgs) Handles BtnTp03_Brg.Click
        psShowLsBrg(True)
    End Sub

    Private Sub GrvLsBrg_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsBrg.PageIndexChanging
        GrvLsBrg.PageIndex = e.NewPageIndex
        psFillGrvLsBrg(Val(TxtTransID.Text))
    End Sub

    Private Sub GrvLsBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsBrg.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvLsBrg.Rows(vnRowIdx)
        Dim vnBrgCode As String = fbuFormatString(vnGRow.Cells(ensColLsBrg.BRGCODE).Text)

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            Dim vnSOHOID As Integer = Val(TxtTransID.Text)
            Dim vnQuery As String
            Dim vnBrgExist As Byte
            vnQuery = "Select 1 From Sys_SsoSOCcBarang_TR with(nolock) Where SOHOID=" & vnSOHOID & " and BRGCODE='" & vnBrgCode & "'"
            vnBrgExist = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction("upd")
            vnBeginTrans = True
            If e.CommandName = "vSelect" Then
                If vnBrgExist = 0 Then
                    vnQuery = "Insert into Sys_SsoSOCcBarang_TR values(" & vnSOHOID & ",'" & vnBrgCode & "')"
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                End If
                DirectCast(vnGRow.Cells(ensColLsBrg.vSelect).Controls(0), LinkButton).Text = ""
                DirectCast(vnGRow.Cells(ensColLsBrg.vRemove).Controls(0), LinkButton).Text = "Remove"
            Else
                If vnBrgExist = 1 Then
                    vnQuery = "Delete Sys_SsoSOCcBarang_TR Where SOHOID=" & vnSOHOID & " and BrgCode='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)
                End If
                DirectCast(vnGRow.Cells(ensColLsBrg.vSelect).Controls(0), LinkButton).Text = "Select"
                DirectCast(vnGRow.Cells(ensColLsBrg.vRemove).Controls(0), LinkButton).Text = ""
            End If

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
            vnBeginTrans = False

        Catch ex As Exception
            LblMsgLsBrg.Text = ex.Message
            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub GrvLsBrg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsBrg.SelectedIndexChanged

    End Sub

    Private Sub BtnLsStoClose_Click(sender As Object, e As EventArgs) Handles BtnLsStoClose.Click
        psShowLsSto(False)
    End Sub

    Private Sub BtnTp02_Loc_Click(sender As Object, e As EventArgs) Handles BtnTp02_Loc.Click
        psShowLsSto(True)
    End Sub

    Private Sub BtnLsSto_Click(sender As Object, e As EventArgs) Handles BtnLsSto.Click
        If HdfActionStatus.Value = cbuActionNew Then Exit Sub
        psClearMessage()

        psFillGrvLsSto(Val(TxtTransID.Text))
    End Sub

    Private Sub GrvLsSto_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsSto.PageIndexChanging
        GrvLsSto.PageIndex = e.NewPageIndex
        psFillGrvLsSto(Val(TxtTransID.Text))
    End Sub

    Private Sub GrvLsSto_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsSto.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvLsSto.Rows(vnRowIdx)
        Dim vnStorageOID As String = fbuFormatString(vnGRow.Cells(ensColLsSto.vStorageOID).Text)

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            Dim vnSOHOID As Integer = Val(TxtTransID.Text)
            Dim vnQuery As String
            Dim vnStoExist As Byte
            vnQuery = "Select 1 From Sys_SsoSOCcStorage_TR with(nolock) Where SOHOID=" & vnSOHOID & " and StorageOID=" & vnStorageOID
            vnStoExist = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction("upd")
            vnBeginTrans = True
            If e.CommandName = "vSelect" Then
                If vnStoExist = 0 Then
                    vnQuery = "Insert into Sys_SsoSOCcStorage_TR values(" & vnSOHOID & "," & vnStorageOID & ")"
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                End If
                DirectCast(vnGRow.Cells(ensColLsSto.vSelect).Controls(0), LinkButton).Text = ""
                DirectCast(vnGRow.Cells(ensColLsSto.vRemove).Controls(0), LinkButton).Text = "Remove"
            Else
                If vnStoExist = 1 Then
                    vnQuery = "Delete Sys_SsoSOCcStorage_TR Where SOHOID=" & vnSOHOID & " and StorageOID=" & vnStorageOID
                    pbuExecuteSQLTrans(vnQuery, cbuActionDel, vnSQLConn, vnSQLTrans)
                End If
                DirectCast(vnGRow.Cells(ensColLsSto.vSelect).Controls(0), LinkButton).Text = "Select"
                DirectCast(vnGRow.Cells(ensColLsSto.vRemove).Controls(0), LinkButton).Text = ""
            End If

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
            vnBeginTrans = False

        Catch ex As Exception
            LblMsgLsSto.Text = ex.Message
            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Private Sub psSetUpLsSto()
        If HdfSetupLsSto.Value = "1" Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuFillDstStorageType(DstLsStoStorageType, True, vnSQLConn)
        pbuFillDstLantai(DstLsStoLantai, True, vnSQLConn)
        pbuFillDstZona(DstLsStoZona, True, vnSQLConn)

        pbuFillDstWarehouse_Pr(HdfSOWarehouseOID.Value, DstLsStoWhs, False, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        HdfSetupLsSto.Value = "1"
    End Sub

    Private Sub DstLsStoStorageType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstLsStoStorageType.SelectedIndexChanged
        If DstLsStoStorageType.SelectedValue = enuStorageType.Floor Then
            PanListRackN.Visible = True
            PanListRackY.Visible = False
            PanListStagging.Visible = False
        ElseIf DstLsStoStorageType.SelectedValue = enuStorageType.Rack Then
            PanListRackN.Visible = False
            PanListRackY.Visible = True
            PanListStagging.Visible = False
        ElseIf DstLsStoStorageType.SelectedValue = enuStorageType.Staging Then
            PanListRackN.Visible = False
            PanListRackY.Visible = False
            PanListStagging.Visible = True
        Else
            PanListRackN.Visible = False
            PanListRackY.Visible = False
            PanListStagging.Visible = False
        End If
    End Sub
    Private Sub psDownloadStock()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnSOHOID As Integer = Val(TxtTransID.Text)
            Dim vnQuery As String
            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery = "Select count(1) From Sys_SsoSOCcBarang_TR with(nolock) Where SOHOID=" & vnSOHOID
            Else
                vnQuery = "Select count(1) From Sys_SsoSOCcStorage_TR with(nolock) Where SOHOID=" & vnSOHOID
            End If

            If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                LblMsgError.Text = "Belum Ada Barang atau Lokasi yang dipilih...Download Gagal."
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnWhsOID As String = HdfSOWarehouseOID.Value
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoSOStock_TR("
            vnQuery += vbCrLf & "SOHOID,StorageOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
            vnQuery += vbCrLf & "Select " & vnSOHOID & ",sc.StorageOID,sc.BRGCODE,mb.BRGNAME,mb.BRGUNIT,sum(sc.QtyOnHand) SOStockQty"
            vnQuery += vbCrLf & "  From Sys_SsoStorageStock_MA sc with(nolock)"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=sc.BRGCODE and mb.CompanyCode=sc.CompanyCode and mb.CompanyCode='" & HdfSOCompanyCode.Value & "'"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) ms on ms.vStorageOID=sc.StorageOID and"
            If vnWhsOID = enuWarehouseOID.Prancis Or vnWhsOID = enuWarehouseOID.Prancis2 Then
                vnQuery += vbCrLf & "                  ms.WarehouseOID in(" & enuWarehouseOID.Prancis & "," & enuWarehouseOID.Prancis2 & ")"
            Else
                vnQuery += vbCrLf & "                  ms.WarehouseOID=" & vnWhsOID
            End If
            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery += vbCrLf & "       inner join Sys_SsoSOCcBarang_TR sob with(nolock) on sob.BRGCODE=sc.BRGCODE"
            Else
                vnQuery += vbCrLf & "       inner join Sys_SsoSOCcStorage_TR sob with(nolock) on sob.StorageOID=sc.StorageOID"
            End If
            vnQuery += vbCrLf & " Where sob.SOHOID=" & vnSOHOID
            vnQuery += vbCrLf & " Group by sc.StorageOID,sc.BRGCODE,mb.BRGNAME,mb.BRGUNIT"
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery = "Insert into Sys_SsoSOCcBarangStorage_TR(SOHOID,BRGCODE,StorageOID)"
                vnQuery += vbCrLf & "Select distinct " & vnSOHOID & ",BrgCode,StorageOID From Sys_SsoStorageStock_MA sc with(nolock)"
                vnQuery += vbCrLf & " Where BrgCode in(Select b.BrgCode From Sys_SsoSOCcBarang_TR b with(nolock) Where b.SOHOID=" & vnSOHOID & ")"
                vnQuery += vbCrLf & "       and StorageOID in(Select pm.vStorageOID From " & vnDBMaster & "fnTbl_SsoStorageInfo(0) pm"
                If vnWhsOID = enuWarehouseOID.Prancis Or vnWhsOID = enuWarehouseOID.Prancis2 Then
                    vnQuery += vbCrLf & "                         Where pm.WarehouseOID in(" & enuWarehouseOID.Prancis & "," & enuWarehouseOID.Prancis2 & ")  )"
                Else
                    vnQuery += vbCrLf & "                         Where pm.WarehouseOID=" & vnWhsOID & ")"
                End If

                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            vnQuery = "Update Sys_SsoSOHeader_TR set SOCutOff=getdate(),SOStockDownload=1,"
            vnQuery += vbCrLf & "SOStockDownloadUserOID=" & Session("UserOID") & ",SOStockDownloadDatetime=getdate() Where OID=" & vnSOHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

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

    Private Sub psDownloadStock_20230827_Bef_Pra1_Sama_Dengan_Pra2()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnSOHOID As Integer = Val(TxtTransID.Text)
            Dim vnQuery As String
            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery = "Select count(1) From Sys_SsoSOCcBarang_TR with(nolock) Where SOHOID=" & vnSOHOID
            Else
                vnQuery = "Select count(1) From Sys_SsoSOCcStorage_TR with(nolock) Where SOHOID=" & vnSOHOID
            End If

            If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                LblMsgError.Text = "Belum Ada Barang atau Lokasi yang dipilih...Download Gagal."
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Insert into Sys_SsoSOStock_TR("
            vnQuery += vbCrLf & "SOHOID,StorageOID,BRGCODE,BRGNAME,BRGUNIT,SOStockQty)"
            vnQuery += vbCrLf & "Select " & vnSOHOID & ",sc.StorageOID,sc.BRGCODE,mb.BRGNAME,mb.BRGUNIT,sum(sc.QtyOnHand) SOStockQty"
            vnQuery += vbCrLf & "  From Sys_SsoStorageStock_MA sc with(nolock)"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=sc.BRGCODE and mb.CompanyCode=sc.CompanyCode and mb.CompanyCode='" & HdfSOCompanyCode.Value & "'"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) ms on ms.vStorageOID=sc.StorageOID and ms.WarehouseOID=" & HdfSOWarehouseOID.Value
            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery += vbCrLf & "       inner join Sys_SsoSOCcBarang_TR sob with(nolock) on sob.BRGCODE=sc.BRGCODE"
            Else
                vnQuery += vbCrLf & "       inner join Sys_SsoSOCcStorage_TR sob with(nolock) on sob.StorageOID=sc.StorageOID"
            End If
            vnQuery += vbCrLf & " Where sob.SOHOID=" & vnSOHOID
            vnQuery += vbCrLf & " Group by sc.StorageOID,sc.BRGCODE,mb.BRGNAME,mb.BRGUNIT"
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            If HdfSOType.Value = enuSOType.CcBarang Then
                vnQuery = "Insert into Sys_SsoSOCcBarangStorage_TR(SOHOID,BRGCODE,StorageOID)"
                vnQuery += vbCrLf & "Select distinct " & vnSOHOID & ",BrgCode,StorageOID From Sys_SsoStorageStock_MA sc with(nolock)"
                vnQuery += vbCrLf & " Where BrgCode in(Select b.BrgCode From Sys_SsoSOCcBarang_TR b with(nolock) Where b.SOHOID=" & vnSOHOID & ")"
                vnQuery += vbCrLf & "       and StorageOID in(Select pm.vStorageOID From " & vnDBMaster & "fnTbl_SsoStorageInfo(0) pm Where pm.WarehouseOID=" & HdfSOWarehouseOID.Value & ")"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            vnQuery = "Update Sys_SsoSOHeader_TR set SOCutOff=getdate(),SOStockDownload=1,"
            vnQuery += vbCrLf & "SOStockDownloadUserOID=" & Session("UserOID") & ",SOStockDownloadDatetime=getdate() Where OID=" & vnSOHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

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

    Protected Sub BtnDownloadBrg_Click(sender As Object, e As EventArgs) Handles BtnDownloadBrg.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Download Stock " & TxtSONo.Text & " ?<br />WARNING : Download Stock Tidak Dapat Dibatalkan"
        HdfProcess.Value = "DownloadStock"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Protected Sub BtnTp03_Loc_Click(sender As Object, e As EventArgs) Handles BtnTp03_Loc.Click
        psShowLsBStg(True)
    End Sub

    Protected Sub BtnLsBStgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBStgClose.Click
        psShowLsBStg(False)
    End Sub

    Protected Sub BtnLsBStg_Click(sender As Object, e As EventArgs) Handles BtnLsBStg.Click
        psFillGrvLsBStg(Val(TxtTransID.Text))
    End Sub

    Protected Sub DstLsStoWhs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstLsStoWhs.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuFillDstBuilding_ByWarehouse(DstLsStoBuilding, True, DstLsStoWhs.SelectedValue, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class