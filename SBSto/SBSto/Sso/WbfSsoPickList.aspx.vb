Imports System.Data.SqlClient
Public Class WbfSsoPickList
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoPickList"
    Const csTNoPrefix = "PICK"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
    End Enum
    Enum ensColListDoc
        CompanyCode = 0
        no_nota = 1
        vtanggal = 2
        kode_cust = 3
        CUSTOMER = 4
        ALAMAT = 5
        kota = 6
        WarehouseOID = 7
    End Enum
    Enum ensColListTRB
        OID = 0
        CompanyCode = 1
        NoBukti = 2
        vTanggal = 3
        GudangAsal = 4
        GudangTujuan = 5
        WarehouseAsalOID = 6
        WarehouseTujuanOID = 7
    End Enum

    Enum ensColListPKDOT
        OID = 0
        PKDOTCompanyCode = 1
        PKDOTNo = 2
        vPKDOTScheduleDate = 3
        WarehouseName = 4
        vWarehouseName_Dest = 5
        NotaHOID = 6
        NotaNo = 7
        vCustomer = 8
        PKDOTDate = 9
        WarehouseOID = 10
        WarehouseOID_Dest = 11
    End Enum
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPickList, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstWhs, False, vnSQLConn)
            pbuFillDstWarehouse(DstWhsDest, False, vnSQLConn)
            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPCLScheduleDate.Text = ""
        TxtPCLNo.Text = ""
        TxtPCLDescr.Text = ""
        TxtPCLPrint.Text = ""

        TxtPCLRefNo.Text = ""
        HdfPCLRefOID.Value = "0"
        TxtPCLRefOID.Text = ""

        HdfPrioritas.Value = "0"
        LblPrioritas.Text = ""

        RdlPickType.Items(0).Selected = False
        RdlPickType.Items(1).Selected = False
        RdlPickType.Items(2).Selected = False
        RdlPickType.Items(3).Selected = False

        ChkDest.Visible = False

        TxtPCLScheduleDate.Text = ""
        TxtPickingNo.Text = ""
        TxtPickingStatus.Text = ""
        HdfPickingHOID.Value = "0"
        HdfPickingStatus.Value = enuTCPCKG.None
        HdfEnableVoid.Value = "0"

        HdfPCLStorageTypeList.Value = ""
        HdfCompanyCode.Value = "0"
        HdfWhs.Value = "0"

        ChkPCL_Rack.Checked = False
        ChkPCL_Floor.Checked = False
        ChkPCL_CrossDock.Checked = False
        ChkPCL_DOTitip.Checked = False

        HdfTransStatus.Value = enuTCPICK.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListDoc.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListDoc.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListTRB.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListPKDOT.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListPKDOT.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrintHS.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrintHS.Style(HtmlTextWriterStyle.Position) = "absolute"
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

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_OnPicking.Checked = False And ChkSt_Prepared.Checked = False And ChkSt_PickingDone.Checked = False And ChkSt_Void.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_Prepared.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCPICK.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPICK.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCPICK.Prepared & ","
        End If
        If ChkSt_OnPicking.Checked = True Then
            vnCrStatus += enuTCPICK.On_Picking & ","
        End If
        If ChkSt_PickingDone.Checked = True Then
            vnCrStatus += enuTCPICK.Picking_Done & ","
        End If
        If ChkSt_Void.Checked = True Then
            vnCrStatus += enuTCPICK.Void & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.PCLNo,convert(varchar(11),PM.PCLDate,106)vPCLDate,convert(varchar(11),PM.PCLScheduleDate,106)vPCLScheduleDate,"
        vnQuery += vbCrLf & "     PM.PCLCompanyCode,TP.SchDTypeName,PM.PCLRefHOID,PM.PCLRefHNo,"
        vnQuery += vbCrLf & "     case when PM.SchDTypeOID=1 then case when abs(nh.NotaPRIO)=1 then 'Y' else 'N' end else 'N' end vInvoicePrio,"
        vnQuery += vbCrLf & "     PM.PCLRefHInfo,pck.PCKNo,WM.WarehouseName,whd.WarehouseName vWarehouseName_Dest,"
        vnQuery += vbCrLf & "     mds.DestTypeName,PM.PCLNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared,"
        vnQuery += vbCrLf & "     PCLCancelNote,PCLVoidNote"

        vnQuery += vbCrLf & "From Sys_SsoPCLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=PM.WarehouseOID_Dest"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmSchDType_MA TP with(nolock) on TP.OID=PM.SchDTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoDestType_MA mds with(nolock) on mds.OID=PM.DestTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     left outer join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=PM.OID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock) on nh.OID=PM.PCLRefHOID and PM.SchDTypeOID=1"

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
            vnQuery += vbCrLf & " and PM.PCLNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If
        If Trim(TxtListRefNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.PCLRefHNo like '%" & fbuFormatString(Trim(TxtListRefNo.Text)) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.PCLCompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWhs.SelectedValue
        End If
        If RdlListPickType.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.SchDTypeOID = " & RdlListPickType.SelectedValue
        End If
        If ChkCrPrioritas.Checked Then
            vnQuery += vbCrLf & "            and case when PM.SchDTypeOID=1 then case when abs(nh.NotaPRIO)=1 then 'Y' else 'N' end else 'N' end='Y'"
        End If
        vnQuery += vbCrLf & "Order by PM.CreationDatetime Desc"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvDetail(vriHOID As Integer, vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String = ""

        If HdfActionStatus.Value = cbuActionNorm Then
            If vriHOID = 0 Then
                vnQuery = "Select ''BRGCODE,''BRGNAME,0 RefQty,0 PCLDQty,''BRGUNIT where 1=2"
            Else
                vnQuery = "Select dt.BRGCODE,mb.BRGNAME,RefQty,PCLDQty,BRGUNIT"
                vnQuery += vbCrLf & "       From Sys_SsoPCLDetail_TR dt with(nolock)"
                vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.BRGCODE"
                vnQuery += vbCrLf & "      Where PCLHOID=" & vriHOID & " and mb.CompanyCode='" & HdfCompanyCode.Value & "'"
                vnQuery += vbCrLf & "   Order by dt.BRGCODE"
            End If
        Else
            If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
                '<---19 Sep 2023 Barang Paketan
                'vnQuery = "Select dt.KodeBarang BRGCODE,mb.BRGNAME,sum(dt.Qty + dt.QtyBonus - dt.QtyOnPickList) RefQty,0 PCLDQty,mb.BRGUNIT"
                'vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmNotaHeader_TR hd with(nolock)"
                'vnQuery += vbCrLf & "            inner join " & vnDBDcm & "Sys_DcmNotaDetail_TR dt with(nolock) on dt.NotaHOID=hd.OID"
                'vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.KodeBarang and mb.CompanyCode=hd.CompanyCode"
                vnQuery = "Select dt.KodeBarang BRGCODE,isnull(mb.BRGNAME,dt.NamaBarang)BRGNAME,sum(dt.Qty + dt.QtyBonus - dt.QtyOnPickList) RefQty,0 PCLDQty,mb.BRGUNIT"
                vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmNotaHeader_TR hd with(nolock)"
                vnQuery += vbCrLf & "            inner join " & vnDBDcm & "Sys_DcmNotaDetail_TR dt with(nolock) on dt.NotaHOID=hd.OID"
                vnQuery += vbCrLf & "            left outer join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.KodeBarang and mb.CompanyCode=hd.CompanyCode"

                vnQuery += vbCrLf & "      Where dt.NotaHOID=" & vriHOID
                vnQuery += vbCrLf & "   Group by dt.KodeBarang,isnull(mb.BRGNAME,dt.NamaBarang),mb.BRGUNIT"
                vnQuery += vbCrLf & "  Having sum(dt.Qty + dt.QtyBonus - dt.QtyOnPickList)>0"
                vnQuery += vbCrLf & "   Order by dt.KodeBarang"
                '<<==19 Sep 2023 Barang Paketan

            ElseIf HdfPickTypeOID.Value = enuSchDType.TRB Then
                vnQuery = "Select dt.KodeBrg BRGCODE,mb.BRGNAME,sum(dt.Qty - dt.QtyOnPickList) RefQty,0 PCLDQty,mb.BRGUNIT"
                vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmTRBHeader_TR hd with(nolock)"
                vnQuery += vbCrLf & "            inner join " & vnDBDcm & "Sys_DcmTRBDetail_TR dt with(nolock) on dt.TRBHOID=hd.OID"
                vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.KodeBrg and mb.CompanyCode=hd.CompanyCode"
                vnQuery += vbCrLf & "      Where dt.TRBHOID=" & vriHOID
                vnQuery += vbCrLf & "   Group by dt.KodeBrg,mb.BRGNAME,mb.BRGUNIT"
                vnQuery += vbCrLf & "  Having sum(dt.Qty - dt.QtyOnPickList)>0"
                vnQuery += vbCrLf & "   Order by dt.KodeBrg"

            ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
                vnQuery = "Select dt.BRGCODE,mb.BRGNAME,sum(dt.PKDOTDQty - dt.QtyOnPickList) RefQty,0 PCLDQty,mb.BRGUNIT"
                vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmPKDOTHeader_TR hd with(nolock)"
                vnQuery += vbCrLf & "            inner join " & vnDBDcm & "Sys_DcmPKDOTDetail_TR dt with(nolock) on dt.PKDOTHOID=hd.OID"
                vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.BRGCODE and mb.CompanyCode=hd.PKDOTCompanyCode"
                vnQuery += vbCrLf & "      Where dt.PKDOTHOID=" & vriHOID
                vnQuery += vbCrLf & "   Group by dt.BRGCODE,mb.BRGNAME,mb.BRGUNIT"
                vnQuery += vbCrLf & "  Having sum(dt.PKDOTDQty - dt.QtyOnPickList)>0"
                vnQuery += vbCrLf & "   Order by dt.BRGCODE"
            End If
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()
    End Sub

    Private Sub psFillGrvInv(vriHOID As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If vriHOID = 0 Then
            LblInv.Text = ""
            vnQuery = "Select ''KodeBarang,''NamaBarang,0 Qty,0 QtyBonus where 1=2"
        Else
            LblInv.Text = "DATA INVOICE " & TxtPCLRefNo.Text
            vnQuery = "Select KodeBarang,NamaBarang,Qty,QtyBonus From " & vnDBDcm & "Sys_DcmNotaDetail_TR"
            vnQuery += vbCrLf & "Where NotaHOID=" & vriHOID
            vnQuery += vbCrLf & "Order by KodeBarang"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvInv.DataSource = vnDtb
        GrvInv.DataBind()
    End Sub

    Private Sub psFillGrvReserve(vriHOID As String, vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If vriHOID = 0 Then
            vnQuery = "Select ''BRGCODE,''BRGNAME,0 ReservedQty,0 QtyOnPicking,''RcvPONo,''vStorageInfoHtml,0 StorageOID,0 StorageStockOID where 1=2"
        Else
            vnQuery = "Select pcs.BRGCODE,mbr.BRGNAME,pcs.ReservedQty,pcs.QtyOnPicking,"
            vnQuery += vbCrLf & "      rch.RcvPONo,convert(varchar(11),rch.RcvPODate,106)vRcvPODate,sti.vStorageInfoHtml,pcs.StorageOID,pcs.StorageStockOID"
            vnQuery += vbCrLf & " From Sys_SsoPCLReserve_TR pcs with(nolock)"
            vnQuery += vbCrLf & "      inner join Sys_SsoStorageStock_MA sto with(nolock) on sto.OID=pcs.StorageStockOID"
            vnQuery += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=sto.RcvPOHOID"
            vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mbr with(nolock) on mbr.BRGCODE=pcs.BRGCODE and mbr.CompanyCode='" & HdfCompanyCode.Value & "'"
            vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') sti on sti.vStorageOID=pcs.StorageOID"
            vnQuery += vbCrLf & "Where pcs.PCLHOID=" & vriHOID
            vnQuery += vbCrLf & "Order by pcs.BRGCODE"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvReserved.DataSource = vnDtb
        GrvReserved.DataBind()

        LblMsgReserved.Text = "DETAIL STOCK PICK LIST"
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

    Private Sub psShowPrintHS(vriBo As Boolean)
        If vriBo Then
            DivPrintHS.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPrintHS.Style(HtmlTextWriterStyle.Visibility) = "hidden"
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
        vnQuery += vbCrLf & "        + replicate(0,4-len(isnull(max(convert(int,substring(PCLNo,len(PCLNo)-3,4))),0)+1))"
        vnQuery += vbCrLf & "        + cast(isnull(max(convert(int,substring(PCLNo,len(PCLNo)-3,4))),0)+1 as varchar)"
        vnQuery += vbCrLf & "        From Sys_SsoPCLHeader_TR with(nolock)"
        vnQuery += vbCrLf & "       Where substring(PCLNo,1,len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'))="
        vnQuery += vbCrLf & "                                     '" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/'"
        vnQuery += vbCrLf & "                                 and len(PCLNo)=len('" & csTNoPrefix & "/" & vriCompanyCode & "/" & vriWarehouseCode & "/'+substring(convert(varchar(10),getdate(),111),3,5)+'/')+4"
        TxtPCLNo.Text = fbuGetDataStrSQL(vnQuery, vriSQLConn)
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

        ChkPCL_Rack.Checked = True
        ChkPCL_Floor.Checked = True
        ChkPCL_CrossDock.Checked = True
        ChkPCL_DOTitip.Checked = False

        RdlPickType.Items(0).Selected = False
        RdlPickType.Items(1).Selected = False
        RdlPickType.Items(2).Selected = False
        RdlPickType.Items(3).Selected = False

        If DstCompany.Items.Count > 0 Then
            DstCompany.SelectedIndex = 0
        End If

        TxtPCLScheduleDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)
        psFillGrvReserve(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgPCLScheduleDate.Text = ""
        LblMsgWhs.Text = ""
        LblMsgWhsDest.Text = ""
        LblMsgPCLRefNo.Text = ""
        LblMsgPCL_Pick.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPCLNo.ReadOnly = Not vriBo
        TxtPCLScheduleDate.ReadOnly = Not vriBo
        TxtPCLDescr.ReadOnly = Not vriBo
        RdlPickType.Enabled = vriBo

        BtnPCLRefNo.Enabled = False
        BtnPCLRefNo.Visible = BtnPCLRefNo.Enabled
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoPickList
            vnParam += "&vqTrNo=" & TxtPCLNo.Text

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
            psFillGrvReserve(0, vnSQLConn)

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
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        Dim vnTransOID As String = TxtTransID.Text
        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,convert(varchar(11),PM.PCLScheduleDate,106)vPCLScheduleDate,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoPCLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     Where PM.OID=" & vnTransOID
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
            psFillGrvDetail(0, vriSQLConn)
            psFillGrvReserve(0, vriSQLConn)
            psFillGrvInv(0, vriSQLConn)
        Else
            TxtPCLNo.Text = vnDtb.Rows(0).Item("PCLNo")

            TxtPCLRefNo.Text = vnDtb.Rows(0).Item("PCLRefHNo")
            TxtPCLDescr.Text = vnDtb.Rows(0).Item("PCLRefHInfo")
            HdfPCLRefOID.Value = vnDtb.Rows(0).Item("PCLRefHOID")
            TxtPCLRefOID.Text = vnDtb.Rows(0).Item("PCLRefHOID")

            TxtPCLScheduleDate.Text = vnDtb.Rows(0).Item("vPCLScheduleDate")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")
            TxtPCLPrint.Text = vnDtb.Rows(0).Item("PCLPrintNo")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("PCLCompanyCode"))
            HdfCompanyCode.Value = DstCompany.SelectedValue

            DstWhs.SelectedValue = Trim(vnDtb.Rows(0).Item("WarehouseOID"))
            HdfWhs.Value = DstWhs.SelectedValue

            RdlPickType.SelectedValue = vnDtb.Rows(0).Item("SchDTypeOID")
            HdfPickTypeOID.Value = RdlPickType.SelectedValue

            HdfPrioritas.Value = "0"
            LblPrioritas.Text = ""
            If RdlPickType.SelectedValue = enuSchDType.TRB Then
                DstWhsDest.SelectedValue = Trim(vnDtb.Rows(0).Item("WarehouseOID_Dest"))
                HdfWhsDest.Value = DstWhsDest.SelectedValue
            ElseIf RdlPickType.SelectedValue = enuSchDType.Invoice Then
                DstWhsDest.SelectedIndex = -1
                HdfWhsDest.Value = "0"

                vnQuery = "Select abs(nh.NotaPRIO) From " & fbuGetDBDcm() & "Sys_DcmNotaHeader_TR nh with(nolock) where nh.OID=" & vnDtb.Rows(0).Item("PCLRefHOID")
                HdfPrioritas.Value = fbuGetDataNumSQL(vnQuery, vriSQLConn)
                If HdfPrioritas.Value = "1" Then
                    LblPrioritas.Text = "URGENT"
                End If
            Else
                DstWhsDest.SelectedIndex = -1
                HdfWhsDest.Value = "0"
            End If

            ChkDest.Visible = (RdlPickType.SelectedValue = enuSchDType.Invoice)
            ChkDest.Checked = (vnDtb.Rows(0).Item("DestTypeOID") = enuDestType.Luar_Kota)

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            HdfPCLStorageTypeList.Value = vnDtb.Rows(0).Item("PCLStorageTypeList")

            ChkPCL_Rack.Checked = (InStr(HdfPCLStorageTypeList.Value, enuStorageType.Rack) > 0)
            ChkPCL_Floor.Checked = (InStr(HdfPCLStorageTypeList.Value, enuStorageType.Floor) > 0)
            ChkPCL_CrossDock.Checked = (InStr(HdfPCLStorageTypeList.Value, enuStorageType.CrossDock) > 0)
            ChkPCL_DOTitip.Checked = (InStr(HdfPCLStorageTypeList.Value, enuStorageType.Staging) > 0)

            psEnableVoid(vnTransOID, HdfTransStatus.Value, vriSQLConn)

            psButtonStatus()
            psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)
            psFillGrvReserve(Val(TxtTransID.Text), vriSQLConn)

            If RdlPickType.SelectedValue = enuSchDType.Invoice Or RdlPickType.SelectedValue = enuSchDType.DO_Titip Then
                psFillGrvInv(Val(HdfPCLRefOID.Value), vriSQLConn)
            Else
                psFillGrvInv(0, vriSQLConn)
            End If
        End If
        vnDtb.Dispose()
    End Sub

    Private Sub psEnableVoid(vriTransOID As Integer, vriTransStatus As Integer, vriSQLConn As SqlConnection)
        If vriTransStatus = enuTCPICK.Cancelled Or vriTransStatus = enuTCPICK.Baru Or vriTransStatus = enuTCPICK.Prepared Then
            HdfPickingHOID.Value = "0"
            HdfPickingStatus.Value = enuTCPCKG.None
            HdfEnableVoid.Value = "0"
            TxtPickingStatus.Text = ""
        Else
            Dim vnDtbPCK As New DataTable
            pbuGetDtbPCKHOID_By_PickListHOID(vnDtbPCK, vriTransOID, vriSQLConn)
            If vnDtbPCK.Rows.Count = 0 Then
                HdfPickingHOID.Value = "0"
                HdfPickingStatus.Value = enuTCPCKG.None
                TxtPickingNo.Text = ""
                TxtPickingStatus.Text = ""
            Else
                HdfPickingHOID.Value = vnDtbPCK.Rows(0).Item("OID")
                HdfPickingStatus.Value = vnDtbPCK.Rows(0).Item("TransStatus")
                TxtPickingNo.Text = vnDtbPCK.Rows(0).Item("PCKNo")
                TxtPickingStatus.Text = vnDtbPCK.Rows(0).Item("TransStatusDescr")
            End If
            vnDtbPCK.Dispose()
            vnDtbPCK = Nothing

            If vriTransStatus = enuTCPICK.Picking_Done Then
                '03 Oct 2023, kasus Picking udah move dari stg out lt 4, ke stg out bs, ternyata ada info salah harga, shg harus void
                'If HdfPickingStatus.Value = enuTCPCKG.Picking_Done Then
                If HdfPickingStatus.Value = enuTCPCKG.Picking_Done Or HdfPickingStatus.Value = enuTCPCKG.Move_Antar_StagingOut_Done Then
                    HdfEnableVoid.Value = "1"
                Else
                    HdfEnableVoid.Value = "0"
                End If
            Else
                HdfEnableVoid.Value = "0"
            End If
        End If
        BtnVoidPCL.Enabled = (HdfEnableVoid.Value = "1")
        BtnVoidPCL.Visible = BtnVoidPCL.Enabled
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
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCPICK.Baru)

            BtnCancelPCL.Enabled = (HdfTransStatus.Value = enuTCPICK.Baru Or HdfTransStatus.Value = enuTCPICK.Prepared)

            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCPICK.Baru)
            BtnPreview.Enabled = (HdfTransStatus.Value = enuTCPICK.Prepared)

            If HdfTransStatus.Value = enuTCPICK.Baru Then
                BtnPrepare.Text = "Prepare"
            ElseIf HdfTransStatus.Value = enuTCPICK.Prepared Then
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
        psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        psFillGrvReserve(HdfPCLRefOID.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        RdlPickType.Enabled = False

        psEnableSave(True)
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        If HdfTransStatus.Value = enuTCPICK.Baru Then
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
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If
        If DstWhs.SelectedValue = "0" Then
            LblMsgWhs.Text = "Pilih Warehouse"
            vnSave = False
        End If
        If HdfPickTypeOID.Value = enuSchDType.TRB Or (HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip And HdfIsExpedition.Value = "0") Then
            If DstWhsDest.SelectedValue = "0" Then
                LblMsgWhsDest.Text = "Pilih Warehouse Tujuan"
                vnSave = False
            End If
        End If
        If HdfPCLRefOID.Value = "0" Then
            LblMsgPCLRefNo.Text = "Pilih Referensi"
            vnSave = False
        End If
        If Not IsDate(Trim(TxtPCLScheduleDate.Text)) Then
            LblMsgPCLScheduleDate.Text = "Isi Schedule"
            vnSave = False
        End If

        If HdfPickTypeOID.Value = enuSchDType.TRB Or HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
            ChkPCL_Rack.Checked = True
            ChkPCL_Floor.Checked = True
            ChkPCL_CrossDock.Checked = False
            ChkPCL_DOTitip.Checked = False
        ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
            ChkPCL_Rack.Checked = False
            ChkPCL_Floor.Checked = False
            ChkPCL_CrossDock.Checked = False
            ChkPCL_DOTitip.Checked = True
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

            HdfPCLStorageTypeList.Value = ""
            If ChkPCL_Rack.Checked Then
                HdfPCLStorageTypeList.Value += enuStorageType.Rack & ","
            End If
            If ChkPCL_Floor.Checked Then
                HdfPCLStorageTypeList.Value += enuStorageType.Floor & ","
            End If
            If ChkPCL_CrossDock.Checked Then
                HdfPCLStorageTypeList.Value += enuStorageType.CrossDock & ","
            End If
            If ChkPCL_DOTitip.Checked Then
                HdfPCLStorageTypeList.Value += enuStorageType.DO_Titip & ","
            End If
            HdfPCLStorageTypeList.Value = Mid(HdfPCLStorageTypeList.Value, 1, Len(HdfPCLStorageTypeList.Value) - 1)

            Dim vnQuery As String
            Dim vnUserNIP As String = Session("EmpNIP")

            Dim vnPLNo As String

            Dim vnWhsOID As String = HdfWhs.Value
            Dim vnWhsCode As String = fbuGetWhsCode_ByOID(vnWhsOID, vnSQLConn)

            Dim vnWhsOID_Dest As String

            If HdfPickTypeOID.Value = enuSchDType.TRB Or HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
                vnWhsOID_Dest = HdfWhsDest.Value
            Else
                vnWhsOID_Dest = "Null"
            End If

            Dim vnDestTypeOID As String
            If HdfPickTypeOID.Value = enuSchDType.Invoice Then
                vnDestTypeOID = IIf(ChkDest.Checked, enuDestType.Luar_Kota, enuDestType.Dalam_Kota)
            Else
                vnDestTypeOID = enuDestType.None
            End If

            Dim vnRefHOID As String = HdfPCLRefOID.Value
            Dim vnRefHNo As String = Trim(TxtPCLRefNo.Text)

            If HdfActionStatus.Value = cbuActionNew Then
                Dim vnDBMaster As String = fbuGetDBMaster()
                Dim vnCompanyCode As String = Trim(HdfCompanyCode.Value)

                If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
                    If fbuValNotaDetail_BrgCode(vnCompanyCode, vnRefHOID, vnSQLConn) = False Then
                        LblMsgError.Text = pbMsgError
                        LblMsgError.Visible = True

                        vnSQLConn.Close()
                        vnSQLConn.Dispose()
                        vnSQLConn = Nothing
                        Exit Sub
                    End If
                End If

                psSetTransNo(vnCompanyCode, vnWhsCode, vnSQLConn)
                vnPLNo = Trim(TxtPCLNo.Text)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoPCLHeader_TR with(nolock)"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoPCLHeader_TR(OID,PCLNo,PCLDate,PCLScheduleDate,"
                vnQuery += vbCrLf & "PCLCompanyCode,WarehouseOID,WarehouseOID_Dest,PCLStorageTypeList,"
                vnQuery += vbCrLf & "SchDTypeOID,PCLRefHOID,PCLRefHNo,PCLRefHInfo,"
                vnQuery += vbCrLf & "DestTypeOID,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & vnPLNo & "',getdate(),getdate(),"

                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "'," & vnWhsOID & "," & vnWhsOID_Dest & ",'" & HdfPCLStorageTypeList.Value & "',"
                vnQuery += vbCrLf & HdfPickTypeOID.Value & "," & vnRefHOID & ",'" & vnRefHNo & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtPCLDescr.Text)) & "',"

                vnQuery += vbCrLf & vnDestTypeOID & ","

                vnQuery += vbCrLf & "'" & stuTransCode.SsoPickList & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnOID, Trim(vnCompanyCode), vnRefHOID, vnSQLConn, vnSQLTrans)

                pbuInsertStatusPCL(vnOID, enuTCPICK.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCPICK.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

            Else
                vnPLNo = Trim(TxtPCLNo.Text)

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoPCLHeader_TR set"
                vnQuery += vbCrLf & "DestTypeOID=" & vnDestTypeOID & ","
                vnQuery += vbCrLf & "PCLScheduleDate='" & TxtPCLScheduleDate.Text & "',"
                vnQuery += vbCrLf & "PCLNote='" & fbuFormatString(Trim(TxtPCLDescr.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                pbuInsertStatusPCL(TxtTransID.Text, enuTCPICK.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

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
    Private Sub psSaveDetail(vriPCLHOID As String, vriCompanyCode As String, vriRefHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String

        If RdlPickType.SelectedValue = enuSchDType.TRB Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "KodeBrg,sum(Qty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmTRBDetail_TR with(nolock) Where TRBHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by KodeBrg"
            vnQuery += vbCrLf & "Having sum(Qty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        ElseIf RdlPickType.SelectedValue = enuSchDType.Invoice Or RdlPickType.SelectedValue = enuSchDType.DO_Titip Then
            '21 Sep 2023 Sys_DcmNotaDetail_ByBarang_TR
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "       KodeBarang,TotalQty + TotalQtyBonus - TotalQtyOnPickList,0"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR with(nolock) Where NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "       and (TotalQty + TotalQtyBonus - TotalQtyOnPickList) > 0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
            '21 Sep 2023 Sys_DcmNotaDetail_ByBarang_TR

            vnQuery = "insert into Sys_SsoPCLDetail_Paket_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "PAKETCODE,RefQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "       KodeBarang,(Qty + QtyBonus - QtyOnPickList)"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "       and KodeBarang in(Select b.PAKETCODE From " & vnDBMaster & "Sys_MstPaketH_MA b with(nolock) Where b.CompanyCode='" & vriCompanyCode & "')"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        ElseIf RdlPickType.SelectedValue = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "BRGCODE,Sum(PKDOTDQty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR with(nolock) Where PKDOTHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by BRGCODE"
            vnQuery += vbCrLf & "Having Sum(PKDOTDQty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End If
    End Sub
    Private Sub psSaveDetail_20230921_Bef_Point_Ke_NotaDetail_SummByBrg(vriPCLHOID As String, vriCompanyCode As String, vriRefHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String

        If RdlPickType.SelectedValue = enuSchDType.TRB Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "KodeBrg,sum(Qty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmTRBDetail_TR with(nolock) Where TRBHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by KodeBrg"
            vnQuery += vbCrLf & "Having sum(Qty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        ElseIf RdlPickType.SelectedValue = enuSchDType.Invoice Or RdlPickType.SelectedValue = enuSchDType.DO_Titip Then
            '<---19 Sep 2023 Original Bef Paketan
            'vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            'vnQuery += vbCrLf & "(PCLHOID,"
            'vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            'vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            'vnQuery += vbCrLf & "       KodeBarang,Sum(Qty + QtyBonus - QtyOnPickList),0"
            'vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriRefHOID
            'vnQuery += vbCrLf & "       and NOT KodeBarang in(Select b.PAKETCODE From " & vnDBMaster & "Sys_MstPaketH_MA b with(nolock) Where b.CompanyCode='" & vriCompanyCode & "')"
            'vnQuery += vbCrLf & "Group by KodeBarang"
            'vnQuery += vbCrLf & "Having Sum(Qty + QtyBonus - QtyOnPickList)>0"
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
            '<<==19 Sep 2023 Original Bef Paketan

            '<---19 Sep 2023 Masukin Barang Paket
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ",KodeBarang,sum(vTotalQty),0"
            vnQuery += vbCrLf & "From("
            vnQuery += vbCrLf & "Select KodeBarang,Sum(Qty + QtyBonus - QtyOnPickList) vTotalQty"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "       and NOT KodeBarang in(Select b.PAKETCODE From " & vnDBMaster & "Sys_MstPaketH_MA b with(nolock) Where b.CompanyCode='" & vriCompanyCode & "')"
            vnQuery += vbCrLf & "Group by KodeBarang"
            vnQuery += vbCrLf & "Having Sum(Qty + QtyBonus - QtyOnPickList)>0"

            vnQuery += vbCrLf & "UNION"

            vnQuery += vbCrLf & "Select pd.BRGCODE,Sum(nh.Qty + nh.QtyBonus - nh.QtyOnPickList)*pd.PaketQty vTotalQty"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR nh with(nolock)"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstPaketH_MA ph with(nolock) on ph.PAKETCODE=nh.KodeBarang"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstPaketD_MA pd with(nolock) on pd.PAKETHOID=ph.OID"
            vnQuery += vbCrLf & " Where nh.NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by pd.BRGCODE,pd.PaketQty"
            vnQuery += vbCrLf & "Having Sum(nh.Qty + nh.QtyBonus - nh.QtyOnPickList)>0)tb"

            vnQuery += vbCrLf & "Group by KodeBarang"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "insert into Sys_SsoPCLDetail_Paket_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "PAKETCODE,RefQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "       KodeBarang,(Qty + QtyBonus - QtyOnPickList)"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "       and KodeBarang in(Select b.PAKETCODE From " & vnDBMaster & "Sys_MstPaketH_MA b with(nolock) Where b.CompanyCode='" & vriCompanyCode & "')"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
            '<<==19 Sep 2023 Masukin Barang Paket

        ElseIf RdlPickType.SelectedValue = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "BRGCODE,Sum(PKDOTDQty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR with(nolock) Where PKDOTHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by BRGCODE"
            vnQuery += vbCrLf & "Having Sum(PKDOTDQty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End If
    End Sub

    Private Sub psSaveDetail_20230919_Orig_Bef_Paketan(vriPCLHOID As String, vriCompanyCode As String, vriRefHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String

        If RdlPickType.SelectedValue = enuSchDType.TRB Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "KodeBrg,sum(Qty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmTRBDetail_TR with(nolock) Where TRBHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by KodeBrg"
            vnQuery += vbCrLf & "Having sum(Qty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        ElseIf RdlPickType.SelectedValue = enuSchDType.Invoice Or RdlPickType.SelectedValue = enuSchDType.DO_Titip Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "       KodeBarang,Sum(Qty + QtyBonus - QtyOnPickList),0"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by KodeBarang"
            vnQuery += vbCrLf & "Having Sum(Qty + QtyBonus - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        ElseIf RdlPickType.SelectedValue = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Insert into Sys_SsoPCLDetail_TR"
            vnQuery += vbCrLf & "(PCLHOID,"
            vnQuery += vbCrLf & "BRGCODE,RefQty,PCLDQty)"
            vnQuery += vbCrLf & "Select " & vriPCLHOID & ","
            vnQuery += vbCrLf & "BRGCODE,Sum(PKDOTDQty - QtyOnPickList),0 From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR with(nolock) Where PKDOTHOID=" & vriRefHOID
            vnQuery += vbCrLf & "Group by BRGCODE"
            vnQuery += vbCrLf & "Having Sum(PKDOTDQty - QtyOnPickList)>0"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
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
        psStatusRefresh()
        If HdfProcess.Value = "CancelPCL" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelPCL()
        ElseIf HdfProcess.Value = "PreparePL" Then
            psPreparePCL()
        ElseIf HdfProcess.Value = "VoidPL" Then
            psVoidPCL()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub BtnPrepare_Click(sender As Object, e As EventArgs) Handles BtnPrepare.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Prepare) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Prepare Pick List No. " & TxtPCLNo.Text & " ?<br />WARNING : Prepare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "PreparePL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub psCancelPCL()
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
            Dim vnPCLHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoPCLHeader_TR with(nolock) Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus = enuTCPICK.Cancelled Or vnTransStatus > enuTCPICK.Prepared Then
                LblMsgError.Text = "Status Sudah Batal atau Picking"
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

            vnQuery = "Execute spSsoReservePicklist_Cancel " & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoPCLHeader_TR set TransStatus=" & enuTCPICK.Cancelled & ",PreparedFailed=0,PCLCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            '<---05 Sep 2023 baru dipindahin, asalnya after status diupdate jadi cancel
            If vnTransStatus = enuTCPICK.Prepared Then
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnTransStatus = enuTCPICK.Prepared = " & enuTCPICK.Prepared)
                vsTextStream.WriteLine("4")
                psUpdateQtyOnPicklist_Cancel(vsTextStream, vnPCLHOID, HdfPCLRefOID.Value, vnSQLConn, vnSQLTrans)
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("5")
            Else
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnTransStatus = enuTCPICK.Baru = " & enuTCPICK.Baru)
            End If
            '<<==05 Sep 2023 baru dipindahin, asalnya after status diupdate jadi cancel

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusPCL...Start")
            pbuInsertStatusPCL(vnPCLHOID, enuTCPICK.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCL...End")

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

    Private Sub psPreparePCL()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psPreparePCL", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnPCLHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoPCLHeader_TR with(nolock) Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            HdfTransStatus.Value = vnTransStatus

            If vnTransStatus > enuTCPICK.Baru Then
                LblMsgError.Text = "Status Sudah Prepared"
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

            vnQuery = "Update Sys_SsoPCLHeader_TR set TransStatus=" & enuTCPICK.Prepared & ",PreparedFailed=0,PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPCL...Start")
            pbuInsertStatusPCL(vnPCLHOID, enuTCPICK.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCL...End")

            vnQuery = "Execute spSsoReservePicklist " & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Select count(1) From Sys_SsoPCLReserve_TR with(nolock) Where PCLHOID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)

            If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                LblMsgError.Text = "TIDAK ADA STOCK...PREPARE GAGAL"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)

                '<---01 Oct 2023
                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoPCLHeader_TR set PreparedFailed=1 Where OID=" & vnPCLHOID
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("104 - Update Status Prepared Failed")
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("105")
                vsTextStream.WriteLine("pbuInsertStatusPCL...Start")
                pbuInsertStatusPCL(vnPCLHOID, enuTCPICK.Prepared_Failed, Session("UserOID"), vnSQLConn, vnSQLTrans)
                vsTextStream.WriteLine("pbuInsertStatusPCL...End")

                vnSQLTrans.Commit()
                vnSQLTrans = Nothing
                vnBeginTrans = False
                '<<==01 Oct 2023

                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("4")
            psUpdateQtyOnPicklist(vsTextStream, HdfPCLRefOID.Value, vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("5")

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
                vnSQLTrans.Dispose()
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

    Private Sub psUpdateQtyOnPicklist(vriTextStream As Scripting.TextStream, vriPCLRefHOID As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("<-----psUpdateQtyOnPicklist")
        vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("vriPCLRefHOID = " & vriPCLRefHOID)

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String
        Dim vnCheck As Integer
        If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnNotaDOID int"

            '21 Sep 2023 Detail_ByBarang
            'vnQuery += vbCrLf & "	Set @cr = cursor for Selecxt distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"

            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID in(" & enuSchDType.Invoice & "," & enuSchDType.DO_Titip & ") and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"

            '21 Sep 2023
            'vnQuery += vbCrLf & "			Select toxp 1 @vnNotaDOID = OID From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID and nd.KodeBarang = @vnBrgCode"
            'vnQuery += vbCrLf & "			Updatxe " & vnDBDcm & "Sys_DcmNotaDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnNotaDOID and NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR Set TotalQtyOnPickList=@vnQtyPCLD  Where NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"

            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            '21 Sep 2023
            'vnQuery = "Select sum(Qty + QtyBonus - QtyOnPickList) From " & vnDBDcm & "Sys_DcmNotaDetail_TR where NotaHOID=" & vriPCLRefHOID & " group by KodeBarang having sum(Qty + QtyBonus - QtyOnPickList)>0"
            vnQuery = "Select sum(TotalQty + TotalQtyBonus - TotalQtyOnPickList) From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR where NotaHOID=" & vriPCLRefHOID & " group by KodeBarang having sum(TotalQty + TotalQtyBonus - TotalQtyOnPickList)>0"

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmNotaHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                vnQuery = "Update " & vnDBDcm & "Sys_DcmJUAL set IsPickListClosed=1 where NO_NOTA='" & TxtPCLRefNo.Text & "' and CompanyCode='" & HdfCompanyCode.Value & "'"
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If

        ElseIf HdfPickTypeOID.Value = enuSchDType.TRB Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnTRBDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBrg From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.TRB & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnTRBDOID = OID From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID and nd.KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmTRBDetail_TR Set QtyOnPickList=@vnQtyPCLD Where OID = @vnTRBDOID and TRBHOID=@vnRefHOID and KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Select Sum(Qty - QtyOnPickList) From " & vnDBDcm & "Sys_DcmTRBDetail_TR where TRBHOID=" & vriPCLRefHOID & " group by KodeBrg having sum(Qty - QtyOnPickList)>0"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmTRBHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If

        ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "   Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnPKDTDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.BRGCODE From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.Perintah_Kirim_DO_Titip & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnPKDTDOID = OID From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID and nd.BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmPKDOTDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnPKDTDOID and PKDOTHOID=@vnRefHOID and BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Select Sum(RequestQty - QtyOnPickList) From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR where PKDOTHOID=" & vriPCLRefHOID & " group by BRGCODE having sum(RequestQty - QtyOnPickList)>0"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmPKDOTHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If
        End If

        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("<<====psUpdateQtyOnPicklist")
    End Sub

    Private Sub psUpdateQtyOnPicklist_Cancel(vriTextStream As Scripting.TextStream, vriPCLTransOID As Integer, vriPCLRefHOID As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("<-----psUpdateQtyOnPicklist_Cancel")
        vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("vriPCLRefHOID = " & vriPCLRefHOID)

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String

        If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnNotaDOID int"

            '21 Sep 2023
            'vnQuery += vbCrLf & "	Set @cr = cursor for Selecxt distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"

            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & HdfPickTypeOID.Value & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"

            '21 Sep 2023
            'vnQuery += vbCrLf & "			Select toxp 1 @vnNotaDOID = OID From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID and nd.KodeBarang = @vnBrgCode"
            'vnQuery += vbCrLf & "			Updatxe " & vnDBDcm & "Sys_DcmNotaDetail_TR Set QtyOnPickList=QtyOnPickList - @vnQtyPCLD  Where OID = @vnNotaDOID and NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"

            '22 Sep 2023
            'vnQuery += vbCrLf & "			Updatxe " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR Set TotalQtyOnPickList=TotalQtyOnPickList - @vnQtyPCLD  Where NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR Set TotalQtyOnPickList=@vnQtyPCLD Where NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"

            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmNotaHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmJUAL set IsPickListClosed=0 where NO_NOTA='" & TxtPCLRefNo.Text & "' and CompanyCode='" & HdfCompanyCode.Value & "'"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

        ElseIf HdfPickTypeOID.Value = enuSchDType.TRB Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnTRBDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBrg From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & HdfPickTypeOID.Value & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > 0 and dt.BRGCODE=@vnBrgCode"

            vnQuery += vbCrLf & "			Select top 1 @vnTRBDOID = OID From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID and nd.KodeBrg = @vnBrgCode"

            '22 Sep 2023
            'vnQuery += vbCrLf & "			Updatxe " & vnDBDcm & "Sys_DcmTRBDetail_TR Set QtyOnPickList=QtyOnPickList - @vnQtyPCLD  Where OID = @vnTRBDOID and TRBHOID=@vnRefHOID and KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmTRBDetail_TR Set QtyOnPickList=@vnQtyPCLD Where OID = @vnTRBDOID and TRBHOID=@vnRefHOID and KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmTRBHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

        ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "   Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnPKDTDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.BRGCODE From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & HdfPickTypeOID.Value & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > 0 and dt.BRGCODE=@vnBrgCode"

            vnQuery += vbCrLf & "			Select top 1 @vnPKDTDOID = OID From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID and nd.BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmPKDOTDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnPKDTDOID and PKDOTHOID=@vnRefHOID and BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmPKDOTHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
        End If

        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("<<====psUpdateQtyOnPicklist_Cancel")
    End Sub

    Private Sub psUpdateQtyOnPicklist_20230921_Orig_Bef_NotaDetail_ByBarang(vriTextStream As Scripting.TextStream, vriPCLRefHOID As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("<-----psUpdateQtyOnPicklist")
        vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("vriPCLRefHOID = " & vriPCLRefHOID)

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String
        Dim vnCheck As Integer
        If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnNotaDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID in(" & enuSchDType.Invoice & "," & enuSchDType.DO_Titip & ") and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnNotaDOID = OID From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID and nd.KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmNotaDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnNotaDOID and NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Select sum(Qty + QtyBonus - QtyOnPickList) From " & vnDBDcm & "Sys_DcmNotaDetail_TR where NotaHOID=" & vriPCLRefHOID & " group by KodeBarang having sum(Qty + QtyBonus - QtyOnPickList)>0"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmNotaHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                vnQuery = "Update " & vnDBDcm & "Sys_DcmJUAL set IsPickListClosed=1 where NO_NOTA='" & TxtPCLRefNo.Text & "' and CompanyCode='" & HdfCompanyCode.Value & "'"
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If

        ElseIf HdfPickTypeOID.Value = enuSchDType.TRB Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnTRBDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBrg From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.TRB & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnTRBDOID = OID From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID and nd.KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmTRBDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnTRBDOID and TRBHOID=@vnRefHOID and KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Select Sum(Qty - QtyOnPickList) From " & vnDBDcm & "Sys_DcmTRBDetail_TR where TRBHOID=" & vriPCLRefHOID & " group by KodeBrg having sum(Qty - QtyOnPickList)>0"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmTRBHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If

        ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "   Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnPKDTDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.BRGCODE From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.Perintah_Kirim_DO_Titip & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > " & enuTCPICK.Baru & " and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnPKDTDOID = OID From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID and nd.BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmPKDOTDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnPKDTDOID and PKDOTHOID=@vnRefHOID and BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Select Sum(RequestQty - QtyOnPickList) From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR where PKDOTHOID=" & vriPCLRefHOID & " group by BRGCODE having sum(RequestQty - QtyOnPickList)>0"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)

            vnCheck = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
            If vnCheck = 0 Then
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = 0")

                vnQuery = "Update " & vnDBDcm & "Sys_DcmPKDOTHeader_TR set IsPickListClosed=1 where OID=" & vriPCLRefHOID
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnQuery")
                vriTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("vnCheck = " & vnCheck)
            End If
        End If

        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("<<====psUpdateQtyOnPicklist")
    End Sub

    Private Sub psUpdateQtyOnPicklist_Cancel_20230921_Orig_Bef_NotaDetail_ByBarang(vriTextStream As Scripting.TextStream, vriPCLRefHOID As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("<-----psUpdateQtyOnPicklist_Cancel")
        vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("vriPCLRefHOID = " & vriPCLRefHOID)

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnQuery As String

        If HdfPickTypeOID.Value = enuSchDType.Invoice Or HdfPickTypeOID.Value = enuSchDType.DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnNotaDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBarang From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID in(" & enuSchDType.Invoice & "," & enuSchDType.DO_Titip & ") and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > 0 and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnNotaDOID = OID From " & vnDBDcm & "Sys_DcmNotaDetail_TR nd with(nolock) Where nd.NotaHOID=@vnRefHOID and nd.KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmNotaDetail_TR Set QtyOnPickList=QtyOnPickList - @vnQtyPCLD  Where OID = @vnNotaDOID and NotaHOID=@vnRefHOID and KodeBarang = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmNotaHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmJUAL set IsPickListClosed=0 where NO_NOTA='" & TxtPCLRefNo.Text & "' and CompanyCode='" & HdfCompanyCode.Value & "'"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

        ElseIf HdfPickTypeOID.Value = enuSchDType.TRB Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "	Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnTRBDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.KodeBrg From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.TRB & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > 0 and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnTRBDOID = OID From " & vnDBDcm & "Sys_DcmTRBDetail_TR nd with(nolock) Where nd.TRBHOID=@vnRefHOID and nd.KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmTRBDetail_TR Set QtyOnPickList=QtyOnPickList - @vnQtyPCLD  Where OID = @vnTRBDOID and TRBHOID=@vnRefHOID and KodeBrg = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmTRBHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

        ElseIf HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip Then
            vnQuery = "Begin"
            vnQuery += vbCrLf & "	Declare @vnRefHOID int"
            vnQuery += vbCrLf & "	Set @vnRefHOID = " & vriPCLRefHOID
            vnQuery += vbCrLf & "   Declare @cr cursor,@vnBrgCode varchar(45),@vnQtyPCLD int,@vnPKDTDOID int"
            vnQuery += vbCrLf & "	Set @cr = cursor for Select distinct nd.BRGCODE From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID"
            vnQuery += vbCrLf & "	Open @cr"
            vnQuery += vbCrLf & "	Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "	While @@FETCH_STATUS = 0"
            vnQuery += vbCrLf & "		Begin"
            vnQuery += vbCrLf & "			Select @vnQtyPCLD  = isnull(sum(dt.PCLDQty),0)"
            vnQuery += vbCrLf & "				From Sys_SsoPCLDetail_TR dt with(nolock)"
            vnQuery += vbCrLf & "					 inner join Sys_SsoPCLHeader_TR dh with(nolock) on dh.OID=dt.PCLHOID"
            vnQuery += vbCrLf & "				Where dh.SchDTypeOID=" & enuSchDType.Perintah_Kirim_DO_Titip & " and dh.PCLRefHOID=@vnRefHOID and dh.TransStatus > 0 and dt.BRGCODE=@vnBrgCode"
            vnQuery += vbCrLf & "			Select top 1 @vnPKDTDOID = OID From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR nd with(nolock) Where nd.PKDOTHOID=@vnRefHOID and nd.BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Update " & vnDBDcm & "Sys_DcmPKDOTDetail_TR Set QtyOnPickList=@vnQtyPCLD  Where OID = @vnPKDTDOID and PKDOTHOID=@vnRefHOID and BRGCODE = @vnBrgCode"
            vnQuery += vbCrLf & "			Fetch @cr into @vnBrgCode"
            vnQuery += vbCrLf & "		End"
            vnQuery += vbCrLf & "	Close @cr"
            vnQuery += vbCrLf & "	Deallocate @cr"
            vnQuery += vbCrLf & "End"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

            vnQuery = "Update " & vnDBDcm & "Sys_DcmPKDOTHeader_TR set IsPickListClosed=0 where OID=" & vriPCLRefHOID
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
        End If

        vriTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("<<====psUpdateQtyOnPicklist_Cancel")
    End Sub
    Protected Sub BtnCancelPCL_Click(sender As Object, e As EventArgs) Handles BtnCancelPCL.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Pick List No. " & TxtPCLNo.Text & " ?<br />WARNING : Batal Pick List Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelPCL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        If fsPrintKe() = False Then
            Exit Sub
        End If
        Dim vnCrpFileName As String = ""
        psGenerateCrp(vnCrpFileName)

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

    Private Sub psGenerateCrp(ByRef vriCrpFileName As String)
        Dim vnTransOID As String = TxtTransID.Text

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoPickList
        vbuCrpSubReport1 = stuSsoCrp.CrpSsoPickListReserve_SR

        vbuCrpQuery = "Select ta.*,mc.CompanyCode,mc.CompanyName,sd.SchDTypeName,"
        vbuCrpQuery += vbCrLf & "      wha.WarehouseName vWarehouseNameAsal,"
        vbuCrpQuery += vbCrLf & "      whd.WarehouseName vWarehouseNameTujuan,"
        vbuCrpQuery += vbCrLf & "      row_number()over(order by mb.BRGCODE)vDSeqNo,mb.BRGNAME,mb.BRGUNIT," & HdfPrioritas.Value & " vPrioritas"
        vbuCrpQuery += vbCrLf & " From fnTbl_SsoPCLHeader_Detail(" & vnTransOID & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery += vbCrLf & "      inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.PCLCompanyCode"
        vbuCrpQuery += vbCrLf & "      inner join " & vnDBDcm & "Sys_DcmSchDType_MA sd with(nolock) on sd.OID=ta.SchDTypeOID"
        vbuCrpQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.PCLCompanyCode and mb.BRGCODE=ta.BRGCODE"

        vbuCrpQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA wha with(nolock) on wha.OID=ta.WarehouseOID"
        vbuCrpQuery += vbCrLf & "      left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=ta.WarehouseOID_Dest"
        vbuCrpQuery += vbCrLf & " order by mb.BRGCODE"

        vbuCrpQuery1 = "Select ta.*,mc.CompanyCode,mc.CompanyName,sd.SchDTypeName,"
        vbuCrpQuery1 += vbCrLf & "      wha.WarehouseName vWarehouseNameAsal,"
        vbuCrpQuery1 += vbCrLf & "      whd.WarehouseName vWarehouseNameTujuan,"
        vbuCrpQuery1 += vbCrLf & "      row_number()over(order by mb.BRGCODE)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,sti.vStorageInfo"
        vbuCrpQuery1 += vbCrLf & " From fnTbl_SsoPCLHeader_Reserve(" & vnTransOID & ",'" & Session("UserID") & "')ta"
        vbuCrpQuery1 += vbCrLf & "      inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.PCLCompanyCode"
        vbuCrpQuery1 += vbCrLf & "      inner join " & vnDBDcm & "Sys_DcmSchDType_MA sd with(nolock) on sd.OID=ta.SchDTypeOID"
        vbuCrpQuery1 += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.PCLCompanyCode and mb.BRGCODE=ta.BRGCODE"

        vbuCrpQuery1 += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA wha with(nolock) on wha.OID=ta.WarehouseOID"
        vbuCrpQuery1 += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') sti on sti.vStorageOID=ta.StorageOID"
        vbuCrpQuery1 += vbCrLf & "      left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=ta.WarehouseOID_Dest"
        vbuCrpQuery1 += vbCrLf & " order by mb.BRGCODE"
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Private Sub psShowListDoc(vriBo As Boolean)
        If vriBo Then
            DivListDoc.Style(HtmlTextWriterStyle.Visibility) = "visible"
            HdfOnList.Value = enuSchDType.Invoice

            psShowListTRB(False)
            psShowListPKDOT(False)

            TxtListDocNota.Focus()
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
            psShowListPKDOT(False)

            TxtListTRBNo.Focus()
        Else
            DivListTRB.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub psShowListPKDOT(vriBo As Boolean)
        If vriBo Then
            DivListPKDOT.Style(HtmlTextWriterStyle.Visibility) = "visible"
            HdfOnList.Value = enuSchDType.Perintah_Kirim_DO_Titip

            psShowListDoc(False)
            psShowListTRB(False)

            TxtListPKDOTNo.Focus()
        Else
            DivListPKDOT.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnListDocClose_Click(sender As Object, e As EventArgs) Handles BtnListDocClose.Click
        psShowListDoc(False)
    End Sub
    Protected Sub BtnListDocFind_Click(sender As Object, e As EventArgs) Handles BtnListDocFind.Click
        LblMsgListDoc.Text = ""
        If Trim(TxtListDocCustomer.Text) = "" And Trim(TxtListDocNota.Text) = "" Then
            LblMsgListDoc.Text = "Pilih Nomor Nota atau Customer"
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
    Private Sub psFillGrvListDoc(vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select Distinct CompanyCode,no_nota,convert(varchar(11),tanggal,106)vtanggal,kode_cust,CUSTOMER,ALAMAT,kota,WarehouseOID,tanggal"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmJUAL"
        vnQuery += vbCrLf & "Where IsPickListClosed=0 and CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and WarehouseOID=" & DstWhs.SelectedValue
        If Trim(TxtListDocCustomer.Text) <> "" Then
            vnQuery += vbCrLf & "      and CUSTOMER like '%" & fbuFormatString(Trim(TxtListDocCustomer.Text)) & "%'"
        End If
        If Trim(TxtListDocNota.Text) <> "" Then
            vnQuery += vbCrLf & "      and no_nota like '%" & fbuFormatString(Trim(TxtListDocNota.Text)) & "%'"
        End If
        vnQuery += vbCrLf & "Order by tanggal,no_nota"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListDoc.DataSource = vnDtb
        GrvListDoc.DataBind()

        TxtListDocNota.Focus()
    End Sub

    Private Sub psGetReferensi_Invoice()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select Distinct CompanyCode,no_nota,convert(varchar(11),tanggal,106)vtanggal,kode_cust,CUSTOMER,ALAMAT,kota,WarehouseOID,tanggal,"
        vnQuery += vbCrLf & "      case when abs(NotaPrio)=1 then 4 else 5 end,"
        vnQuery += vbCrLf & "      case when abs(NotaPrio)=1 then NotaPRIODatetime else getdate() end"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmJUAL with(nolock)"
        vnQuery += vbCrLf & "Where NotaCancel=0 and NotaDOT=0 and IsPickListClosed=0 and CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and GDG in('" & stuWarehouse.Kepu_Baru & "','" & stuWarehouse.Prancis_Baru & "')"
        vnQuery += vbCrLf & "      and WarehouseOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And not rtrim(CompanyCode)+'x'+no_nota in(Select b.PCLCompanyCode+'x'+b.PCLRefHNo From Sys_SsoPCLHeader_TR b with(nolock) where b.TransStatus=" & enuTCPICK.Baru & " and b.SchDTypeOID=" & enuSchDType.Invoice & ")"
        vnQuery += vbCrLf & "Order by case when abs(NotaPrio)=1 then 4 else 5 end,"
        vnQuery += vbCrLf & "      case when abs(NotaPrio)=1 then NotaPRIODatetime else getdate() end,"
        vnQuery += vbCrLf & "      tanggal,no_nota"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""

            TxtPCLRefNo.Text = ""
            HdfPCLRefOID.Value = "0"
            TxtPCLRefOID.Text = ""

            HdfWhs.Value = "0"

            TxtPCLDescr.Text = ""

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
            psFillGrvInv(HdfPCLRefOID.Value, vnSQLConn)
        Else
            Dim vnDRow As DataRow
            vnDRow = vnDtb.Rows(0)

            Dim vnCompanyCode As String = vnDRow.Item("CompanyCode")
            Dim vnNotaNo As String = vnDRow.Item("no_nota")

            Dim vnSQLTrans As SqlTransaction = Nothing
            Dim vnBeginTrans As Boolean
            Try
                Dim vnNotaHOID As String = "0"
                HdfPCLRefOID.Value = vnNotaHOID
                TxtPCLRefOID.Text = vnNotaHOID

                pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psGetReferensi_Invoice", 0, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

                vnSQLTrans = vnSQLConn.BeginTransaction("inp")
                vnBeginTrans = True

                pbuInsertNotaByPrepare(vsTextStream, vnCompanyCode, vnNotaNo, vnNotaHOID, vnSQLConn, vnSQLTrans)

                vnSQLTrans.Commit()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                DstCompany.SelectedValue = Trim(vnCompanyCode)
                HdfCompanyCode.Value = Trim(vnCompanyCode)

                TxtPCLRefNo.Text = vnNotaNo
                HdfPCLRefOID.Value = vnNotaHOID
                TxtPCLRefOID.Text = vnNotaHOID

                HdfPickTypeOID.Value = enuSchDType.Invoice

                HdfWhs.Value = vnDRow.Item("WarehouseOID")

                TxtPCLDescr.Text = "Customer=" & vnDRow.Item("kode_cust") & " " & vnDRow.Item("CUSTOMER") & Chr(10) & vnDRow.Item("ALAMAT") & " " & vnDRow.Item("kota")

                psFillGrvInv(HdfPCLRefOID.Value, vnSQLConn)

                If fbuValNotaDetail_BrgCode(vnCompanyCode, vnNotaHOID, vnSQLConn) = False Then
                    LblMsgError.Text = pbMsgError
                    LblMsgError.Visible = True

                    psFillGrvDetail(0, vnSQLConn)
                Else
                    'pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "pbuInsertNotaDetail_ByBarang", HdfPCLRefOID.Value, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

                    vnSQLTrans = vnSQLConn.BeginTransaction()
                    vnBeginTrans = True

                    pbuInsertNotaDetail_ByBarang(vsTextStream, vnCompanyCode, HdfPCLRefOID.Value, vnSQLConn, vnSQLTrans)

                    vnSQLTrans.Commit()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing

                    psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
                End If

                psShowListDoc(False)
            Catch ex As Exception
                LblMsgListDoc.Text = ex.Message
                LblMsgListDoc.Visible = True

                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing
                End If

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("ERROR")
                vsTextStream.WriteLine(ex.Message)
            Finally
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Date.Now)
                vsTextStream.WriteLine("--------------------EOF--------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
            End Try
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psGetReferensi_DO_Titip()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select Distinct CompanyCode,no_nota,convert(varchar(11),tanggal,106)vtanggal,kode_cust,CUSTOMER,ALAMAT,kota,WarehouseOID,tanggal"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmJUAL with(nolock)"
        vnQuery += vbCrLf & "Where NotaDOT=1 and IsPickListClosed=0 and CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and WarehouseOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And not rtrim(CompanyCode)+'x'+no_nota in(Select b.PCLCompanyCode+'x'+b.PCLRefHNo From Sys_SsoPCLHeader_TR b with(nolock) where b.TransStatus=" & enuTCPICK.Baru & ")"
        vnQuery += vbCrLf & "Order by tanggal,no_nota"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""

            TxtPCLRefNo.Text = ""
            HdfPCLRefOID.Value = "0"
            TxtPCLRefOID.Text = ""

            HdfWhs.Value = "0"

            TxtPCLDescr.Text = ""

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        Else
            Dim vnDRow As DataRow
            vnDRow = vnDtb.Rows(0)

            Dim vnCompanyCode As String = vnDRow.Item("CompanyCode")
            Dim vnNotaNo As String = vnDRow.Item("no_nota")

            Dim vnSQLTrans As SqlTransaction = Nothing
            Dim vnBeginTrans As Boolean
            Try
                Dim vnNotaHOID As String = "0"
                pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psGetReferensi_DO_Titip", 0, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

                vnSQLTrans = vnSQLConn.BeginTransaction("inp")
                vnBeginTrans = True
                pbuInsertNotaByPrepare(vsTextStream, vnCompanyCode, vnNotaNo, vnNotaHOID, vnSQLConn, vnSQLTrans)

                vnSQLTrans.Commit()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                DstCompany.SelectedValue = Trim(vnCompanyCode)
                HdfCompanyCode.Value = Trim(vnCompanyCode)

                TxtPCLRefNo.Text = vnNotaNo
                HdfPCLRefOID.Value = vnNotaHOID
                TxtPCLRefOID.Text = vnNotaHOID

                HdfPickTypeOID.Value = enuSchDType.DO_Titip

                HdfWhs.Value = vnDRow.Item("WarehouseOID")

                TxtPCLDescr.Text = "Customer=" & vnDRow.Item("kode_cust") & " " & vnDRow.Item("CUSTOMER") & Chr(10) & vnDRow.Item("ALAMAT") & " " & vnDRow.Item("kota")

                '<---26 Sep 2023
                'psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
                psFillGrvInv(HdfPCLRefOID.Value, vnSQLConn)

                If fbuValNotaDetail_BrgCode(vnCompanyCode, vnNotaHOID, vnSQLConn) = False Then
                    LblMsgError.Text = pbMsgError
                    LblMsgError.Visible = True

                    psFillGrvDetail(0, vnSQLConn)
                Else
                    pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "pbuInsertNotaDetail_ByBarang", HdfPCLRefOID.Value, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

                    vnSQLTrans = vnSQLConn.BeginTransaction()
                    vnBeginTrans = True

                    pbuInsertNotaDetail_ByBarang(vsTextStream, vnCompanyCode, HdfPCLRefOID.Value, vnSQLConn, vnSQLTrans)

                    vnSQLTrans.Commit()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing

                    psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
                End If
                '<<==26 Sep 2023
                psShowListDoc(False)
            Catch ex As Exception
                LblMsgListDoc.Text = ex.Message
                LblMsgListDoc.Visible = True

                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing
                End If

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("ERROR")
                vsTextStream.WriteLine(ex.Message)
            Finally
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Date.Now)
                vsTextStream.WriteLine("--------------------EOF--------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
            End Try
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
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
    Private Sub GrvListDoc_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListDoc.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "no_nota" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowList As GridViewRow = GrvListDoc.Rows(vnIdx)

            Dim vnCompanyCode As String = Trim(vnGRowList.Cells(ensColListDoc.CompanyCode).Text)
            Dim vnNotaNo As String = DirectCast(vnGRowList.Cells(ensColListDoc.no_nota).Controls(0), LinkButton).Text

            LblMsgListDoc.Text = ""
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgListDoc.Text = pbMsgError
                LblMsgListDoc.Visible = True
                Exit Sub
            End If

            Dim vnSQLTrans As SqlTransaction = Nothing
            Dim vnBeginTrans As Boolean
            Try
                Dim vnNotaHOID As String = "0"
                pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "GrvListDoc_RowCommand", 0, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)

                vnSQLTrans = vnSQLConn.BeginTransaction("inp")
                vnBeginTrans = True
                pbuInsertNotaByPrepare(vsTextStream, vnCompanyCode, vnNotaNo, vnNotaHOID, vnSQLConn, vnSQLTrans)

                vnSQLTrans.Commit()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                DstCompany.SelectedValue = Trim(vnCompanyCode)
                HdfCompanyCode.Value = Trim(vnCompanyCode)

                TxtPCLRefNo.Text = vnNotaNo
                HdfPCLRefOID.Value = vnNotaHOID
                TxtPCLRefOID.Text = vnNotaHOID

                HdfPickTypeOID.Value = enuSchDType.Invoice

                HdfWhs.Value = vnGRowList.Cells(ensColListDoc.WarehouseOID).Text

                TxtPCLDescr.Text = "Customer=" & vnGRowList.Cells(ensColListDoc.kode_cust).Text & " " & vnGRowList.Cells(ensColListDoc.CUSTOMER).Text & Chr(10) & vnGRowList.Cells(ensColListDoc.ALAMAT).Text & " " & vnGRowList.Cells(ensColListDoc.kota).Text

                psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)

                psShowListDoc(False)
            Catch ex As Exception
                LblMsgListDoc.Text = ex.Message
                LblMsgListDoc.Visible = True

                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing
                End If

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("ERROR")
                vsTextStream.WriteLine(ex.Message)
            Finally
                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing

                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Date.Now)
                vsTextStream.WriteLine("--------------------EOF--------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
            End Try
        End If
    End Sub
    Private Sub BtnListTRBClose_Click(sender As Object, e As EventArgs) Handles BtnListTRBClose.Click
        psShowListTRB(False)
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

    Private Sub psFillGrvListTRB(vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserOID As String = Session("UserOID")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.OID,PM.CompanyCode,PM.NoBukti,convert(varchar(11),PM.Tanggal,106)vTanggal,PM.GudangAsal,PM.GudangTujuan,PM.WarehouseAsalOID,PM.WarehouseTujuanOID,PM.Tanggal"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=PM.CompanyCode and mu.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "      inner join Sys_SsoUserWarehouse_MA mw with(nolock) on mw.WarehouseOID=PM.WarehouseAsalOID and mw.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "Where PM.IsPickListClosed=0 and PM.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      And PM.WarehouseAsalOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And PM.WarehouseTujuanOID=" & DstWhsDest.SelectedValue
        vnQuery += vbCrLf & "      And not PM.OID in(Select b.PCLRefHOID From Sys_SsoPCLHeader_TR b where b.TransStatus=" & enuTCPICK.Baru & ")"
        If Trim(TxtListTRBNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and PM.NoBukti like '%" & fbuFormatString(Trim(TxtListTRBNo.Text)) & "%'"
        End If

        vnQuery += vbCrLf & "Order by PM.Tanggal,PM.NoBukti"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListTRB.DataSource = vnDtb
        GrvListTRB.DataBind()

        TxtListTRBNo.Focus()
    End Sub

    Private Sub psGetReferensi_TRB()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserOID As String = Session("UserOID")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select Top 1 PM.OID,PM.CompanyCode,PM.NoBukti,convert(varchar(11),PM.Tanggal,106)vTanggal,PM.GudangAsal,PM.GudangTujuan,PM.WarehouseAsalOID,PM.WarehouseTujuanOID,PM.Tanggal"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=PM.CompanyCode and mu.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "      inner join Sys_SsoUserWarehouse_MA mw with(nolock) on mw.WarehouseOID=PM.WarehouseAsalOID and mw.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "Where PM.TRBCancel=0 and PM.IsPickListClosed=0 and PM.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      And PM.WarehouseAsalOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And PM.WarehouseTujuanOID=" & DstWhsDest.SelectedValue
        vnQuery += vbCrLf & "      And not PM.OID in(Select b.PCLRefHOID From Sys_SsoPCLHeader_TR b where b.TransStatus=" & enuTCPICK.Baru & " and b.SchDTypeOID=" & enuSchDType.TRB & ")"

        vnQuery += vbCrLf & "Order by PM.Tanggal,PM.NoBukti"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""

            TxtPCLRefNo.Text = ""
            HdfPCLRefOID.Value = "0"
            TxtPCLRefOID.Text = ""

            HdfWhs.Value = "0"
            HdfWhsDest.Value = "0"

            TxtPCLDescr.Text = ""
            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        Else
            Dim vnDRow As DataRow
            vnDRow = vnDtb.Rows(0)
            Dim vnCompanyCode As String = vnDRow.Item("CompanyCode")
            Dim vnNoBukti As String = vnDRow.Item("NoBukti")

            DstCompany.SelectedValue = Trim(vnCompanyCode)
            HdfCompanyCode.Value = vnCompanyCode

            TxtPCLRefNo.Text = vnNoBukti
            HdfPCLRefOID.Value = vnDRow.Item("OID")
            TxtPCLRefOID.Text = vnDRow.Item("OID")

            HdfPickTypeOID.Value = enuSchDType.TRB

            HdfWhs.Value = vnDRow.Item("WarehouseAsalOID")
            HdfWhsDest.Value = vnDRow.Item("WarehouseTujuanOID")

            TxtPCLDescr.Text = "Gudang Asal=" & vnDRow.Item("GudangAsal") & Chr(10) & "Gudang Tujuan=" & vnDRow.Item("GudangTujuan")
            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        End If

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

    Private Sub GrvListTRB_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListTRB.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "NoBukti" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowList As GridViewRow = GrvListTRB.Rows(vnIdx)

            Dim vnCompanyCode As String = Trim(vnGRowList.Cells(ensColListTRB.CompanyCode).Text)
            Dim vnNoBukti As String = DirectCast(vnGRowList.Cells(ensColListTRB.NoBukti).Controls(0), LinkButton).Text

            DstCompany.SelectedValue = vnCompanyCode
            HdfCompanyCode.Value = vnCompanyCode

            TxtPCLRefNo.Text = vnNoBukti
            HdfPCLRefOID.Value = vnGRowList.Cells(ensColListTRB.OID).Text
            TxtPCLRefOID.Text = vnGRowList.Cells(ensColListTRB.OID).Text

            HdfPickTypeOID.Value = enuSchDType.TRB

            HdfWhs.Value = vnGRowList.Cells(ensColListTRB.WarehouseAsalOID).Text
            HdfWhsDest.Value = vnGRowList.Cells(ensColListTRB.WarehouseTujuanOID).Text

            TxtPCLDescr.Text = "Gudang Asal=" & vnGRowList.Cells(ensColListTRB.GudangAsal).Text & Chr(10) & "Gudang Tujuan=" & vnGRowList.Cells(ensColListTRB.GudangTujuan).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowListTRB(False)
        End If
    End Sub
    Private Sub psFillGrvListPKDOT(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserOID As String = Session("UserOID")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.OID,PM.PKDOTCompanyCode,PM.PKDOTNo,convert(varchar(11),PM.PKDOTScheduleDate,106)vPKDOTScheduleDate,"
        vnQuery += vbCrLf & "     WM.WarehouseName,whd.WarehouseName vWarehouseName_Dest,"
        vnQuery += vbCrLf & "     PM.NotaHOID,nh.NotaNo,nh.CustCode+' '+nh.CustName vCustomer,"
        vnQuery += vbCrLf & "     PM.PKDOTScheduleDate,PM.WarehouseOID,PM.WarehouseOID_Dest"

        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=PM.PKDOTCompanyCode and mu.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA mw with(nolock) on mw.WarehouseOID=PM.WarehouseOID and mw.UserOID=" & vnUserOID

        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock) on nh.OID=PM.NotaHOID and nh.NotaNo like '%" & Trim(TxtListRefNo.Text) & "%'"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=PM.WarehouseOID_Dest"
        vnQuery += vbCrLf & "Where PM.TransStatus in(" & enuTCPerintahKirimDOT.Prepared & "," & enuTCPerintahKirimDOT.Dalam_Picklist & ") and PM.IsPickListClosed=0 and PM.PKDOTCompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      And PM.WarehouseOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And PM.WarehouseOID_Dest=" & DstWhsDest.SelectedValue

        If Trim(TxtListPKDOTNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and PM.PKDOTNo like '%" & fbuFormatString(Trim(TxtListPKDOTNo.Text)) & "%'"
        End If

        vnQuery += vbCrLf & "Order by PM.PKDOTDate,PM.PKDOTNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListPKDOT.DataSource = vnDtb
        GrvListPKDOT.DataBind()

        TxtListPKDOTNo.Focus()
    End Sub

    Private Sub psGetReferensi_Perintah_Kirim_DO_Titip()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserOID As String = Session("UserOID")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select Top 1 PM.OID,PM.PKDOTCompanyCode,PM.PKDOTNo,convert(varchar(11),PM.PKDOTScheduleDate,106)vPKDOTScheduleDate,"
        vnQuery += vbCrLf & "     WM.WarehouseName,whd.WarehouseName vWarehouseName_Dest,"
        vnQuery += vbCrLf & "     PM.CustCode+' '+mc.CustName +"
        vnQuery += vbCrLf & "     char(10)+'Ship To: '+PKDOTShipToName +"
        vnQuery += vbCrLf & "     char(10)+'Ship Address: '+PKDOTShipToAddress vCustomer,"
        vnQuery += vbCrLf & "     PM.PKDOTScheduleDate,PM.WarehouseOID,PM.WarehouseOID_Dest,abs(IsExpedition)vIsExpedition"

        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=PM.PKDOTCompanyCode and mu.UserOID=" & vnUserOID
        vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA mw with(nolock) on mw.WarehouseOID=PM.WarehouseOID and mw.UserOID=" & vnUserOID

        vnQuery += vbCrLf & "     inner join (Select distinct CUSTSUB,CUSTNAME From " & vnDBMaster & "Sys_MstCustomer_MA mcs with(nolock)) mc on mc.CUSTSUB=PM.CustCode"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=PM.WarehouseOID_Dest"
        vnQuery += vbCrLf & "Where PM.TransStatus in(" & enuTCPerintahKirimDOT.Prepared & "," & enuTCPerintahKirimDOT.Dalam_Picklist & ") and PM.IsPickListClosed=0 and PM.PKDOTCompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      And PM.WarehouseOID=" & DstWhs.SelectedValue
        vnQuery += vbCrLf & "      And PM.WarehouseOID_Dest=" & DstWhsDest.SelectedValue
        vnQuery += vbCrLf & "      And not PM.OID in(Select b.PCLRefHOID From Sys_SsoPCLHeader_TR b with(nolock) where b.TransStatus=" & enuTCPICK.Baru & " and b.SchDTypeOID=" & enuSchDType.Perintah_Kirim_DO_Titip & ")"

        vnQuery += vbCrLf & "Order by PM.PKDOTDate,PM.PKDOTNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""

            TxtPCLRefNo.Text = ""
            HdfPCLRefOID.Value = "0"
            TxtPCLRefOID.Text = ""

            HdfWhs.Value = "0"
            HdfWhsDest.Value = "0"
            HdfIsExpedition.Value = "0"

            TxtPCLDescr.Text = ""

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        Else
            Dim vnDRow As DataRow
            vnDRow = vnDtb.Rows(0)

            Dim vnCompanyCode As String = vnDRow.Item("PKDOTCompanyCode")
            Dim vnPKDOTNo As String = vnDRow.Item("PKDOTNo")

            DstCompany.SelectedValue = Trim(vnCompanyCode)
            HdfCompanyCode.Value = vnCompanyCode

            TxtPCLRefNo.Text = vnPKDOTNo
            HdfPCLRefOID.Value = vnDRow.Item("OID")
            TxtPCLRefOID.Text = vnDRow.Item("OID")

            HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip

            HdfWhs.Value = vnDRow.Item("WarehouseOID")
            HdfWhsDest.Value = vnDRow.Item("WarehouseOID_Dest")
            HdfIsExpedition.Value = vnDRow.Item("vIsExpedition")

            If HdfIsExpedition.Value = "0" Then
                TxtPCLDescr.Text = vnPKDOTNo & Chr(10) & vnDRow.Item("vCustomer") & Chr(10) & vnDRow.Item("WarehouseName") & Chr(10) & " ke " & vnDRow.Item("vWarehouseName_Dest")
            Else
                TxtPCLDescr.Text = vnPKDOTNo & Chr(10) & vnDRow.Item("vCustomer") & Chr(10) & vnDRow.Item("WarehouseName")
            End If

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnPCLRefNo_Click(sender As Object, e As EventArgs) Handles BtnPCLRefNo.Click
        If HdfActionStatus.Value = cbuActionNorm Then Exit Sub
        If RdlPickType.SelectedValue = enuSchDType.Invoice Then
            psShowListDoc(True)
        ElseIf RdlPickType.SelectedValue = enuSchDType.TRB Then
            psShowListTRB(True)
        ElseIf RdlPickType.SelectedValue = enuSchDType.Perintah_Kirim_DO_Titip Then
            psShowListPKDOT(True)
        End If
    End Sub

    Private Sub BtnListPKDOTClose_Click(sender As Object, e As EventArgs) Handles BtnListPKDOTClose.Click
        psShowListPKDOT(False)
    End Sub

    Private Sub BtnListPKDOTFind_Click(sender As Object, e As EventArgs) Handles BtnListPKDOTFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListPKDOT(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListPKDOT_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListPKDOT.PageIndexChanging
        GrvListPKDOT.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListPKDOT(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListPKDOT_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListPKDOT.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "PKDOTNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRowList As GridViewRow = GrvListPKDOT.Rows(vnIdx)

            Dim vnCompanyCode As String = Trim(vnGRowList.Cells(ensColListPKDOT.PKDOTCompanyCode).Text)
            Dim vnPKDOTNo As String = DirectCast(vnGRowList.Cells(ensColListPKDOT.PKDOTNo).Controls(0), LinkButton).Text

            DstCompany.SelectedValue = vnCompanyCode
            HdfCompanyCode.Value = vnCompanyCode

            TxtPCLRefNo.Text = vnPKDOTNo
            HdfPCLRefOID.Value = vnGRowList.Cells(ensColListPKDOT.OID).Text
            TxtPCLRefOID.Text = vnGRowList.Cells(ensColListPKDOT.OID).Text

            HdfPickTypeOID.Value = enuSchDType.Perintah_Kirim_DO_Titip

            HdfWhs.Value = vnGRowList.Cells(ensColListPKDOT.WarehouseOID).Text
            HdfWhsDest.Value = vnGRowList.Cells(ensColListPKDOT.WarehouseOID_Dest).Text

            TxtPCLDescr.Text = vnGRowList.Cells(ensColListPKDOT.NotaNo).Text & " " & vnPKDOTNo & Chr(10) & vnGRowList.Cells(ensColListPKDOT.vCustomer).Text & Chr(10) & vnGRowList.Cells(ensColListPKDOT.WarehouseName).Text & Chr(10) & " ke " & vnGRowList.Cells(ensColListPKDOT.vWarehouseName_Dest).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvDetail(HdfPCLRefOID.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowListPKDOT(False)
        End If
    End Sub

    Protected Sub RdlPickType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdlPickType.SelectedIndexChanged
        psGetReferensi_OnNew()
    End Sub

    Protected Sub BtnVoidPCL_Click(sender As Object, e As EventArgs) Handles BtnVoidPCL.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Void_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda VOID Pick List No. " & TxtPCLNo.Text & " ?<br />WARNING : VOID Tidak Dapat Dibatalkan"
        HdfProcess.Value = "VoidPL"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Private Sub psVoidPCL()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psVoidPCL", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnPCLHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoPCLHeader_TR with(nolock) Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            HdfTransStatus.Value = vnTransStatus

            If vnTransStatus <> enuTCPICK.Picking_Done Then
                LblMsgError.Text = "Status <> Picking Done"
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

            psEnableVoid(vnPCLHOID, vnTransStatus, vnSQLConn)
            If HdfEnableVoid.Value = "0" Then
                LblMsgError.Text = "Void Gagal...Picking <> Picking Done"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
                Exit Sub
            End If

            Dim vnPCKHOID As String = HdfPickingHOID.Value

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            pbuSsoProcessDataKey(HdfProcessDataKey.Value, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoPCKHeader_TR set TransStatus=" & enuTCPCKG.Void & ",IsVoid=1,VoidUserOID=" & Session("UserOID") & ",VoidDatetime=getdate() Where OID=" & vnPCKHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusPCK...Start")
            pbuInsertStatusPCK(vnPCKHOID, enuTCPCKG.Void, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCK...End")

            vnQuery = "Update Sys_SsoPCLHeader_TR set TransStatus=" & enuTCPICK.Void & ",PCLVoidNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',VoidUserOID=" & Session("UserOID") & ",VoidDatetime=getdate() Where OID=" & vnPCLHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            '<---05 Sep 2023 baru dipindahin, asalnya after status diupdate jadi void
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("5")
            psUpdateQtyOnPicklist_Cancel(vsTextStream, vnPCLHOID, HdfPCLRefOID.Value, vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("6")
            '<<==05 Sep 2023 baru dipindahin, asalnya after status diupdate jadi void

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("4")
            vsTextStream.WriteLine("pbuInsertStatusPCL...Start")
            pbuInsertStatusPCL(vnPCLHOID, enuTCPICK.Void, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusPCL...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Picking Sukses")
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
                vnSQLTrans.Dispose()
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
    Private Sub psStatusRefresh()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfTransStatus.Value = fbuGetPCLTransStatus(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub DstWhs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstWhs.SelectedIndexChanged
        psGetReferensi_OnNew()
    End Sub

    Protected Sub DstWhsDest_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstWhsDest.SelectedIndexChanged
        psGetReferensi_OnNew()
    End Sub

    Protected Sub DstCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstCompany.SelectedIndexChanged
        psGetReferensi_OnNew()
    End Sub

    Private Sub psGetReferensi_OnNew()
        psClearMessage()

        If HdfActionStatus.Value = cbuActionNew Then
            ChkDest.Visible = False
            If RdlPickType.SelectedValue = "" Then
                RdlPickType.SelectedValue = enuSchDType.Invoice
            End If
            If RdlPickType.SelectedValue = enuSchDType.TRB Then
                psGetReferensi_TRB()
            ElseIf RdlPickType.SelectedValue = enuSchDType.Invoice Then
                psGetReferensi_Invoice()
                ChkDest.Visible = True
            ElseIf RdlPickType.SelectedValue = enuSchDType.DO_Titip Then
                psGetReferensi_DO_Titip()
            ElseIf RdlPickType.SelectedValue = enuSchDType.Perintah_Kirim_DO_Titip Then
                psGetReferensi_Perintah_Kirim_DO_Titip()
            End If
        End If
    End Sub

    Private Function fsPrintKe() As Boolean
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Return False
        End If
        Dim vnReturn As Boolean = False

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnHOID As String = TxtTransID.Text
            Dim vnPrintNo As Integer
            Dim vnQuery As String
            vnQuery = "Select PCLPrintNo From Sys_SsoPCLHeader_TR Where OID=" & vnHOID
            vnPrintNo = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPCLHeader_TR set PCLPrintNo=" & vnPrintNo & " Where OID=" & vnHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            psInsertPrintPCL(vnHOID, vnPrintNo, Session("UserOID"), vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            TxtPCLPrint.Text = vnPrintNo
            vnReturn = True
        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If
            vnReturn = False
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
        Return vnReturn
    End Function

    Private Sub psInsertPrintPCL(vriOID As Integer, vriPrintNo As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPCLPrint_TR(PCLHOID,PCLPrintNo,PCLPrintUserOID,PCLPrintDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & "," & vriPrintNo & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Protected Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvRPPrint(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psShowPrintHS(True)
    End Sub

    Private Sub psFillGrvRPPrint(vriSQLConn As SqlConnection)
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select tr.PCLPrintNo,um.UserName PCLPrintBy,"
        vnQuery += vbCrLf & "       Convert(varchar(11),tr.PCLPrintDatetime,106) + ' '+ Convert(varchar(11),tr.PCLPrintDatetime,108)vPCLPrintDatetime"
        vnQuery += vbCrLf & "       From Sys_SsoPCLPrint_TR tr"
        vnQuery += vbCrLf & "			 inner join Sys_SsoUser_MA um on um.OID=tr.PCLPrintUserOID"
        vnQuery += vbCrLf & "	   Where tr.PCLHOID=" & Val(TxtTransID.Text)
        vnQuery += vbCrLf & "Order by tr.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvPrintHS.DataSource = vnDtb
        GrvPrintHS.DataBind()
    End Sub

    Protected Sub BtnRPClose_Click(sender As Object, e As EventArgs) Handles BtnRPClose.Click
        psShowPrintHS(False)
    End Sub
End Class