Imports System.Data.SqlClient

Public Class WbfSsoSmTRB
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSmTRB"

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
    Enum ensColListSO
        ChkCompany = 0
        CompanyCode = 1
        SalesOrderNo = 2
        vSalesOrderDate = 3
        vSUB = 4
        NAMA_CUSTOMER = 5
        GDGOJL = 6
        BRG = 7
        NAMA_BARANG = 8
        QTY = 9
        vSalesOrderDOID = 10
        SalesOrderHOID = 11
    End Enum

    Enum ensColDetail
        OID = 0
        vAddItem = 1
        CompanyCode = 2
        SalesOrderNo = 3
        vSalesOrderDate = 4
        vSUB = 5
        NAMA_CUSTOMER = 6
        GDGOJL = 7
        BRG = 8
        NAMA_BARANG = 9
        QTY = 10
        SalesOrderDOID = 11
        SalesOrderHOID = 12
        vDelItem = 13
    End Enum
    Enum ensColSum
        BRGCODE = 0
        BRGNAME = 1
        vQty_SO = 2
        vQty_PCL = 3
        vQty_Stock = 4
        vQty_RequestTRB = 5
        vQty_Avail_Wh_Dest = 6
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtSmDate.Text = ""
        TxtSmNote.Text = ""

        HdfTransStatus.Value = enuTCSMTB.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListSO.Style(HtmlTextWriterStyle.Position) = "absolute"

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoSummaryTRB, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhsFrom, True, vnSQLConn)
            pbuFillDstWarehouse(DstListWhsTo, True, vnSQLConn)

            psDisplayData_OnLoad(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psDisplayData_OnLoad(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")
        Dim vnUserOID As String = Session("UserOID")
        Dim vnQuery As String

        vnQuery = "Select PM.OID"
        vnQuery += vbCrLf & "     From Sys_SsoSmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "          inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.CompanyCode=PM.SmTRBCompanyCode and sw.OID=PM.SubWarehouseOID_From"
        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.SmTRBCompanyCode and uc.UserOID=" & vnUserOID
        End If
        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=sw.WarehouseOID and uw.UserOID=" & vnUserOID
        End If
        vnQuery += vbCrLf & "Where PM.TransStatus=" & enuTCSMTB.Baru

        Dim vnHOID As String = fbuGetDataNumSQL(vnQuery, vriSQLConn)
        If vnHOID > 0 Then
            TxtTransID.Text = vnHOID
            psDisplayData(vriSQLConn)

            psRefreshDetail(vnHOID, vriSQLConn)
            psFillGrvSum(vnHOID, vriSQLConn)
        End If
    End Sub

    Private Sub psRefreshDetail(vriSmTRBHOID As String, vriSQLConn As SqlConnection)
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnQuery As String
            Dim vn As Integer

            Dim vnDtbSO As New DataTable
            Dim vnSalesOrderDOID As Integer
            Dim vnSalesOrderHOID As Integer

            vnQuery = "Select distinct smd.SourceDOID,sod.SalesOrderHOID"
            vnQuery += vbCrLf & "       From Sys_SsoSmTRBDetail_TR smd"
            vnQuery += vbCrLf & "	         inner join Sys_SsoSalesOrderDetail_TR sod on sod.OID=smd.SourceDOID"
            vnQuery += vbCrLf & "	   Where smd.SmTRBHOID=" & vriSmTRBHOID & " and smd.SmTRBDTypeOID=" & enuSmTRBDTypeOID.Sales_Order
            pbuFillDtbSQL(vnDtbSO, vnQuery, vriSQLConn)

            vnSQLTrans = vriSQLConn.BeginTransaction()
            vnBeginTrans = True

            For vn = 0 To vnDtbSO.Rows.Count - 1
                vnSalesOrderDOID = vnDtbSO.Rows(vn).Item("SourceDOID")
                vnSalesOrderHOID = vnDtbSO.Rows(vn).Item("SalesOrderHOID")

                vnQuery = "Update Sys_SsoSalesOrderDetail_TR Set QTY_TRB=0 Where OID=" & vnSalesOrderDOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vnSQLTrans)

                vnQuery = "Update Sys_SsoSalesOrderHeader_TR Set TransStatus=" & enuTCCSSO.Baru
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vnSQLTrans)
            Next

            vnQuery = "Delete Sys_SsoSmTRBDetail_TR Where SmTRBHOID=" & vriSmTRBHOID
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)

            psSaveDetail(vriSmTRBHOID, vriSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If
        End Try
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

        If ChkSt_Baru.Checked = False And ChkSt_Closed.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_Baru.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCSMTB.Baru & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += enuTCSMTB.Closed_Sudah_TRB & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCSMTB.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,convert(varchar(11),PM.SmTRBDate,106)vSmTRBDate,"
        vnQuery += vbCrLf & "     PM.SmTRBCompanyCode,SW1.SubWhsName vSubWarehouseFrom,SW2.SubWhsName vSubWarehouseTo,"
        vnQuery += vbCrLf & "     PM.SmTRBNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ClosedDatetime,106)+' '+convert(varchar(5),PM.ClosedDatetime,108)+' '+ PR.UserName vClosed"

        vnQuery += vbCrLf & "From Sys_SsoSmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_SubWarehouse_MA SW1 with(nolock) on SW1.OID=PM.SubWarehouseOID_From"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_SubWarehouse_MA SW2 with(nolock) on SW2.OID=PM.SubWarehouseOID_To"

        If DstListWhsFrom.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WH1 with(nolock) on WH1.OID=SW1.WarehouseOID"
        End If
        If DstListWhsTo.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WH2 with(nolock) on WH2.OID=SW2.WarehouseOID"
        End If

        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.ClosedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListSalesOrderNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.OID in("
            vnQuery += vbCrLf & "              Select pod.SmTRBHOID"
            vnQuery += vbCrLf & "                     From Sys_SsoSmTRBDetail_TR pod with(nolock)"
            vnQuery += vbCrLf & "                          inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.OID=pod.SalesOrderDOID"
            vnQuery += vbCrLf & "                          inner join Sys_SsoSalesOrderHeader_TR soh with(nolock) on soh.OID=sod.SalesOrderHOID Where soh.SalesOrderNo like '%" & Trim(TxtListSalesOrderNo.Text) & "%')"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.SmTRBDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.SmTRBDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhsFrom.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and SW1.WarehouseOID= " & DstListWhsFrom.SelectedValue
        End If
        If DstListWhsTo.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and SW2.WarehouseOID = " & DstListWhsTo.SelectedValue
        End If

        vnQuery += vbCrLf & "Order by PM.SmTRBDate"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvDetail_SO(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = "0" Then
            vnQuery = "Select 0 OID,''vAddItem,''CompanyCode,''SalesOrderNo,''vSalesOrderDate,''vSUB,''NAMA_CUSTOMER,''GDGOJL,"
            vnQuery += vbCrLf & "            ''BRG,''NAMA_BARANG,0 QTY,"
            vnQuery += vbCrLf & "            0 SourceDOID,0 SalesOrderHOID,''vDelItem"
            vnQuery += vbCrLf & " Where 1=2"
        Else
            vnQuery = "Select smd.OID,''vAddItem,soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,soh.SUB vSUB,soh.NAMA_CUSTOMER,sod.GDGOJL,"
            vnQuery += vbCrLf & "            sod.BRG,sod.NAMA_BARANG,sod.QTY,"
            vnQuery += vbCrLf & "            smd.SourceDOID,sod.SalesOrderHOID,'Hapus Item'vDelItem"
            vnQuery += vbCrLf & "       From Sys_SsoSmTRBDetail_TR smd with(nolock)"
            vnQuery += vbCrLf & "            inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.OID=smd.SourceDOID"
            vnQuery += vbCrLf & "            inner join Sys_SsoSalesOrderHeader_TR soh with(nolock) on soh.OID=sod.SalesOrderHOID"
            vnQuery += vbCrLf & " Where smd.SmTRBHOID=" & vriHOID & " and smd.SmTRBDTypeOID=" & enuSmTRBDTypeOID.Sales_Order
            vnQuery += vbCrLf & "Order by sod.NAMA_BARANG,soh.SalesOrderNo"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvDetail_SO.DataSource = vnDtb
        GrvDetail_SO.DataBind()

        GrvDetail_SO.Visible = True
        GrvSum.Visible = False
    End Sub

    Private Sub psFillGrvDetail_20230930_Orig(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnOID As String = "0"
        Dim vnvAddItem As String = "..."
        Dim vnCompanyCode As String = ""
        Dim vnSalesOrderNo As String = ""
        Dim vnvSalesOrderDate As String = ""
        Dim vnvSUB As String = ""
        Dim vnNAMA_CUSTOMER As String = ""
        Dim vnGDGOJL As String = ""
        Dim vnBRG As String = ""
        Dim vnNAMA_BARANG As String = ""
        Dim vnQTY As String = "0"
        Dim vnSourceDOID As String = "0"
        Dim vnSalesOrderHOID As String = "0"
        Dim vnvDelItem As String = ""

        If vriHOID = "0" Then
            vnQuery = "Select 0 OID,''vAddItem,''CompanyCode,''SalesOrderNo,''vSalesOrderDate,''vSUB,''NAMA_CUSTOMER,''GDGOJL,"
            vnQuery += vbCrLf & "            ''BRG,''NAMA_BARANG,0 QTY,"
            vnQuery += vbCrLf & "            0 SourceDOID,0 SalesOrderHOID,''vDelItem"
            vnQuery += vbCrLf & " Where 1=2"
        Else
            vnQuery = "Select smd.OID,''vAddItem,soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,soh.SUB vSUB,soh.NAMA_CUSTOMER,sod.GDGOJL,"
            vnQuery += vbCrLf & "            sod.BRG,sod.NAMA_BARANG,sod.QTY,"
            vnQuery += vbCrLf & "            smd.SourceDOID,sod.SalesOrderHOID,'Hapus Item'vDelItem"
            vnQuery += vbCrLf & "       From Sys_SsoSmTRBDetail_TR smd with(nolock)"
            vnQuery += vbCrLf & "            inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.OID=smd.SourceDOID"
            vnQuery += vbCrLf & "            inner join Sys_SsoSalesOrderHeader_TR soh with(nolock) on soh.OID=sod.SalesOrderHOID"
            vnQuery += vbCrLf & " Where smd.SmTRBHOID=" & vriHOID & " and smd.SmTRBDTypeOID=" & enuSmTRBDTypeOID.Sales_Order
            vnQuery += vbCrLf & "Order by sod.NAMA_BARANG,soh.SalesOrderNo"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        Dim vn As Integer
        If HdfActionStatus.Value = cbuActionNorm Then
            GrvDetail_SO.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail_SO.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

            If HdfTransStatus.Value = enuTCSMTB.Baru Then
                GrvDetail_SO.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = ""
                GrvDetail_SO.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = ""
            Else
                GrvDetail_SO.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
                GrvDetail_SO.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"
            End If
        Else
            GrvDetail_SO.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
            GrvDetail_SO.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

            GrvDetail_SO.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail_SO.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

            For vn = 0 To 40
                vnDtb.Rows.Add(New Object() {vnOID, vnvAddItem, vnCompanyCode, vnSalesOrderNo, vnvSalesOrderDate, vnvSUB, vnNAMA_CUSTOMER, vnGDGOJL, vnBRG, vnNAMA_BARANG, vnQTY, vnSalesOrderHOID, vnSourceDOID, vnvDelItem})
            Next
        End If

        GrvDetail_SO.DataSource = vnDtb
        GrvDetail_SO.DataBind()

        GrvDetail_SO.Visible = True
        GrvSum.Visible = False
    End Sub

    Private Sub psFillGrvSum(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = "0" Then
            vnQuery = "Select ''BRGCODE,''BRGNAME,0 vQty_SO,0 vQty_PCL,0 vQty_Stock,0 vQty_RequestTRB,0 vQty_Avail_Wh_Dest"
            vnQuery += vbCrLf & " Where 1=2"
        Else
            Dim vnSubWhsCode As String = fbuGetSubWhsCode_ByOID(DstSubWhsFrom.SelectedValue, vriSQLConn)

            vnQuery = "Select mbr.BRGCODE,mbr.BRGNAME,smt.vQty_SO,smt.vQty_PCL,smt.vQty_Stock,smt.vQty_RequestTRB,"

            If HdfTransStatus.Value = enuTCSMTB.Baru Then
                vnQuery += vbCrLf & " vQty_Avail_Wh_Dest"
            Else
                vnQuery += vbCrLf & " 0 vQty_Avail_Wh_Dest"
            End If

            vnQuery += vbCrLf & " From fnTbl_SsoSmTRBSummary(" & vriHOID & ")smt"
            vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mbr with(nolock) on mbr.BRGCODE=smt.BRGCODE"

            If HdfTransStatus.Value = enuTCSMTB.Baru Then
                vnQuery += vbCrLf & "      inner join ("
                vnQuery += vbCrLf & "            Select sto.BRGCODE,sum(sto.vQtyAvailable)vQty_Avail_Wh_Dest"
                vnQuery += vbCrLf & "              From fnTbl_SsoStorageStock()sto"
                vnQuery += vbCrLf & "                   inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0)sti on sti.vStorageOID=sto.StorageOID"

                If vnSubWhsCode = stuWarehouse.Prancis_Baru Then
                    vnQuery += vbCrLf & "             Where sto.CompanyCode='" & DstCompany.SelectedValue & "' and sti.StorageTypeOID in(" & enuStorageType.Rack & "," & enuStorageType.Floor & ")"
                    vnQuery += vbCrLf & "                   and sti.WarehouseOID in(" & enuWarehouseOID.Prancis & "," & enuWarehouseOID.Prancis2 & ")"
                Else
                    vnQuery += vbCrLf & "                   inner join " & vnDBMaster & "Sys_SubWarehouse_MA swh with(nolock) on swh.CompanyCode=sto.CompanyCode and swh.WarehouseOID=sti.WarehouseOID"
                    vnQuery += vbCrLf & "             Where swh.CompanyCode='" & DstCompany.SelectedValue & "' and sti.StorageTypeOID in(" & enuStorageType.Rack & "," & enuStorageType.Floor & ")"
                    vnQuery += vbCrLf & "                   and swh.SubWhsCode='" & vnSubWhsCode & "'"
                End If

                vnQuery += vbCrLf & "             Group by sto.BRGCODE) sto on sto.BRGCODE=smt.BRGCODE"
            End If

            vnQuery += vbCrLf & "Where mbr.CompanyCode='" & DstCompany.SelectedValue & "'"

            If ChkSummNonZero.Checked Then
                vnQuery += vbCrLf & "      and smt.vQty_RequestTRB > 0"
            End If
            vnQuery += vbCrLf & " Order by mbr.BRGCODE"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSum.DataSource = vnDtb
        GrvSum.DataBind()

        GrvDetail_SO.Visible = False
        GrvSum.Visible = True

        If HdfTransStatus.Value = enuTCSMTB.Baru Then
            GrvSum.Columns(ensColSum.vQty_Avail_Wh_Dest).HeaderStyle.CssClass = ""
            GrvSum.Columns(ensColSum.vQty_Avail_Wh_Dest).ItemStyle.CssClass = ""
        Else
            GrvSum.Columns(ensColSum.vQty_Avail_Wh_Dest).HeaderStyle.CssClass = "myDisplayNone"
            GrvSum.Columns(ensColSum.vQty_Avail_Wh_Dest).ItemStyle.CssClass = "myDisplayNone"
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
            psFillGrvList()
            psButtonShowList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psButtonStatus()
        End If
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnSudahTRB.Enabled = False
        BtnCancelSm.Enabled = False
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

        TxtSmDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvSum(0, vnSQLConn)
        psFillGrvDetail_SO(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgSmDate.Text = ""
        LblMsgSubWhsFrom.Text = ""
        LblMsgSubWhsTo.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSmNote.ReadOnly = Not vriBo
        RdbDS.Enabled = Not vriBo

        If vriBo Then
            RdbDS.SelectedValue = "D"
        End If
        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = False
        BtnBaru.Visible = Not vriBo

        BtnSudahTRB.Visible = False
        BtnCancelSm.Visible = False
        BtnPreview.Visible = False

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail_SO.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvDetail_SO.Rows(vnIdx)
        If e.CommandName = "vAddItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If
            HdfDetailRowIdx.Value = vnIdx

            psShowListSO(True)

        ElseIf e.CommandName = "vDelItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If
            HdfDetailRowIdx.Value = vnIdx
            LblConfirmMessage.Text = "Anda Hapus Item " & vnGRow.Cells(ensColDetail.NAMA_BARANG).Text & " ?"
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
            vnParam += "&vqTrCode=" & stuTransCode.SsoSummaryTRB
            vnParam += "&vqTrNo=None"

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

            psFillGrvDetail_SO(0, vnSQLConn)

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

        vnQuery = "Select PM.*,convert(varchar(11),PM.SmTRBDate,106)vSmTRBDate,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoSmTRBHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoSummaryTRB & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtSmDate.Text = vnDtb.Rows(0).Item("vSmTRBDate")
            TxtSmNote.Text = vnDtb.Rows(0).Item("SmTRBNote")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("SmTRBCompanyCode"))

            pbuFillDstSubWarehouse_ByUserOID(Session("UserOID"), DstCompany.SelectedValue, DstSubWhsFrom, False, vriSQLConn)
            pbuFillDstSubWarehouse(DstSubWhsTo, False, vriSQLConn)

            DstSubWhsFrom.SelectedValue = Trim(vnDtb.Rows(0).Item("SubWarehouseOID_From"))
            DstSubWhsTo.SelectedValue = Trim(vnDtb.Rows(0).Item("SubWarehouseOID_To"))

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        psFillGrvSum(Val(TxtTransID.Text), vriSQLConn)

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnSudahTRB.Visible = BtnSudahTRB.Enabled
        BtnCancelSm.Visible = BtnCancelSm.Enabled
        BtnPreview.Visible = BtnPreview.Enabled

        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnSudahTRB.Enabled = False
        BtnCancelSm.Enabled = False
        BtnPreview.Enabled = False
        BtnRefreshSum.Enabled = False

        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = False

            BtnSudahTRB.Enabled = (HdfTransStatus.Value = enuTCSMTB.Baru)
            BtnCancelSm.Enabled = (HdfTransStatus.Value = enuTCSMTB.Baru)
            BtnRefreshSum.Enabled = (HdfTransStatus.Value = enuTCSMTB.Baru)

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
        psFillGrvDetail_SO(TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        If HdfTransStatus.Value = enuTCSMTB.Baru Then
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
        If DstSubWhsFrom.SelectedValue = "0" Then
            LblMsgSubWhsFrom.Text = "Pilih Gudang Asal"
            vnSave = False
        End If
        If DstSubWhsTo.SelectedValue = "0" Then
            LblMsgSubWhsTo.Text = "Pilih Gudang Tujuan"
            vnSave = False
        End If
        If Not IsDate(Trim(TxtSmDate.Text)) Then
            LblMsgSmDate.Text = "Isi Tanggal"
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

                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)
                Dim vnWhsOID_From As String = Trim(DstSubWhsFrom.SelectedValue)
                Dim vnWhsOID_To As String = Trim(DstSubWhsTo.SelectedValue)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoSmTRBHeader_TR with(nolock)"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoSmTRBHeader_TR(OID,SmTRBDate,"
                vnQuery += vbCrLf & "SmTRBCompanyCode,"
                vnQuery += vbCrLf & "SubWarehouseOID_From,SubWarehouseOID_To,"
                vnQuery += vbCrLf & "SmTRBNote,"
                vnQuery += vbCrLf & "TransCode,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & TxtSmDate.Text & "',"

                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "'," & vnWhsOID_From & "," & vnWhsOID_To & ","
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtSmNote.Text)) & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoSummaryTRB & "'," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnOID, vnSQLConn, vnSQLTrans)

                pbuInsertStatusSmTRB(vnOID, enuTCSMTB.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCSMTB.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

            Else
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoSmTRBHeader_TR set"
                vnQuery += vbCrLf & "SmTRBDate='" & TxtSmDate.Text & "',"
                vnQuery += vbCrLf & "SubWarehouseOID_From=" & DstSubWhsFrom.SelectedValue & ","
                vnQuery += vbCrLf & "SubWarehouseOID_To=" & DstSubWhsTo.SelectedValue & ","

                vnQuery += vbCrLf & "SmTRBNote='" & fbuFormatString(Trim(TxtSmNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                psSaveDetail(TxtTransID.Text, vnSQLConn, vnSQLTrans)

                pbuInsertStatusSmTRB(TxtTransID.Text, enuTCSMTB.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"
            End If

            psEnableInput(False)
            psEnableSave(False)

            HdfActionStatus.Value = cbuActionNorm
            RdbDS.SelectedValue = "S"
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
    Private Sub psSaveDetail(vriSmTRBHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnQuery As String

        Dim vnSubWhsCode As String = fbuGetSubWhsCode_ByOID_Trans(DstSubWhsTo.SelectedValue, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSmTRBDetail_TR"
        vnQuery += vbCrLf & "(SmTRBHOID,SmTRBDTypeOID,SourceDOID,BRGCODE,SourceDQty)"
        vnQuery += vbCrLf & "  Select " & vriSmTRBHOID & "," & enuSmTRBDTypeOID.Sales_Order & ",sod.OID,sod.BRG,sod.QTY"
        vnQuery += vbCrLf & " From Sys_SsoSalesOrderHeader_TR soh with(nolock) "
        vnQuery += vbCrLf & "      inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.SalesOrderHOID=soh.OID"
        vnQuery += vbCrLf & "Where soh.SOVoid=0 and sod.QTY > sod.QTY_TRB and soh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and sod.GDGOJL='" & vnSubWhsCode & "'"
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSmTRBDetail_TR"
        vnQuery += vbCrLf & "(SmTRBHOID,SmTRBDTypeOID,SourceDOID,BRGCODE,SourceDQty)"
        vnQuery += vbCrLf & "Select " & vriSmTRBHOID & "," & enuSmTRBDTypeOID.Picklist_Gantung & ",nd.OID,nd.KodeBarang,(nd.TotalQty + nd.TotalQtyBonus) - nd.TotalQtyOnPickList"
        vnQuery += vbCrLf & " From " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR nd with(nolock) on nd.NotaHOID=nh.OID"
        vnQuery += vbCrLf & "Where nh.IsPickListClosed=0 and nh.NotaCancel=0 and NotaDOT=0"
        vnQuery += vbCrLf & "      and (nd.TotalQty + nd.TotalQtyBonus) > nd.TotalQtyOnPickList"
        vnQuery += vbCrLf & "      and nh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and nh.GDG='" & vnSubWhsCode & "'"
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoSmTRBDetail_TR"
        vnQuery += vbCrLf & "(SmTRBHOID,SmTRBDTypeOID,SourceDOID,BRGCODE,SourceDQty)"
        vnQuery += vbCrLf & "  Select " & vriSmTRBHOID & "," & enuSmTRBDTypeOID.Stock_Gudang & ",0,sto.BRGCODE,sum(sto.vQtyAvailable)"
        vnQuery += vbCrLf & "	 From fnTbl_SsoStorageStock()sto"
        vnQuery += vbCrLf & "	      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0)sti on sti.vStorageOID=sto.StorageOID"
        vnQuery += vbCrLf & "	      inner join " & vnDBMaster & "Sys_SubWarehouse_MA swh with(nolock) on swh.CompanyCode=sto.CompanyCode and swh.WarehouseOID=sti.WarehouseOID"
        vnQuery += vbCrLf & "   Where swh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "         and swh.SubWhsCode='" & vnSubWhsCode & "'"
        vnQuery += vbCrLf & "         and sti.StorageTypeOID in(" & enuStorageType.Rack & "," & enuStorageType.Floor & ")"
        vnQuery += vbCrLf & "   Group by sto.BRGCODE"
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Update sod set QTY_TRB=QTY"
        vnQuery += vbCrLf & " From Sys_SsoSalesOrderHeader_TR soh with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.SalesOrderHOID=soh.OID"
        vnQuery += vbCrLf & "Where soh.SOVoid=0 and sod.QTY > sod.QTY_TRB and soh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and sod.GDGOJL='" & vnSubWhsCode & "'"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

        vnQuery = "Update soh set TransStatus=" & enuTCCSSO.In_TRB_Calculation
        vnQuery += vbCrLf & " From Sys_SsoSalesOrderHeader_TR soh with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.SalesOrderHOID=soh.OID"
        vnQuery += vbCrLf & "Where soh.SOVoid=0 and sod.QTY = sod.QTY_TRB and soh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      and sod.GDGOJL='" & vnSubWhsCode & "'"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Private Sub psSaveDetail_20230930_Orig(vriSmTRBHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow
        Dim vnSalesOrderDOID As String
        Dim vnSalesOrderHOID As String
        For vn = 0 To GrvDetail_SO.Rows.Count - 1
            vnGRow = GrvDetail_SO.Rows(vn)
            vnSalesOrderDOID = vnGRow.Cells(ensColDetail.SalesOrderDOID).Text
            vnSalesOrderHOID = vnGRow.Cells(ensColDetail.SalesOrderHOID).Text
            If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRG).Text) <> "" Then
                    vnQuery = "Select SalesOrderDOID From Sys_SsoSmTRBDetail_TR Where SmTRBHOID=" & vriSmTRBHOID & " and SalesOrderDOID=" & vnSalesOrderDOID
                    If fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = 0 Then
                        vnQuery = "Insert into Sys_SsoSmTRBDetail_TR"
                        vnQuery += vbCrLf & "(SmTRBHOID,SalesOrderDOID)"
                        vnQuery += vbCrLf & "values(" & vriSmTRBHOID & "," & vnSalesOrderDOID & ")"
                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                        vnQuery = "Update Sys_SsoSalesOrderDetail_TR Set QTY_TRB=QTY Where OID=" & vnSalesOrderDOID
                        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                        vnQuery = "Update Sys_SsoSalesOrderHeader_TR Set TransStatus=" & enuTCCSSO.In_TRB_Calculation & " Where OID=" & vnSalesOrderHOID
                        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
                    End If
                End If
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
            TxtTransID.Text = DirectCast(vnRow.Cells(ensColList.OID).Controls(0), LinkButton).Text

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
        If HdfProcess.Value = "CancelSm" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelSm()
        ElseIf HdfProcess.Value = "CloseSm" Then
            psCloseSm()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub BtnSudahTRB_Click(sender As Object, e As EventArgs) Handles BtnSudahTRB.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Close Sudah TRB untuk No " & TxtTransID.Text & " ?<br />WARNING : Close Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CloseSm"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub psCancelSm()
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

            Dim vnDtbSO As New DataTable
            Dim vnSalesOrderDOID As Integer
            Dim vnSalesOrderHOID As Integer

            vnQuery = "Select distinct smd.SourceDOID,sod.SalesOrderHOID"
            vnQuery += vbCrLf & "       From Sys_SsoSmTRBDetail_TR smd"
            vnQuery += vbCrLf & "	         inner join Sys_SsoSalesOrderDetail_TR sod on sod.OID=smd.SourceDOID"
            vnQuery += vbCrLf & "	   Where smd.SmTRBHOID=" & TxtTransID.Text & " and smd.SmTRBDTypeOID=" & enuSmTRBDTypeOID.Sales_Order
            pbuFillDtbSQL(vnDtbSO, vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoSmTRBHeader_TR set TransStatus=" & enuTCSMTB.Cancelled & ",SmTRBCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusSmTRB(TxtTransID.Text, enuTCSMTB.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

            For vn = 0 To vnDtbSO.Rows.Count - 1
                vnSalesOrderDOID = vnDtbSO.Rows(vn).Item("SourceDOID")
                vnSalesOrderHOID = vnDtbSO.Rows(vn).Item("SalesOrderHOID")

                vnQuery = "Update Sys_SsoSalesOrderDetail_TR Set QTY_TRB=0 Where OID=" & vnSalesOrderDOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnQuery = "Update Sys_SsoSalesOrderHeader_TR Set TransStatus=" & enuTCCSSO.Baru
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
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

    Private Sub psCloseSm()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psCloseSm", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnSmTRBHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoSmTRBHeader_TR set TransStatus=" & enuTCSMTB.Closed_Sudah_TRB & ",ClosedUserOID=" & Session("UserOID") & ",ClosedDatetime=getdate() Where OID=" & vnSmTRBHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSmTRB...Start")
            pbuInsertStatusSmTRB(vnSmTRBHOID, enuTCSMTB.Closed_Sudah_TRB, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSmTRB...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Close Sudah TRB Sukses")
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

    Protected Sub BtnCancelSm_Click(sender As Object, e As EventArgs) Handles BtnCancelSm.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan Summary Barang untuk TRB No " & TxtTransID.Text & " ?<br />WARNING : Batal Summary Barang untuk TRB No Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelSm"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Private Sub psGenerateCrp(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSmTRBSummary

        vbuCrpQuery = "Select sm.*,mc.CompanyName,"
        vbuCrpQuery += "      sw1.SubWhsCode vSubWhsCode_From,sw1.SubWhsName vSubWhsName_From,"
        vbuCrpQuery += "      sw2.SubWhsCode vSubWhsCode_To,sw2.SubWhsName vSubWhsName_To"
        vbuCrpQuery += " From fnTbl_SsoSmTRB_Summary(" & TxtTransID.Text & ",'" & Session("UserID") & "')sm"
        vbuCrpQuery += "      inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=sm.CompanyCode"

        vbuCrpQuery += "      inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw1 with(nolock) on sw1.OID=sm.SubWarehouseOID_From"
        vbuCrpQuery += "      inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw2 with(nolock) on sw2.OID=sm.SubWarehouseOID_To"
        vbuCrpQuery += " order by sm.NAMA_BARANG"
    End Sub
    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

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

    Private Sub psShowListSO(vriBo As Boolean)
        If vriBo Then
            DivListSO.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListSONo.Focus()
        Else
            DivListSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub
    Protected Sub BtnListPOFind_Click(sender As Object, e As EventArgs) Handles BtnListSOFind.Click
        LblMsgListPO.Text = ""

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListSO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvListSO(vriSQLConn As SqlConnection)
        LblMsgListPO.Text = ""

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String
        Dim vnSubWhsCode As String = fbuGetSubWhsCode_ByOID(DstSubWhsFrom.SelectedValue, vriSQLConn)
        Dim vnCustomer = fbuFormatString(Trim(TxtListSOCust.Text))

        vnQuery = "Select soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,soh.SUB vSUB,soh.NAMA_CUSTOMER,sod.GDGOJL,"
        vnQuery += vbCrLf & "            sod.BRG,sod.NAMA_BARANG,sod.QTY,"
        vnQuery += vbCrLf & "            sod.OID vSalesOrderDOID,sod.SalesOrderHOID"
        vnQuery += vbCrLf & "       From Sys_SsoSalesOrderHeader_TR soh with(nolock)"
        vnQuery += vbCrLf & "            inner join Sys_SsoSalesOrderDetail_TR sod with(nolock) on sod.SalesOrderHOID=soh.OID"

        vnCriteria = "      Where QTY>QTY_TRB"
        vnCriteria += vbCrLf & "            and soh.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnCriteria += vbCrLf & "            and sod.GDGOJL='" & vnSubWhsCode & "'"
        If vnCustomer <> "" Then
            vnCriteria += vbCrLf & "            and (soh.SUB like '%" & vnCustomer & "%' or soh.NAMA_CUSTOMER like '%" & vnCustomer & "%')"
        End If
        If Trim(TxtListSONo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderNo like '%" & fbuFormatString(Trim(TxtListSONo.Text)) & "%'"
        End If
        If IsDate(TxtSOStart.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate >= '" & TxtSOStart.Text & "'"
        End If
        If IsDate(TxtSOEnd.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate <= '" & TxtSOEnd.Text & "'"
        End If
        vnQuery += vbCrLf & vnCriteria
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListSO.DataSource = vnDtb
        GrvListSO.DataBind()

        TxtListSONo.Focus()
    End Sub

    Protected Sub BtnListSOClose_Click(sender As Object, e As EventArgs) Handles BtnListSOClose.Click
        psShowListSO(False)
    End Sub

    Private Sub GrvListSO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListSO.PageIndexChanging
        GrvListSO.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListSO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListSOSelect_Click(sender As Object, e As EventArgs) Handles BtnListSOSelect.Click
        If GrvListSO.Rows.Count > 0 Then
            Dim vn As Integer
            Dim vnRowIdxDetail As Integer

            For vn = 0 To GrvDetail_SO.Rows.Count - 1
                If fbuValStrHtml(GrvDetail_SO.Rows(vn).Cells(ensColDetail.BRG).Text) = "" Then
                    vnRowIdxDetail = vn
                    Exit For
                End If
            Next

            Dim vnGRowDetail As GridViewRow
            Dim vnGRowSO As GridViewRow
            Dim vnChkSelect As CheckBox
            For vn = 0 To GrvListSO.Rows.Count - 1
                vnGRowSO = GrvListSO.Rows(vn)
                vnChkSelect = vnGRowSO.FindControl("ChkSelect")
                If vnChkSelect.Checked Then
                    vnGRowDetail = GrvDetail_SO.Rows(vnRowIdxDetail)
                    vnGRowDetail.Cells(ensColDetail.CompanyCode).Text = vnGRowSO.Cells(ensColListSO.CompanyCode).Text
                    vnGRowDetail.Cells(ensColDetail.SalesOrderNo).Text = vnGRowSO.Cells(ensColListSO.SalesOrderNo).Text
                    vnGRowDetail.Cells(ensColDetail.vSalesOrderDate).Text = vnGRowSO.Cells(ensColListSO.vSalesOrderDate).Text
                    vnGRowDetail.Cells(ensColDetail.vSUB).Text = vnGRowSO.Cells(ensColListSO.vSUB).Text
                    vnGRowDetail.Cells(ensColDetail.NAMA_CUSTOMER).Text = vnGRowSO.Cells(ensColListSO.NAMA_CUSTOMER).Text
                    vnGRowDetail.Cells(ensColDetail.GDGOJL).Text = vnGRowSO.Cells(ensColListSO.GDGOJL).Text
                    vnGRowDetail.Cells(ensColDetail.BRG).Text = vnGRowSO.Cells(ensColListSO.BRG).Text
                    vnGRowDetail.Cells(ensColDetail.NAMA_BARANG).Text = vnGRowSO.Cells(ensColListSO.NAMA_BARANG).Text
                    vnGRowDetail.Cells(ensColDetail.QTY).Text = vnGRowSO.Cells(ensColListSO.QTY).Text
                    vnGRowDetail.Cells(ensColDetail.SalesOrderDOID).Text = vnGRowSO.Cells(ensColListSO.vSalesOrderDOID).Text
                    vnGRowDetail.Cells(ensColDetail.SalesOrderHOID).Text = vnGRowSO.Cells(ensColListSO.SalesOrderHOID).Text
                    vnRowIdxDetail = vnRowIdxDetail + 1
                End If
            Next
        End If
    End Sub

    Protected Sub RdbDS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdbDS.SelectedIndexChanged
        If BtnSimpan.Visible Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If RdbDS.SelectedValue = "D" Then
            psFillGrvDetail_SO(Val(TxtTransID.Text), vnSQLConn)
        Else
            psFillGrvSum(Val(TxtTransID.Text), vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub DstCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstCompany.SelectedIndexChanged
        If BtnSimpan.Visible = False Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuFillDstSubWarehouse_ByUserOID(Session("UserOID"), DstCompany.SelectedValue, DstSubWhsFrom, False, vnSQLConn)
        pbuFillDstSubWarehouse_ByCompanyCode(DstSubWhsTo, False, DstCompany.SelectedValue, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvDetail_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvDetail_SO.SelectedIndexChanged

    End Sub

    Protected Sub ChkSummNonZero_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSummNonZero.CheckedChanged
        If BtnSimpan.Visible Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        RdbDS.SelectedValue = "S"
        psFillGrvSum(Val(TxtTransID.Text), vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnRefreshSum_Click(sender As Object, e As EventArgs) Handles BtnRefreshSum.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnHOID As String = TxtTransID.Text

        psDisplayData(vnSQLConn)
        psRefreshDetail(vnHOID, vnSQLConn)
        psFillGrvSum(vnHOID, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class