Imports System.Data.SqlClient
Public Class WbfSsoStockCard
    Inherits System.Web.UI.Page
    Enum ensColLsRcvPO
        RcvPONo = 0
        vRcvPODate = 1
        RcvPOSupplierName = 2
        RcvPOTypeName = 3
        OID = 4
    End Enum

    Private Sub psDefaultDisplay()
        DivLsRcvPO.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsRcvPO.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Private Sub psShowLsRcvPO(vriBo As Boolean)
        If vriBo Then
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Reporting"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoStorageStock, vnSQLConn)

            pbuFillDstStorageType(DstListStorageType, False, vnSQLConn)
            pbuFillDstWarehouse(DstListWarehouse, False, vnSQLConn)
            pbuFillDstBuilding(DstListBuilding, False, vnSQLConn)
            pbuFillDstLantai(DstListLantai, False, vnSQLConn)
            pbuFillDstZona(DstListZona, False, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub DstListStorageType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstListStorageType.SelectedIndexChanged
        If DstListStorageType.SelectedValue = enuStorageType.Floor Then
            PanListRackN.Visible = True
            PanListRackY.Visible = False
            PanListStagging.Visible = False
        ElseIf DstListStorageType.SelectedValue = enuStorageType.Rack Then
            PanListRackN.Visible = False
            PanListRackY.Visible = True
            PanListStagging.Visible = False
        ElseIf DstListStorageType.SelectedValue = enuStorageType.Staging Then
            PanListRackN.Visible = False
            PanListRackY.Visible = False
            PanListStagging.Visible = True
        Else
            PanListRackN.Visible = False
            PanListRackY.Visible = False
            PanListStagging.Visible = False
        End If
    End Sub

    Private Sub DstListWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstListWarehouse.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If DstListWarehouse.SelectedValue = "0" Then
            pbuFillDstBuilding(DstListBuilding, False, vnSQLConn)
        Else
            pbuFillDstBuilding_ByWarehouse(DstListBuilding, False, DstListWarehouse.SelectedValue, vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblMsgListWarehouse.Text = ""
        LblMsgListLantai.Text = ""
        LblMsgListStorageType.Text = ""
        LblMsgListBuilding.Text = ""
        LblMsgListZona.Text = ""
        LblMsgListRackN_Start.Text = ""
        LblMsgListRackY.Text = ""
        LblMsgListCompany.Text = ""
        LblMsgListBrg.Text = ""
        LblMsgListRcvNo.Text = ""
        LblStorageOID.ForeColor = System.Drawing.Color.Blue
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        psClearMessage()

        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        If ChkStorageOID.Checked Then
            If Val(TxtStorageOID.Text) = 0 Then
                LblStorageOID.ForeColor = System.Drawing.Color.Red
                Exit Sub
            End If
        Else
            If DstListWarehouse.SelectedIndex = 0 Then
                LblMsgListWarehouse.Text = "Pilih Warehouse"
                Exit Sub
            End If
            If ChkSummByWarehouse.Checked = False Then
                If DstListLantai.SelectedIndex = 0 Then
                    LblMsgListLantai.Text = "Pilih Lantai"
                    Exit Sub
                End If
                If DstListStorageType.SelectedIndex = 0 Then
                    LblMsgListStorageType.Text = "Pilih Storage Type"
                    Exit Sub
                End If
                If DstListBuilding.SelectedIndex = 0 Then
                    LblMsgListBuilding.Text = "Pilih Building"
                    Exit Sub
                End If
                If DstListZona.SelectedIndex = 0 Then
                    LblMsgListZona.Text = "Pilih Zona"
                    Exit Sub
                End If
                If DstListZona.SelectedIndex = 0 Then
                    LblMsgListZona.Text = "Pilih Zona"
                    Exit Sub
                End If
                If PanListRackN.Visible = True Then
                    If Trim(TxtListRackN_Start.Text) = "" Then
                        LblMsgListRackN_Start.Text = "Isi Storage Number"
                        Exit Sub
                    End If
                End If
                If PanListRackY.Visible = True Then
                    If Trim(TxtListRackY_SeqNo.Text) = "" Or Trim(TxtListRackY_Column.Text) = "" Or Trim(TxtListRackY_Level.Text) = "" Then
                        LblMsgListRackY.Text = "Isi Sequence,Column dan Level"
                        Exit Sub
                    End If
                End If
            End If
            If ChkSummAllRcv.Checked = False Then
                If HdfListRcvPOHOID.Value = "0" Then
                    LblMsgListRcvNo.Text = "Isi Nomor Penerimaan"
                    Exit Sub
                End If
            End If
            If Trim(TxtListBrgCode.Text) = "" Then
                LblMsgListBrg.Text = "Isi Kode Barang"
                Exit Sub
            End If
        End If
        psFillGrvList()
    End Sub

    Private Sub psFillGrvList()
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
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrRcvPOHOID As String = HdfListRcvPOHOID.Value

        Dim vnQuery As String

        Dim vnStorageOID As String

        If ChkStorageOID.Checked Then
            vnStorageOID = TxtStorageOID.Text
        Else
            If ChkSummByWarehouse.Checked Then
                vnStorageOID = 0
            Else
                vnQuery = "Select vStorageOID From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm"
                vnQuery += vbCrLf & "Where pm.WarehouseOID=" & DstListWarehouse.SelectedValue
                vnQuery += vbCrLf & "      and pm.BuildingOID=" & DstListBuilding.SelectedValue
                vnQuery += vbCrLf & "      and pm.LantaiOID=" & DstListLantai.SelectedValue
                vnQuery += vbCrLf & "      and pm.ZonaOID=" & DstListZona.SelectedValue
                vnQuery += vbCrLf & "      and pm.StorageTypeOID=" & DstListStorageType.SelectedValue

                If DstListStorageType.SelectedValue = enuStorageType.Floor Then
                    vnQuery += vbCrLf & "      and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(TxtListRackN_Start.Text)) & "'"
                ElseIf DstListStorageType.SelectedValue = enuStorageType.Rack Then
                    vnQuery += vbCrLf & "      and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(TxtListRackY_SeqNo.Text)) & "'"
                    vnQuery += vbCrLf & "      and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(TxtListRackY_Column.Text)) & "'"
                    vnQuery += vbCrLf & "      and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(TxtListRackY_Level.Text)) & "'"
                ElseIf DstListStorageType.SelectedValue = enuStorageType.Staging Then
                    vnQuery += vbCrLf & "      and pm.StorageStagIO=" & RdbListStagging.SelectedValue
                End If
                vnStorageOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)
            End If
        End If

        Dim vnFromCriteria As String

        vnFromCriteria = " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoStockCard_TR sm with(nolock) on sm.StorageOID=pm.vStorageOID"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR rc with(nolock) on rc.OID=sm.RcvPOHOID"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoTransName_MA tm with(nolock) on tm.TransCode=sm.TransCode"
        vnFromCriteria += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
        vnFromCriteria += vbCrLf & "Where 1=1"

        vnFromCriteria += vbCrLf & "            and mb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnFromCriteria += vbCrLf & "            and mb.BRGCODE='" & vnCrBrgCode & "'"

        If ChkSummAllRcv.Checked = False Then
            vnFromCriteria += vbCrLf & "            and sm.RcvPOHOID=" & vnCrRcvPOHOID
        End If
        If vnStorageOID = 0S Then
            vnFromCriteria += vbCrLf & "            and pm.WarehouseOID=" & DstListWarehouse.SelectedValue
        Else
            vnFromCriteria += vbCrLf & "            and pm.vStorageOID=" & vnStorageOID
        End If

        Dim vnDtb As New DataTable
        vnQuery = "Select * From("
        vnQuery += vbCrLf & "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
        vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
        vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
        vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
        vnQuery += vbCrLf & "     pm.vStorageOID,isnull(sm.OID,0)vStockCardOID,"
        vnQuery += vbCrLf & "     pm.vStorageInfoHtml,mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
        vnQuery += vbCrLf & "     sm.TransCode,tm.TransName,sm.TransOID,"
        vnQuery += vbCrLf & "     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,"
        vnQuery += vbCrLf & "     sm.TransQty,"
        vnQuery += vbCrLf & "     0 vQtyOnHand"
        vnQuery += vbCrLf & vnFromCriteria

        vnQuery += vbCrLf & "UNION"

        vnQuery += vbCrLf & "Select ''WarehouseName,''BuildingName,''LantaiDescription,''ZonaName,"
        vnQuery += vbCrLf & "     ''StorageTypeName,''vIsRack,''StorageSequenceNumber,''StorageColumn,''StorageLevel,''StorageNumber,"
        vnQuery += vbCrLf & "     ''vStorageStagIO,"
        vnQuery += vbCrLf & "     Null vStorageOID,Null vStockCardOID,"
        vnQuery += vbCrLf & "     ''vStorageInfoHtml,''CompanyCode,'',''vRcvPODate,''BRGCODE,''BRGNAME,"
        vnQuery += vbCrLf & "     ''TransCode,'TOTAL'TransName,Null TransOID,"
        vnQuery += vbCrLf & "     ''vCreationDatetime,"
        vnQuery += vbCrLf & "     sum(sm.TransQty)TransQty,"
        vnQuery += vbCrLf & "     dbo.fnSsoGet_StorageStock_QtyOnHand_ByKey('" & vnStorageOID & "','" & DstListCompany.SelectedValue & "','" & vnCrBrgCode & "'," & vnCrRcvPOHOID & ")vQtyOnHand"
        vnQuery += vbCrLf & vnFromCriteria

        vnQuery += vbCrLf & ")tb Order by case when isnull(vStockCardOID,0)=0 then 19 else case when TransCode='" & stuTransCode.SsoStockOB & "' then 4 else 5 end end,vStockCardOID"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")

        If vnDtb.Rows.Count > 0 Then
            GrvList.Rows(GrvList.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
        End If
    End Sub

    Private Sub psFillGrvList_20230918_Bef_ByWhs()
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
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrRcvPOHOID As String = HdfListRcvPOHOID.Value

        Dim vnQuery As String

        Dim vnStorageOID As String

        If ChkStorageOID.Checked Then
            vnStorageOID = TxtStorageOID.Text
        Else
            vnQuery = "Select vStorageOID From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm"
            vnQuery += vbCrLf & "Where pm.WarehouseOID=" & DstListWarehouse.SelectedValue
            vnQuery += vbCrLf & "      and pm.BuildingOID=" & DstListBuilding.SelectedValue
            vnQuery += vbCrLf & "      and pm.LantaiOID=" & DstListLantai.SelectedValue
            vnQuery += vbCrLf & "      and pm.ZonaOID=" & DstListZona.SelectedValue
            vnQuery += vbCrLf & "      and pm.StorageTypeOID=" & DstListStorageType.SelectedValue

            If DstListStorageType.SelectedValue = enuStorageType.Floor Then
                vnQuery += vbCrLf & "      and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(TxtListRackN_Start.Text)) & "'"
            ElseIf DstListStorageType.SelectedValue = enuStorageType.Rack Then
                vnQuery += vbCrLf & "      and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(TxtListRackY_SeqNo.Text)) & "'"
                vnQuery += vbCrLf & "      and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(TxtListRackY_Column.Text)) & "'"
                vnQuery += vbCrLf & "      and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(TxtListRackY_Level.Text)) & "'"
            ElseIf DstListStorageType.SelectedValue = enuStorageType.Staging Then
                vnQuery += vbCrLf & "      and pm.StorageStagIO=" & RdbListStagging.SelectedValue
            End If
            vnStorageOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)
        End If

        Dim vnFromCriteria As String

        vnFromCriteria = " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "')pm"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoStockCard_TR sm with(nolock) on sm.StorageOID=pm.vStorageOID"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR rc with(nolock) on rc.OID=sm.RcvPOHOID"
        vnFromCriteria += vbCrLf & "      inner join Sys_SsoTransName_MA tm with(nolock) on tm.TransCode=sm.TransCode"
        vnFromCriteria += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
        vnFromCriteria += vbCrLf & "Where 1=1"

        vnFromCriteria += vbCrLf & "            and mb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnFromCriteria += vbCrLf & "            and mb.BRGCODE='" & vnCrBrgCode & "'"
        vnFromCriteria += vbCrLf & "            and sm.RcvPOHOID=" & vnCrRcvPOHOID
        vnFromCriteria += vbCrLf & "            and pm.vStorageOID=" & vnStorageOID

        Dim vnDtb As New DataTable
        vnQuery = "Select * From("
        vnQuery += vbCrLf & "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
        vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
        vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
        vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
        vnQuery += vbCrLf & "     pm.vStorageOID,isnull(sm.OID,0)vStockCardOID,"
        vnQuery += vbCrLf & "     pm.vStorageInfoHtml,mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
        vnQuery += vbCrLf & "     sm.TransCode,tm.TransName,sm.TransOID,"
        vnQuery += vbCrLf & "     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,"
        vnQuery += vbCrLf & "     sm.TransQty,"
        vnQuery += vbCrLf & "     0 vQtyOnHand"
        vnQuery += vbCrLf & vnFromCriteria

        vnQuery += vbCrLf & "UNION"

        vnQuery += vbCrLf & "Select ''WarehouseName,''BuildingName,''LantaiDescription,''ZonaName,"
        vnQuery += vbCrLf & "     ''StorageTypeName,''vIsRack,''StorageSequenceNumber,''StorageColumn,''StorageLevel,''StorageNumber,"
        vnQuery += vbCrLf & "     ''vStorageStagIO,"
        vnQuery += vbCrLf & "     Null vStorageOID,Null vStockCardOID,"
        vnQuery += vbCrLf & "     ''vStorageInfoHtml,''CompanyCode,'',''vRcvPODate,''BRGCODE,''BRGNAME,"
        vnQuery += vbCrLf & "     ''TransCode,'TOTAL'TransName,Null TransOID,"
        vnQuery += vbCrLf & "     ''vCreationDatetime,"
        vnQuery += vbCrLf & "     sum(sm.TransQty)TransQty,"
        vnQuery += vbCrLf & "     dbo.fnSsoGet_StorageStock_QtyOnHand_ByKey('" & vnStorageOID & "','" & DstListCompany.SelectedValue & "','" & vnCrBrgCode & "'," & vnCrRcvPOHOID & ")vQtyOnHand"
        vnQuery += vbCrLf & vnFromCriteria

        vnQuery += vbCrLf & ")tb Order by case when isnull(vStockCardOID,0)=0 then 19 else case when TransCode='" & stuTransCode.SsoStockOB & "' then 4 else 5 end end,vStockCardOID"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")

        If vnDtb.Rows.Count > 0 Then
            GrvList.Rows(GrvList.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
        End If
    End Sub
    Protected Sub BtnListBrgCode_Click(sender As Object, e As EventArgs) Handles BtnListBrgCode.Click
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psShowLsBrg(True)
    End Sub

    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnLsBrgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBrgClose.Click
        psShowLsBrg(False)
    End Sub

    Private Sub psFillGrvLsBrg()
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
        vnQuery = "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstBarang_MA PM"
        vnQuery += vbCrLf & "Where CompanyCode='" & DstListCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtLsBrg.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtLsBrg.Text) & "%')"
        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsBrg.DataSource = vnDtb
        GrvLsBrg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvLsBrg_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsBrg.PageIndexChanging
        GrvLsBrg.PageIndex = e.NewPageIndex
        psFillGrvLsBrg()
    End Sub

    Private Sub GrvLsBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsBrg.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsBrg.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(0).Controls(0), LinkButton).Text
            TxtListBrgCode.Text = vnValue
            TxtListBrgName.Text = vnRow.Cells(1).Text
            psShowLsBrg(False)
        End If
    End Sub

    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psFillGrvLsBrg()
    End Sub

    Protected Sub BtnLsRcvPOClose_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOClose.Click
        psShowLsRcvPO(False)
    End Sub

    Protected Sub BtnListRcvNo_Click(sender As Object, e As EventArgs) Handles BtnListRcvNo.Click
        psShowLsRcvPO(True)
    End Sub

    Protected Sub BtnLsRcvPOFind_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOFind.Click
        psClearMessage()

        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            LblMsgListCompany.Visible = True
            Exit Sub
        End If

        psFillGrvLsRcvPO()
    End Sub

    Private Sub GrvLsRcvPO_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsRcvPO.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsRcvPO.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(ensColLsRcvPO.RcvPONo).Controls(0), LinkButton).Text
            TxtListRcvNo.Text = vnValue
            HdfListRcvPOHOID.Value = vnRow.Cells(ensColLsRcvPO.OID).Text
            psShowLsRcvPO(False)
        End If
    End Sub

    Private Sub GrvLsRcvPO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsRcvPO.PageIndexChanging
        GrvLsRcvPO.PageIndex = e.NewPageIndex
        psFillGrvLsRcvPO()
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
        vnQuery = "Select PM.RcvPONo,Convert(varchar(11),PM.RcvPODate)vRcvPODate,PM.RcvPOSupplierName,PT.RcvTypeName,PM.OID"
        vnQuery += vbCrLf & " From Sys_SsoRcvPOHeader_TR PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoRcvType_MA PT on PT.OID=PM.RcvRefTypeOID"
        vnQuery += vbCrLf & "Where PM.RcvPOCompanyCode='" & DstListCompany.SelectedValue & "' and PM.RcvPONo like '%" & Trim(TxtLsRcvPONo.Text) & "%'"
        vnQuery += vbCrLf & " Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsRcvPO.DataSource = vnDtb
        GrvLsRcvPO.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkStorageOID_CheckedChanged(sender As Object, e As EventArgs) Handles ChkStorageOID.CheckedChanged
        If ChkStorageOID.Checked Then
            psEnableStorageOID(True)
        Else
            psEnableStorageOID(False)
        End If
    End Sub

    Private Sub psEnableStorageOID(vriBo As Boolean)
        ChkSummByWarehouse.Checked = Not vriBo

        LblStorageOID.Visible = vriBo
        TxtStorageOID.Visible = vriBo

        DstListBuilding.Enabled = Not vriBo
        DstListLantai.Enabled = Not vriBo
        DstListStorageType.Enabled = Not vriBo
        DstListWarehouse.Enabled = Not vriBo
        DstListZona.Enabled = Not vriBo

        If vriBo Then
            PanListRackN.Visible = Not vriBo
            PanListRackY.Visible = Not vriBo
            PanListStagging.Visible = Not vriBo
        End If
    End Sub
    Private Sub psEnableWarehouseOID(vriBo As Boolean)
        If vriBo Then
            ChkStorageOID.Checked = Not vriBo
            LblStorageOID.Visible = vriBo
            TxtStorageOID.Visible = vriBo
        End If

        DstListBuilding.Enabled = Not vriBo
        DstListLantai.Enabled = Not vriBo
        DstListStorageType.Enabled = Not vriBo
        DstListWarehouse.Enabled = vriBo
        DstListZona.Enabled = Not vriBo

        If vriBo Then
            PanListRackN.Visible = Not vriBo
            PanListRackY.Visible = Not vriBo
            PanListStagging.Visible = Not vriBo
        End If
    End Sub

    Protected Sub BtnXLS_Click(sender As Object, e As EventArgs) Handles BtnXLS.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnFileName As String = ""
        pbuCreateXlsx_StockCard(vnFileName, Session("UserOID"), DstListWarehouse, DstListCompany,
                                          HdfListRcvPOHOID.Value, TxtListBrgCode, ChkStorageOID, TxtStorageOID,
                                       RdbListStagging, DstListBuilding, DstListLantai,
                                       DstListZona, DstListStorageType,
                                       TxtListRackN_Start, TxtListRackY_SeqNo, TxtListRackY_Column, TxtListRackY_Level, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkSummAllRcv_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSummAllRcv.CheckedChanged
        If ChkSummAllRcv.Checked Then
            BtnListRcvNo.Enabled = False
            TxtListRcvNo.Text = ""
            HdfListRcvPOHOID.Value = "0"
        Else
            BtnListRcvNo.Enabled = True
        End If
    End Sub

    Private Sub ChkSummByWarehouse_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSummByWarehouse.CheckedChanged
        If ChkSummByWarehouse.Checked Then
            psEnableWarehouseOID(True)
        Else
            psEnableWarehouseOID(False)
        End If
    End Sub
End Class