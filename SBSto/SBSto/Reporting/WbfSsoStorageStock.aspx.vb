Imports System.Data.SqlClient
Public Class WbfSsoStorageStock
    Inherits System.Web.UI.Page

    Enum ensColList
        vStorageStockOID = 0
        vStorageOID = 1
        vStorageInfoHtml = 2
        CompanyCode = 3
        RcvPOHOID = 4
        RcvPONo = 5
        vRcvPODate = 6
        BRGCODE = 7
        BRGNAME = 8
    End Enum
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

        DivStCard.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivStCard.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanStCard.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivCheck.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivCheck.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanCheck.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            pbuFillDstStorageType(DstListStorageType, True, vnSQLConn)
            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)
            pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)
            pbuFillDstLantai(DstListLantai, True, vnSQLConn)
            pbuFillDstZona(DstListZona, True, vnSQLConn)

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
            pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)
        Else
            pbuFillDstBuilding_ByWarehouse(DstListBuilding, True, DstListWarehouse.SelectedValue, vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblMsgListCompany.Text = ""
        LblMsgListBrg.Text = ""
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        psClearMessage()

        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            Exit Sub
        End If
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
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
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select "
        vnQuery += vbCrLf & "     isnull(sm.OID,0)vStorageStockOID,pm.vStorageOID,pm.vStorageInfoHtml,"
        vnQuery += vbCrLf & "     mb.CompanyCode,sm.RcvPOHOID,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
        vnQuery += vbCrLf & "     sm.QtyOnHand,"
        vnQuery += vbCrLf & "     sm.vQtyAvailable,"
        vnQuery += vbCrLf & "     sm.QtyOnPutaway,sm.QtyOnPutawayWh,sm.QtyOnMovement,sm.QtyOnMovementWh,"
        vnQuery += vbCrLf & "     sm.QtyOnPickList,sm.QtyOnPicking,sm.QtyOnSgo,sm.QtyOnDispatch,sm.QtyOnKarantina,sm.QtyOnPutawayKr,"
        vnQuery += vbCrLf & "     sm.QtyOnPutawayDtw,sm.QtyOnPutawayDty,sm.QtyOnPutawayPtv,sm.QtyOnPutawayDsw,sm.QtyOnPutawayDsy"
        vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
        vnQuery += vbCrLf & "      inner join fnTbl_SsoStorageStock() sm on sm.StorageOID=pm.vStorageOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & "            and mb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "            and mb.BRGCODE like '%" & vnCrBrgCode & "%' and mb.BRGNAME like '%" & vnCrBrgName & "%'"
        If Trim(TxtListRcvNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and rc.RcvPONo like '%" & Trim(TxtListRcvNo.Text) & "%'"
        End If

        If ChkStorageOID.Checked Then
            vnQuery += vbCrLf & "            and pm.vStorageOID=" & Val(TxtStorageOID.Text)
        Else
            If Val(DstListWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and pm.WarehouseOID=" & DstListWarehouse.SelectedValue
            End If
            If Val(DstListBuilding.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and pm.BuildingOID=" & DstListBuilding.SelectedValue
            End If
            If Val(DstListLantai.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and pm.LantaiOID=" & DstListLantai.SelectedValue
            End If
            If Val(DstListZona.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and pm.ZonaOID=" & DstListZona.SelectedValue
            End If
            If Val(DstListStorageType.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and pm.StorageTypeOID=" & DstListStorageType.SelectedValue
            End If

            If DstListStorageType.SelectedValue = enuStorageType.Rack Then
                If Trim(TxtListRackY_SeqNo.Text) <> "" Then
                    vnQuery += vbCrLf & "            and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(TxtListRackY_SeqNo.Text)) & "'"
                End If
                If Trim(TxtListRackY_Column.Text) <> "" Then
                    vnQuery += vbCrLf & "            and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(TxtListRackY_Column.Text)) & "'"
                End If
                If Trim(TxtListRackY_Level.Text) <> "" Then
                    vnQuery += vbCrLf & "            and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(TxtListRackY_Level.Text)) & "'"
                End If
                vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,mb.BRGCODE"

            ElseIf DstListStorageType.SelectedValue = enuStorageType.Floor Then
                If Trim(TxtListRackN_Start.Text) <> "" Then
                    vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(TxtListRackN_Start.Text)) & "'"
                End If
                If Trim(TxtListRackN_End.Text) <> "" Then
                    vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')<='" & fbuFormatString(Trim(TxtListRackN_End.Text)) & "'"
                End If
                vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageNumber"

            ElseIf DstListStorageType.SelectedValue = enuStorageType.Staging Then
                vnQuery += vbCrLf & "            and pm.StorageStagIO=" & RdbListStagging.SelectedValue
                vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
            Else
                vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
            End If
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub
    Protected Sub BtnListBrgCode_Click(sender As Object, e As EventArgs) Handles BtnListBrgCode.Click
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psShowLsBrg(True)
    End Sub
    Private Sub psShowLsRcvPO(vriBo As Boolean)
        If vriBo Then
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowCheck(vriBo As Boolean)
        If vriBo Then
            DivCheck.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivCheck.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowStCard(vriBo As Boolean)
        If vriBo Then
            DivStCard.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivStCard.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnLsBrgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBrgClose.Click
        psShowLsBrg(False)
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
            psShowLsRcvPO(False)
        End If
    End Sub

    Private Sub GrvLsRcvPO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsRcvPO.PageIndexChanging
        GrvLsRcvPO.PageIndex = e.NewPageIndex
        psFillGrvLsRcvPO()
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Protected Sub ChkStorageOID_CheckedChanged(sender As Object, e As EventArgs) Handles ChkStorageOID.CheckedChanged
        If ChkStorageOID.Checked Then
            psEnableStorageOID(True)
        Else
            psEnableStorageOID(False)
        End If
    End Sub

    Private Sub psEnableStorageOID(vriBo As Boolean)
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

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnStorageStockOID As String = vnGRow.Cells(ensColList.vStorageStockOID).Text
        Dim vnStorageOID As String = vnGRow.Cells(ensColList.vStorageOID).Text
        Dim vnStorageInfoHtml As String = vnGRow.Cells(ensColList.vStorageInfoHtml).Text
        Dim vnCompanyCode As String = vnGRow.Cells(ensColList.CompanyCode).Text
        Dim vnRcvPOHOID As String = vnGRow.Cells(ensColList.RcvPOHOID).Text
        Dim vnRcvPONo As String = vnGRow.Cells(ensColList.RcvPONo).Text
        Dim vnBrgCode As String = vnGRow.Cells(ensColList.BRGCODE).Text
        Dim vnQuery As String

        If e.CommandName = "QtyOnHand" Then
            Dim vnDtbStCard As New DataTable
            Dim vnCriteria As String
            vnCriteria = "Where 1=1"
            vnCriteria += vbCrLf & "            and sm.CompanyCode='" & vnCompanyCode & "'"
            vnCriteria += vbCrLf & "            and sm.BRGCODE='" & vnBrgCode & "'"
            vnCriteria += vbCrLf & "            and sm.RcvPOHOID=" & vnRcvPOHOID
            vnCriteria += vbCrLf & "            and sm.StorageOID=" & vnStorageOID

            LblStCardTitle.Text = "STOCK CARD " & vnCompanyCode & " " & vnStorageInfoHtml & " " & vnBrgCode & " " & vnRcvPONo

            vnQuery = "Select * From ("
            vnQuery += vbCrLf & "Select sm.OID,sm.TransCode,tm.TransName,sm.TransOID,"
            vnQuery += vbCrLf & "     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,"
            vnQuery += vbCrLf & "     sm.TransQty"
            vnQuery += vbCrLf & " From Sys_SsoStockCard_TR sm with(nolock)"
            vnQuery += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"
            vnQuery += vbCrLf & "      inner join Sys_SsoTransName_MA tm on tm.TransCode=sm.TransCode"
            vnQuery += vbCrLf & vnCriteria
            vnQuery += vbCrLf & "UNION"
            vnQuery += vbCrLf & "Select Null OID,''TransCode,''TransName,Null TransOID,"
            vnQuery += vbCrLf & "     'TOTAL'vCreationDatetime,"
            vnQuery += vbCrLf & "     sum(sm.TransQty)"
            vnQuery += vbCrLf & " From Sys_SsoStockCard_TR sm with(nolock)"
            vnQuery += vbCrLf & vnCriteria
            vnQuery += vbCrLf & ")tb"

            vnQuery += vbCrLf & "Order by case when vCreationDatetime='TOTAL' then 19 else 8 end,case when TransCode='STOB' then 4 else 5 end,OID"

            pbuFillDtbSQL(vnDtbStCard, vnQuery, vnSQLConn)
            GrvStCard.DataSource = vnDtbStCard
            GrvStCard.DataBind()

            If vnDtbStCard.Rows.Count > 0 Then
                GrvStCard.Rows(GrvStCard.Rows.Count - 1).BackColor = System.Drawing.Color.GreenYellow
            End If

            psShowStCard(True)

        ElseIf e.CommandName = "QtyOnPickList" Then
            Dim vnDtbCheck As New DataTable
            LblCheckTitle.Text = vnCompanyCode & " " & vnStorageInfoHtml & " " & vnBrgCode & " " & vnRcvPONo

            vnQuery = "Select pch.TransCode,trn.TransName,pcs.PCLHOID vTransOID,pch.PCLNo vTransNo,trs.TransStatusDescr,pcs.ReservedQty vTransQty"
            vnQuery += vbCrLf & "       From Sys_SsoPCLReserve_TR pcs with(nolock)"
            vnQuery += vbCrLf & "	         inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pcs.PCLHOID"
            vnQuery += vbCrLf & "	         inner join Sys_SsoTransName_MA trn with(nolock) on trn.TransCode=pch.TransCode"
            vnQuery += vbCrLf & "	         inner join Sys_SsoTransStatus_MA trs with(nolock) on trs.TransCode=pch.TransCode and trs.TransStatus=pch.TransStatus"
            vnQuery += vbCrLf & "	   Where pch.TransStatus=" & enuTCPICK.Prepared & " and pcs.StorageStockOID=" & vnStorageStockOID

            pbuFillDtbSQL(vnDtbCheck, vnQuery, vnSQLConn)
            GrvCheck.DataSource = vnDtbCheck
            GrvCheck.DataBind()

            psShowCheck(True)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

    End Sub

    Protected Sub BtnCheckClose_Click(sender As Object, e As EventArgs) Handles BtnCheckClose.Click
        psShowCheck(False)
    End Sub

    Private Sub BtnStCardClose_Click(sender As Object, e As EventArgs) Handles BtnStCardClose.Click
        psShowStCard(False)
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
        pbuCreateXlsx_StockInfo1(vnFileName, Session("UserOID"), DstListWarehouse, DstListCompany, DstListBuilding,
                                        DstListLantai, DstListZona, DstListStorageType, RdbListStagging,
                                        ChkStorageOID,
                                        TxtListRcvNo,
                                        TxtListRackY_SeqNo,
                                        TxtListRackY_Column,
                                        TxtListRackY_Level,
                                        TxtListRackN_Start,
                                        TxtListRackN_End,
                                        TxtListBrgCode, TxtListBrgName,
                                        vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class