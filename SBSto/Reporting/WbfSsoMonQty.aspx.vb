Imports System.Data.SqlClient
Public Class WbfSsoMonQty
    Inherits System.Web.UI.Page
    Enum ensColLsRcvPO
        RcvPONo = 0
        vRcvPODate = 1
        RcvPOSupplierName = 2
        RcvPOTypeName = 3
        OID = 4
        RcvPORefTypeOID = 5
    End Enum

    Enum ensColList
        vStorageOID = 0
        vStorageStockOID = 1
        vStorageInfoHtml = 2
        CompanyCode = 3
        RcvPONo = 4
        vRcvPODate = 5
        BRGCODE = 6
        BRGNAME = 7
        QtyOnHand = 8
        vQtyStockCard = 9
        QtyOnPutaway = 10
        vQtyOnPutaway_Trans = 11
        QtyOnPutawayWh = 12
        vQtyOnPutawayWh_Trans = 13
        QtyOnMovement = 14
        vQtyOnMovement_Trans = 15
        QtyOnMovementWh = 16
        vQtyOnMovementWh_Trans = 17
        QtyOnPickList = 18
        vQtyOnPickList_Trans = 19
        QtyOnPicking = 20
        vQtyOnPicking_Trans = 21
        QtyOnDispatch = 22
        vQtyOnDispatch_Trans = 23
        QtyOnKarantina = 24
        vQtyOnKarantina_Trans = 25
        QtyOnPutawayKr = 26
        vQtyOnPutawayKr_Trans = 27
        QtyOnPutawayDtw = 28
        vQtyOnPutawayDtw_Trans = 29
        QtyOnPutawayDty = 30
        vQtyOnPutawayDty_Trans = 31
    End Enum
    Private Sub psDefaultDisplay()
        DivLsRcvPO.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsRcvPO.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"
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
            psFillDstMonQty()

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
        Dim vnMonType As String = DstMonQty.SelectedValue
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select "
        vnQuery += vbCrLf & "     pm.vStorageOID,isnull(sm.OID,0)vStorageStockOID,pm.vStorageInfoHtml,"
        vnQuery += vbCrLf & "     mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"

        If vnMonType = "OH" Then
            vnQuery += vbCrLf & "     sm.QtyOnHand,isnull(sc.vQtyStockCard,0)vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PTW" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     sm.QtyOnPutaway,isnull(sc.vQtyTrans,0)vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PTW_WH" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayWh,isnull(sc.vQtyTrans,0)vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PDL" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnMovement,isnull(sc.vQtyTrans,0)vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PDL_WH" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnMovementWh,isnull(sc.vQtyTrans,0)vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PCL" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPickList,isnull(sc.vQtyTrans,0)vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PCK" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPicking,isnull(sc.vQtyTrans,0)vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "DSP" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnDispatch,isnull(sc.vQtyTrans,0)vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "KRT" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnKarantina,isnull(sc.vQtyTrans,0)vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "PTK" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayKr,isnull(sc.vQtyTrans,0)vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "DTW" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayDtw,isnull(sc.vQtyTrans,0)vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDty,0 vQtyOnPutawayDty_Trans"
        ElseIf vnMonType = "DTY" Then
            vnQuery += vbCrLf & "     0 QtyOnHand,0 vQtyStockCard,"
            vnQuery += vbCrLf & "     0 QtyOnPutaway,0 vQtyOnPutaway_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayWh,0 vQtyOnPutawayWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovement,0 vQtyOnMovement_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnMovementWh,0 vQtyOnMovementWh_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPickList,0 vQtyOnPickList_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPicking,0 vQtyOnPicking_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnDispatch,0 vQtyOnDispatch_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnKarantina,0 vQtyOnKarantina_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayKr,0 vQtyOnPutawayKr_Trans,"
            vnQuery += vbCrLf & "     0 QtyOnPutawayDtw,0 vQtyOnPutawayDtw_Trans,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayDty,isnull(sc.vQtyTrans,0)vQtyOnPutawayDty_Trans"
        End If

        vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
        vnQuery += vbCrLf & "      inner join fnTbl_SsoStorageStock() sm on sm.StorageOID=pm.vStorageOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"

        If vnMonType = "OH" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyStockCard_Summary_ByKey() sc"
        ElseIf vnMonType = "PTW" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPutaway_Summary_ByKey() sc"
        ElseIf vnMonType = "PTW_WH" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPutawayWh_Summary_ByKey() sc"
        ElseIf vnMonType = "PDL" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnMovement_Summary_ByKey() sc"
        ElseIf vnMonType = "PDL_WH" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnMovementWh_Summary_ByKey() sc"
        ElseIf vnMonType = "PCL" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPicklist_Summary_ByKey() sc"
        ElseIf vnMonType = "PCK" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPicking_Summary_ByKey() sc"
        ElseIf vnMonType = "DSP" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnDispatch_ByKey() sc"
        ElseIf vnMonType = "KRT" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnKarantina_ByKey() sc"
        ElseIf vnMonType = "PTK" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPtw_Karantina_ByKey() sc"
        ElseIf vnMonType = "DTW" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPutawayDtw_Summary_ByKey() sc"
        ElseIf vnMonType = "DTY" Then
            vnQuery += vbCrLf & "      left outer join fnTbl_SsoQtyOnPutawayDty_Summary_ByKey() sc"
        End If

        vnQuery += vbCrLf & "                 on sc.StorageOID=sm.StorageOID and sc.CompanyCode=sm.CompanyCode and sc.BRGCODE=sm.BRGCODE and sc.RcvPOHOID=sm.RcvPOHOID"

        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
        vnQuery += vbCrLf & "Where 1=1"

        If ChkVarianOnly.Checked Then
            If vnMonType = "OH" Then
                vnQuery += vbCrLf & "            and sm.QtyOnHand<>isnull(sc.vQtyStockCard,0)"
            ElseIf vnMonType = "PTW" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPutaway<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PTW_WH" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPutawayWh<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PDL" Then
                vnQuery += vbCrLf & "            and sm.QtyOnMovement<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PDL_WH" Then
                vnQuery += vbCrLf & "            and sm.QtyOnMovementWh<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PCL" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPickList<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PCK" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPicking<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "DSP" Then
                vnQuery += vbCrLf & "            and sm.QtyOnDispatch<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "KRT" Then
                vnQuery += vbCrLf & "            and sm.QtyOnKarantina<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "PTK" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPutawayKr<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "DTW" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPutawayDtw<>isnull(sc.vQtyTrans,0)"
            ElseIf vnMonType = "DTY" Then
                vnQuery += vbCrLf & "            and sm.QtyOnPutawayDty<>isnull(sc.vQtyTrans,0)"
            End If
        End If

        vnQuery += vbCrLf & "            and mb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "            and mb.BRGCODE like '%" & vnCrBrgCode & "%' and mb.BRGNAME like '%" & vnCrBrgName & "%'"

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
        If Trim(TxtListRcvNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and rc.RcvPONo like '%" & Trim(TxtListRcvNo.Text) & "%'"
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

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psGrvList_Default()
        psGrvList_ColsVisible(vnMonType)

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub psGrvList_Default()
        GrvList.Columns(ensColList.QtyOnDispatch).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnDispatch).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnHand).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnHand).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnMovement).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnMovement).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnMovementWh).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnMovementWh).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPicking).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPicking).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPickList).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPickList).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPutaway).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPutaway).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPutawayWh).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPutawayWh).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnKarantina).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnKarantina).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPutawayKr).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPutawayKr).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPutawayDtw).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPutawayDtw).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.QtyOnPutawayDty).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.QtyOnPutawayDty).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnDispatch_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnDispatch_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnMovementWh_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnMovementWh_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnMovement_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnMovement_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPicking_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPicking_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPickList_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPickList_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPutawayWh_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPutawayWh_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPutaway_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPutaway_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyStockCard).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyStockCard).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnKarantina_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnKarantina_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPutawayKr_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPutawayKr_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPutawayDtw_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPutawayDtw_Trans).ItemStyle.CssClass = "myDisplayNone"

        GrvList.Columns(ensColList.vQtyOnPutawayDty_Trans).HeaderStyle.CssClass = "myDisplayNone"
        GrvList.Columns(ensColList.vQtyOnPutawayDty_Trans).ItemStyle.CssClass = "myDisplayNone"
    End Sub

    Private Sub psGrvList_ColsVisible(vriMonType As String)
        If vriMonType = "OH" Then
            GrvList.Columns(ensColList.QtyOnHand).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnHand).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyStockCard).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyStockCard).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PTW" Then
            GrvList.Columns(ensColList.QtyOnPutaway).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPutaway).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPutaway_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPutaway_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PTW_WH" Then
            GrvList.Columns(ensColList.QtyOnPutawayWh).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPutawayWh).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPutawayWh_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPutawayWh_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PDL" Then
            GrvList.Columns(ensColList.QtyOnMovement).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnMovement).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnMovement_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnMovement_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PDL_WH" Then
            GrvList.Columns(ensColList.QtyOnMovementWh).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnMovementWh).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnMovementWh_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnMovementWh_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PCL" Then
            GrvList.Columns(ensColList.QtyOnPickList).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPickList).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPickList_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPickList_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PCK" Then
            GrvList.Columns(ensColList.QtyOnPicking).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPicking).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPicking_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPicking_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "DSP" Then
            GrvList.Columns(ensColList.QtyOnDispatch).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnDispatch).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnDispatch_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnDispatch_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "KRT" Then
            GrvList.Columns(ensColList.QtyOnKarantina).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnKarantina).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnKarantina_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnKarantina_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "PTK" Then
            GrvList.Columns(ensColList.QtyOnPutawayKr).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPutawayKr).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPutawayKr_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPutawayKr_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "DTW" Then
            GrvList.Columns(ensColList.QtyOnPutawayDtw).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPutawayDtw).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPutawayDtw_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPutawayDtw_Trans).ItemStyle.CssClass = ""
        ElseIf vriMonType = "DTY" Then
            GrvList.Columns(ensColList.QtyOnPutawayDty).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.QtyOnPutawayDty).ItemStyle.CssClass = ""

            GrvList.Columns(ensColList.vQtyOnPutawayDty_Trans).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vQtyOnPutawayDty_Trans).ItemStyle.CssClass = ""
        End If
    End Sub
    Private Sub psFillGrvList_20230813_Orig()
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
        vnQuery += vbCrLf & "     pm.vStorageOID,isnull(sm.OID,0)vStorageStockOID,pm.vStorageInfoHtml,"
        vnQuery += vbCrLf & "     mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
        vnQuery += vbCrLf & "     sm.QtyOnHand,"
        vnQuery += vbCrLf & "     isnull(sc.vQtyStockCard,0)vQtyStockCard"
        vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & Session("UserID") & "') pm"
        vnQuery += vbCrLf & "      inner join fnTbl_SsoStorageStock() sm on sm.StorageOID=pm.vStorageOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"
        vnQuery += vbCrLf & "      left outer join (Select StorageOID,CompanyCode,BRGCODE,RcvPOHOID,sum(TransQty)vQtyStockCard"
        vnQuery += vbCrLf & "                         From Sys_SsoStockCard_TR"
        vnQuery += vbCrLf & "                        Group by StorageOID,CompanyCode,BRGCODE,RcvPOHOID) sc on sc.StorageOID=sm.StorageOID and sc.CompanyCode=sm.CompanyCode and sc.BRGCODE=sm.BRGCODE and sc.RcvPOHOID=sm.RcvPOHOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
        vnQuery += vbCrLf & "Where 1=1"

        If ChkVarianOnly.Checked Then
            vnQuery += vbCrLf & "            and sm.QtyOnHand<>isnull(sc.vQtyStockCard,0)"
        End If

        vnQuery += vbCrLf & "            and mb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "            and mb.BRGCODE like '%" & vnCrBrgCode & "%' and mb.BRGNAME like '%" & vnCrBrgName & "%'"

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
        If Trim(TxtListRcvNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and rc.RcvPONo like '%" & Trim(TxtListRcvNo.Text) & "%'"
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

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
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
        vnQuery = "Select PM.RcvPONo,Convert(varchar(11),PM.RcvPODate)vRcvPODate,PM.RcvPOSupplierName,PT.RcvPOTypeName,PM.OID,PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & " From Sys_SsoRcvPOHeader_TR PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoRcvPOType_MA PT on PT.OID=PM.RcvPORefTypeOID"
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
    Private Sub psFillDstMonQty()
        Dim vnDtb As New DataTable
        vnDtb.Columns.Add("MonType")
        vnDtb.Columns.Add("MonTypeDescr")

        vnDtb.Rows.Add(New Object() {"OH", "Qty OnHand"})
        vnDtb.Rows.Add(New Object() {"PTW", "Qty On Putaway"})
        vnDtb.Rows.Add(New Object() {"PTW_WH", "Qty On Putaway Antar Wh"})
        vnDtb.Rows.Add(New Object() {"PDL", "Qty On Pindah Lokasi"})
        vnDtb.Rows.Add(New Object() {"PDL_WH", "Qty On Pindah Lokasi Antar Wh"})
        vnDtb.Rows.Add(New Object() {"PCL", "Qty On Picklist"})
        vnDtb.Rows.Add(New Object() {"PCK", "Qty On Picking"})
        vnDtb.Rows.Add(New Object() {"DSP", "Qty On Dispatch"})
        vnDtb.Rows.Add(New Object() {"KRT", "Qty On Karantina"})
        vnDtb.Rows.Add(New Object() {"PTK", "Qty On Putaway Karantina"})
        vnDtb.Rows.Add(New Object() {"DTW", "Qty On Putaway DO Titip"})
        vnDtb.Rows.Add(New Object() {"DTY", "Qty On Putaway DO Titip Antar WH"})

        DstMonQty.DataSource = vnDtb
        DstMonQty.DataValueField = "MonType"
        DstMonQty.DataTextField = "MonTypeDescr"
        DstMonQty.DataBind()
    End Sub
End Class