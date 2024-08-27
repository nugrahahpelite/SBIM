Imports System.Data.SqlClient
Public Class WbfSsoStockPick
    Inherits System.Web.UI.Page
    Enum ensColLsRcvPO
        RcvPONo = 0
        vRcvPODate = 1
        RcvPOSupplierName = 2
        RcvPOTypeName = 3
        OID = 4
        RcvPORefTypeOID = 5
    End Enum
    Private Sub psDefaultDisplay()
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

            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)

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

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblMsgListCompany.Text = ""
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
        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnCrStatus As String = ""
        If ChkSt_ReadyToDispatch.Checked = False And ChkSt_OnDispatch.Checked = False And ChkSt_DispatchDone.Checked = False Then
            ChkSt_ReadyToDispatch.Checked = True
            ChkSt_OnDispatch.Checked = True
        End If

        If ChkSt_ReadyToDispatch.Checked = True Then
            vnCrStatus += enuTCDISG.Ready_To_Dispatch & ","
        End If
        If ChkSt_OnDispatch.Checked = True Then
            vnCrStatus += enuTCDISG.On_Dispatch & ","
        End If
        If ChkSt_DispatchDone.Checked = True Then
            vnCrStatus += enuTCDISG.Dispatch_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select sti.vStorageOID,sti.WarehouseName,sti.vStorageInfoHtml,"
        vnQuery += vbCrLf & "     pm.DSRCompanyCode,pm.vStockPickHOID,"
        vnQuery += vbCrLf & "     pm.DSRNo,convert(varchar(11),pm.DSRDate,106)vDSRDate,pm.PCKNo,convert(varchar(11),pm.PCKDate,106)vPCKDate,"
        vnQuery += vbCrLf & "     pm.PCLNo,scd.SchDTypeName,pm.vPCLRefHInfoHtml,"
        vnQuery += vbCrLf & "     pm.TransStatus,pm.TransStatusDescr,"
        vnQuery += vbCrLf & "     pm.vReceiveInfoHtml,pm.vDispatchInfoHtml"
        vnQuery += vbCrLf & " From fnTbl_SsoStockPick('" & Session("UserID") & "')pm"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sti on sti.vStorageOID=pm.StorageOID"
        vnQuery += vbCrLf & "      inner join " & vnDBDcm & "Sys_DcmSchDType_MA scd with(nolock) on scd.OID=pm.SchDTypeOID"
        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & "            and pm.DSRCompanyCode='" & DstListCompany.SelectedValue & "'"

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and pm.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If
        vnQuery += vbCrLf & vnCrStatus
        vnQuery += vbCrLf & " Order by sti.WarehouseName,pm.DSRNo"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

End Class