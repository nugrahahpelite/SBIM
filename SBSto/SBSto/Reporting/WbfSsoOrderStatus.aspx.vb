Imports System.Data.SqlClient
Imports System.IO
Public Class WbfSsoOrderStatus
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoOrderStatus"

    Enum ensColList
        CompanyCode = 0
        WarehouseOID = 1
        WarehouseName = 2
        BRGCODE = 3
        BRGNAME = 4
    End Enum

    Private Sub psDefaultDisplay()

    End Sub
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Reporting"
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            If Session("UserCompanyCode") = "" Then
                pbuFillCstCompany(ChkListCompany, True, vnSQLConn)
            Else
                pbuFillCstCompanyByUser(Session("UserOID"), ChkListCompany, True, vnSQLConn)
            End If

            pbuFillCstWarehouse_ByUserOID(Session("UserOID"), ChkListWarehouse, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            ChkListCompany.Items(0).Selected = True
            ChkListWarehouse.Items(0).Selected = True
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnCrWarehouse As String = "''"
        Dim vnCrWhAll As Boolean = False

        For vn = 0 To ChkListWarehouse.Items.Count - 1
            If ChkListWarehouse.Items(vn).Selected Then
                If ChkListWarehouse.Items(vn).Value = 0 Then
                    vnCrWhAll = True
                Else
                    vnCrWarehouse += ",'" & fbuGetWarehouseName(ChkListWarehouse.Items(vn).Value, vnSQLConn) & "'"
                End If
            End If
        Next
        If vnCrWhAll = True Then
            vnCrWarehouse = ""
        Else
            If vnCrWarehouse = "''" Then
                ChkListWarehouse.Items(0).Selected = True
                vnCrWarehouse = ""
            Else
                vnCrWarehouse = " and PM.WarehouseName  IN (" & vnCrWarehouse & ")"
            End If
        End If

        Dim vnCrCompany As String = "''"
        Dim vnCrComAll As Boolean = False

        For vn = 0 To ChkListCompany.Items.Count - 1
            If ChkListCompany.Items(vn).Selected Then
                If ChkListCompany.Items(vn).Value = "" Then
                    vnCrComAll = True
                Else
                    vnCrCompany += ",'" & ChkListCompany.Items(vn).Value & "'"
                End If
            End If
        Next
        If vnCrComAll = True Then
            vnCrCompany = ""
        Else
            If vnCrCompany = "''" Then
                ChkListCompany.Items(0).Selected = True
                vnCrCompany = ""
            Else
                vnCrCompany = " and PM.CompanyCode  IN (" & vnCrCompany & ")"
            End If
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	SELECT PM.[CompanyCode],PM.[WarehouseName],[Order Status] vOrderStatus,[Ref No] vRefNo	"
        vnQuery += vbCrLf & "	  ,[TANGGAL],[Priority] ,[DO Titip] vDoTitip	"
        vnQuery += vbCrLf & "	  ,[KODE_CUST],[CUSTOMER],[uploadDatetime]	"
        vnQuery += vbCrLf & "	  ,[Picklist No] vPicklistNo,[Picklist Date] vPickListDate,[PL Created by]vPLCreate,[PreparedDatetime],[PL Status]	"
        vnQuery += vbCrLf & "	  ,[Picking No],[Picking Created Date] vPickingCreate,[Picking Done] vPickingDone	"
        vnQuery += vbCrLf & "	  ,[Dispatch/Putaway No] vDispatchNo,[Confirm Dispatch/Putaway Date] vDispatchConfirm,[Driver Confirm Date] vDriverConfirm	"
        vnQuery += vbCrLf & "	  ,[Driver Name] vDriverName,[Driver Return Time] vDriverReturn	"

        vnQuery += vbCrLf & " FROM " & fbuGetDBDcm() & "vOrderTracing PM"
        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA mw with(nolock) on mw.WarehouseName=PM.WarehouseName"
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=mw.OID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        If RdlListPickType.SelectedValue = 1 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') in ('Baru','Prepared','On Picking'))"

        ElseIf RdlListPickType.SelectedValue = 2 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') != 'Cancelled' and isnull([PL Status],'') != 'Void') and ([Picking Created Date] is not null) and ([Picking Done] is not null) and ([Confirm Dispatch/Putaway Date] is NULL) and ([Driver Return Time] is null)  "
        ElseIf RdlListPickType.SelectedValue = 3 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and isnull([Picklist No],'')=''"
        Else
            vnQuery += vbCrLf & ""
        End If

        vnQuery += vbCrLf & vnCrWarehouse & vnCrCompany & " Order by PM.TANGGAL Desc"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
    Private Sub psFillGrvList_20231011_Orig()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnCompanyCode As String = "" ' DstListCompany.SelectedValue
        Dim vnWarehouse As String = "" 'DstListWarehouse.SelectedValue

        Dim vnCutOfDate As String = fbuGetCutOfDate(vnCompanyCode, vnSQLConn)
        Dim vnWarehouseName As String = fbuGetWarehouseName(vnWarehouse, vnSQLConn)
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        Dim vnCrCB As String = ""

        Dim vnCrDate As String = ""

        Dim vnCrStatus As String = ""

        vnQuery = "	SELECT PM.[CompanyCode],PM.[WarehouseName],[Order Status] vOrderStatus,[Ref No] vRefNo	"
        vnQuery += vbCrLf & "	  ,[TANGGAL],[Priority] ,[DO Titip] vDoTitip	"
        vnQuery += vbCrLf & "	  ,[KODE_CUST],[CUSTOMER],[uploadDatetime]	"
        vnQuery += vbCrLf & "	  ,[Picklist No] vPicklistNo,[Picklist Date] vPickListDate,[PL Created by]vPLCreate,[PreparedDatetime],[PL Status]	"
        vnQuery += vbCrLf & "	  ,[Picking No],[Picking Created Date] vPickingCreate,[Picking Done] vPickingDone	"
        vnQuery += vbCrLf & "	  ,[Dispatch/Putaway No] vDispatchNo,[Confirm Dispatch/Putaway Date] vDispatchConfirm,[Driver Confirm Date] vDriverConfirm	"
        vnQuery += vbCrLf & "	  ,[Driver Name] vDriverName,[Driver Return Time] vDriverReturn	"

        vnQuery += vbCrLf & " FROM " & fbuGetDBDcm() & "vOrderTracing PM"
        If vnUserCompanyCode <> "" Then
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA mw with(nolock) on mw.WarehouseName=PM.WarehouseName"
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=mw.OID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        'If DstListCompany.SelectedIndex > 0 Then
        '    vnQuery += vbCrLf & "            and PM.CompanyCode = '" & vnCompanyCode & "'"
        'End If
        'If DstListWarehouse.SelectedIndex > 0 Then
        '    vnQuery += vbCrLf & "            and PM.WarehouseName = '" & vnWarehouseName & "'"
        'End If
        If RdlListPickType.SelectedValue = 1 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') in ('Baru','Prepared','On Picking'))"

        ElseIf RdlListPickType.SelectedValue = 2 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') != 'Cancelled' and isnull([PL Status],'') != 'Void') and ([Picking Created Date] is not null) and ([Picking Done] is not null) and ([Confirm Dispatch/Putaway Date] is NULL) and ([Driver Return Time] is null)  "
        ElseIf RdlListPickType.SelectedValue = 3 Then
            vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and isnull([Picklist No],'')=''"
        Else
            vnQuery += vbCrLf & ""
        End If
        vnQuery += vbCrLf & "Order by PM.TANGGAL Desc"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        LblMsgListStart.Text = ""
        psFillGrvList()
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

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
        pbuCreateXlsx_OrderStatus2(vnFileName, Session("UserOID"), ChkListWarehouse, ChkListCompany, RdlListPickType, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub



End Class