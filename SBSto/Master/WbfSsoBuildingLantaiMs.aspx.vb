Imports System.Data.SqlClient
Public Class WbfSsoBuildingLantaiMs
    Inherits System.Web.UI.Page

    Private Sub psClearData()
        TxtOID.Text = ""
        ChkActive.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgBuilding.Visible = False
        LblMsgLantai.Visible = False
        LblMsgErrorNE.Visible = False
    End Sub

    Private Sub psDisplayData()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtOID.Text = "" Then Exit Sub

        Dim vnDBMaster As String = fbuGetDBMaster()
        vnQuery = "Select BM.WarehouseOID,PM.BuildingOID,PM.LantaiOID,PM.Status"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_BuildingLantaiRel_MA PM"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Building_MA BM on BM.OID=PM.BuildingOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Lantai_MA LM on LM.OID=PM.LantaiOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA WM on WM.OID=BM.WarehouseOID"
        vnQuery += vbCrLf & "Where PM.OID=" & TxtOID.Text

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            DstWarehouse.SelectedValue = vnDtb.Rows(0).Item("WarehouseOID")
            pbuFillDstBuilding_ByWarehouse(DstBuilding, False, DstWarehouse.SelectedValue, vnSQLConn)

            DstBuilding.SelectedValue = vnDtb.Rows(0).Item("BuildingOID")
            DstLantai.SelectedValue = vnDtb.Rows(0).Item("LantaiOID")

            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        LblMsgErrorNE.Visible = False
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo
        BtnFind.Enabled = Not vriBo
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Master"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            BtnBaru.Enabled = (Session("UserAdmin") = 1)
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsGudang, vnSQLConn)

            pbuFillDstWarehouse(DstWarehouse, False, vnSQLConn)
            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)

            pbuFillDstLantai(DstLantai, False, vnSQLConn)
            pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If DstBuilding.SelectedValue = 0 Then
            LblMsgBuilding.Text = "Pilih Building"
            LblMsgBuilding.Visible = True
            vnSave = False
        End If
        If DstLantai.SelectedValue = 0 Then
            LblMsgLantai.Text = "Pilih Lantai"
            LblMsgLantai.Visible = True
            vnSave = False
        End If
        Try
            Dim vnSQLConn As New SqlConnection
            Dim vnSQLTrans As SqlTransaction
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnDBMaster As String = fbuGetDBMaster()

            Dim vnBuildingOID As Integer = DstBuilding.SelectedValue
            Dim vnLantaiOID As Integer = DstLantai.SelectedValue

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_BuildingLantaiRel_MA Where BuildingOID=" & vnBuildingOID & " and LantaiOID=" & vnLantaiOID
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgLantai.Text = "Building " & DstBuilding.SelectedItem.Text & " - Lantai " & DstLantai.SelectedItem.Text & " Sudah Ada<br />Silakan Cek Daftar Building - Lantai"
                    LblMsgLantai.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
                Dim vnOID As Integer

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Insert into " & vnDBMaster & "Sys_BuildingLantaiRel_MA("
                vnQuery += vbCrLf & "BuildingOID,"
                vnQuery += vbCrLf & "LantaiOID,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values("
                vnQuery += vbCrLf & vnBuildingOID & ","
                vnQuery += vbCrLf & vnLantaiOID & ","
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnQuery = "Select OID From " & vnDBMaster & "Sys_BuildingLantaiRel_MA with(nolock) Where BuildingOID=" & vnBuildingOID & " and LantaiOID=" & vnLantaiOID
                vnOID = fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_BuildingLantaiRel_MA Where BuildingOID=" & vnBuildingOID & " and LantaiOID=" & vnLantaiOID & " and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgLantai.Text = "Building " & DstBuilding.SelectedItem.Text & " - Lantai " & DstLantai.SelectedItem.Text & " Sudah Ada<br />Silakan Cek Daftar Building - Lantai"
                    LblMsgLantai.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & vnDBMaster & "Sys_BuildingLantaiRel_MA set "
                vnQuery += vbCrLf & "BuildingOID=" & vnBuildingOID & ","
                vnQuery += vbCrLf & "LantaiOID=" & vnLantaiOID & ","
                vnQuery += vbCrLf & "Status='" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & vnUserOID
                vnQuery += vbCrLf & " Where OID=" & TxtOID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psEnableInput(False)
            psEnableSave(False)
            BtnEdit.Enabled = True
            HdfActionStatus.Value = cbuActionNorm

        Catch ex As Exception
            LblMsgErrorNE.Text = ex.Message
            LblMsgErrorNE.Visible = True
        End Try
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psClearData()
        psEnableInput(True)
        psEnableSave(True)
        ChkActive.Checked = True
        HdfActionStatus.Value = cbuActionNew
    End Sub

    Protected Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
        psEnableSave(False)
        HdfActionStatus.Value = cbuActionNorm
        If TxtOID.Text = "" Then
            psClearData()
        Else
            psDisplayData()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Trim(TxtOID.Text) = "" Then Exit Sub
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psEnableInput(True)
        psEnableSave(True)
        HdfActionStatus.Value = cbuActionEdit
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If e.CommandName = "Select" Then
            If BtnSimpan.Visible Then Exit Sub
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)
            'vnValue = DirectCast(vnRow.Cells(1).Controls(0), LinkButton).Text
        End If
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged
        TxtOID.Text = GrvList.SelectedRow.Cells(3).Text
        psDisplayData()
        psEnableInput(False)
        psEnableSave(False)
        HdfActionStatus.Value = cbuActionNorm
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_BuildingLantaiRel_MA PM"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Building_MA BM on BM.OID=PM.BuildingOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Lantai_MA LM on LM.OID=PM.LantaiOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA WM on WM.OID=BM.WarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "Where 1=1"
        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and BM.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If
        If Val(DstListBuilding.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and PM.BuildingOID=" & DstListBuilding.SelectedValue
        End If

        vnQuery += vbCrLf & " Order by WM.WarehouseName,BM.BuildingName,LM.LantaiDescription"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub DstWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstWarehouse.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuFillDstBuilding_ByWarehouse(DstBuilding, False, DstWarehouse.SelectedValue, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class