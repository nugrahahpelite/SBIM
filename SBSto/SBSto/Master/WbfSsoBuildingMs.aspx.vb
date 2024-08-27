Imports System.Data.SqlClient
Public Class WbfSsoBuildingMs
    Inherits System.Web.UI.Page

    Private Sub psClearData()
        TxtBuildCode.Text = ""
        TxtBuildName.Text = ""
        TxtBuildDescr.Text = ""
        TxtBuildPanjang.Text = ""
        TxtBuildTinggi.Text = ""
        TxtBuildLebar.Text = ""
        TxtBuildLuasArea.Text = ""
        TxtOID.Text = ""
        ChkActive.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgBuildCode.Visible = False
        LblMsgBuildName.Visible = False
        LblMsgBuildDescr.Visible = False
        LblMsgBuildPanjang.Visible = False
        LblMsgBuildTinggi.Visible = False
        LblMsgBuildLebar.Visible = False
        LblMsgBuildLuasArea.Visible = False
        LblMsgWarehouse.Visible = False
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

        vnQuery = "Select * From " & fbuGetDBMaster() & "Sys_Building_MA Where OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtBuildCode.Text = fbuValStr(vnDtb.Rows(0).Item("BuildingCode"))
            TxtBuildName.Text = fbuValStr(vnDtb.Rows(0).Item("BuildingName"))
            TxtBuildDescr.Text = fbuValStr(vnDtb.Rows(0).Item("BuildingDescription"))

            TxtBuildPanjang.Text = fbuValNum(vnDtb.Rows(0).Item("BuildingPanjang"))
            TxtBuildTinggi.Text = fbuValNum(vnDtb.Rows(0).Item("BuildingTinggi"))
            TxtBuildLebar.Text = fbuValNum(vnDtb.Rows(0).Item("BuildingLebar"))
            TxtBuildLuasArea.Text = fbuValNum(vnDtb.Rows(0).Item("BuildingLuasArea"))

            DstWarehouse.SelectedValue = fbuValNum(vnDtb.Rows(0).Item("WarehouseOID"))
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtBuildCode.ReadOnly = Not vriBo
        TxtBuildName.ReadOnly = Not vriBo
        TxtBuildDescr.ReadOnly = Not vriBo

        TxtBuildPanjang.ReadOnly = Not vriBo
        TxtBuildTinggi.ReadOnly = Not vriBo
        TxtBuildLebar.ReadOnly = Not vriBo
        TxtBuildLuasArea.ReadOnly = Not vriBo

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

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtBuildCode.Text)) = 0 Then
            LblMsgBuildCode.Text = "Isi Nama Building"
            LblMsgBuildCode.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtBuildName.Text)) = 0 Then
            LblMsgBuildName.Text = "Isi Nama Building"
            LblMsgBuildName.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtBuildDescr.Text)) = 0 Then
            LblMsgBuildDescr.Text = "Isi Deskripsi Building"
            LblMsgBuildDescr.Visible = True
            vnSave = False
        End If
        If DstWarehouse.SelectedValue = 0 Then
            LblMsgWarehouse.Text = "Pilih Warehouse"
            LblMsgWarehouse.Visible = True
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

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            Dim vnBdCode As String = fbuFormatString(Trim(TxtBuildCode.Text))
            Dim vnBdName As String = fbuFormatString(Trim(TxtBuildName.Text))
            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & fbuGetDBMaster() & "Sys_Building_MA Where BuildingName='" & vnBdName & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgBuildName.Text = "Nama Building " & Trim(TxtBuildName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Building"
                    LblMsgBuildName.Visible = True
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

                vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_Building_MA("
                vnQuery += vbCrLf & "WarehouseOID,"
                vnQuery += vbCrLf & "BuildingCode,BuildingName,"
                vnQuery += vbCrLf & "BuildingDescription,"
                vnQuery += vbCrLf & "BuildingPanjang,BuildingTinggi,"
                vnQuery += vbCrLf & "BuildingLebar,BuildingLuasArea,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values("
                vnQuery += vbCrLf & DstWarehouse.SelectedValue & ","
                vnQuery += vbCrLf & "'" & vnBdCode & "',"
                vnQuery += vbCrLf & "'" & vnBdName & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtBuildDescr.Text)) & "',"

                vnQuery += vbCrLf & Val(Trim(TxtBuildPanjang.Text)) & ","
                vnQuery += vbCrLf & Val(Trim(TxtBuildTinggi.Text)) & ","
                vnQuery += vbCrLf & Val(Trim(TxtBuildLebar.Text)) & ","
                vnQuery += vbCrLf & Val(Trim(TxtBuildLuasArea.Text)) & ","

                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnQuery = "Select OID From " & fbuGetDBMaster() & "Sys_Building_MA Where BuildingCode='" & vnBdCode & "' and BuildingName='" & vnBdName & "'"
                vnOID = fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from " & fbuGetDBMaster() & "Sys_Building_MA Where BuildingName='" & vnBdName & "' and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgBuildName.Text = "Nama Building " & Trim(TxtBuildName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Building"
                    LblMsgBuildName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & fbuGetDBMaster() & "Sys_Building_MA set"
                vnQuery += vbCrLf & "WarehouseOID=" & DstWarehouse.SelectedValue & ","
                vnQuery += vbCrLf & "BuildingCode='" & vnBdCode & "',"
                vnQuery += vbCrLf & "BuildingName='" & vnBdName & "',"
                vnQuery += vbCrLf & "BuildingDescription='" & fbuFormatString(Trim(TxtBuildDescr.Text)) & "',"

                vnQuery += vbCrLf & "BuildingPanjang=" & Val(Trim(TxtBuildPanjang.Text)) & ","
                vnQuery += vbCrLf & "BuildingTinggi=" & Val(Trim(TxtBuildTinggi.Text)) & ","
                vnQuery += vbCrLf & "BuildingLebar=" & Val(Trim(TxtBuildLebar.Text)) & ","
                vnQuery += vbCrLf & "BuildingLuasArea=" & Val(Trim(TxtBuildLuasArea.Text)) & ","

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
            vnValue = DirectCast(vnRow.Cells(1).Controls(0), LinkButton).Text
        End If
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged
        TxtOID.Text = GrvList.SelectedRow.Cells(8).Text
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

        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.BuildingCode,PM.BuildingName,PM.BuildingDescription,GM.WarehouseName,"
        vnQuery += vbCrLf & "PM.BuildingPanjang,PM.BuildingTinggi,PM.BuildingLebar,PM.BuildingLuasArea,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_Building_MA PM"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA GM on GM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "Where 1=1"
        If vnCriteria <> "" Then
            vnQuery += vbCrLf & "      and (PM.BuildingCode like '%" & vnCriteria & "%' or PM.BuildingName like '%" & vnCriteria & "%')"
        End If
        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "      and (PM.WarehouseOID=" & DstListWarehouse.SelectedValue & ")"
        End If
        vnQuery += vbCrLf & " Order by PM.BuildingName"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

End Class