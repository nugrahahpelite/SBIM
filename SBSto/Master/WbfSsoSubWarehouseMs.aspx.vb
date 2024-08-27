Imports System.Data.SqlClient
Public Class WbfSsoSubWarehouseMs
    Inherits System.Web.UI.Page


    Private Sub psClearData()
        TxtSubWhsCode.Text = ""
        TxtWhsSubName.Text = ""
        TxtWhsSubDescr.Text = ""
        TxtOID.Text = ""
        ChkActive.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgSubWhsCode.Visible = False
        LblMsgSubWhsName.Visible = False
        LblMsgSubWhsDescr.Visible = False
        LblMsgWhs.Visible = False
        LblMsgCompany.Visible = False
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

        vnQuery = "Select * From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA Where OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = vnDtb.Rows(0).Item("OID")
            TxtSubWhsCode.Text = vnDtb.Rows(0).Item("SubWhsCode")
            TxtWhsSubName.Text = vnDtb.Rows(0).Item("SubWhsName")
            TxtWhsSubDescr.Text = vnDtb.Rows(0).Item("SubWhsDescription")
            DstCompany.SelectedValue = vnDtb.Rows(0).Item("CompanyCode")
            DstWhs.SelectedValue = vnDtb.Rows(0).Item("WarehouseOID")
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSubWhsCode.ReadOnly = Not vriBo
        TxtWhsSubName.ReadOnly = Not vriBo
        TxtWhsSubDescr.ReadOnly = Not vriBo
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

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse(DstWhs, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        Try
            If Len(Trim(TxtSubWhsCode.Text)) = 0 Then
                LblMsgSubWhsCode.Text = "Isi Nama Warehouse"
                LblMsgSubWhsCode.Visible = True
                vnSave = False
            End If
            If Len(Trim(TxtWhsSubName.Text)) = 0 Then
                LblMsgSubWhsName.Text = "Isi Nama Warehouse"
                LblMsgSubWhsName.Visible = True
                vnSave = False
            End If
            If DstCompany.SelectedValue = "" Then
                LblMsgCompany.Text = "Pilih Company"
                LblMsgCompany.Visible = True
                vnSave = False
            End If
            If DstWhs.SelectedValue = 0 Then
                LblMsgWhs.Text = "Pilih Warehouse"
                LblMsgWhs.Visible = True
                vnSave = False
            End If

            Dim vnSQLConn As New SqlConnection
            Dim vnSQLTrans As SqlTransaction
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If
            Dim vnDBMaster As String = fbuGetDBMaster()

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_SubWarehouse_MA Where SubWhsName='" & fbuFormatString(Trim(TxtWhsSubName.Text)) & "' and CompanyCode='" & DstCompany.SelectedValue & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgSubWhsName.Text = "Nama Warehouse " & Trim(TxtWhsSubName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Warehouse untuk Company " & DstCompany.SelectedItem.Text
                    LblMsgSubWhsName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
                Dim vnOID As Integer
                vnQuery = "Select isnull(max(OID),0)+1 From " & vnDBMaster & "Sys_SubWarehouse_MA"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Insert into " & vnDBMaster & "Sys_SubWarehouse_MA("
                vnQuery += vbCrLf & "OID,SubWhsCode,SubWhsName,"
                vnQuery += vbCrLf & "SubWhsDescription,"
                vnQuery += vbCrLf & "CompanyCode,"
                vnQuery += vbCrLf & "WarehouseOID,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values(" & vnOID & ","
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtSubWhsCode.Text)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtWhsSubName.Text)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtWhsSubDescr.Text)) & "',"
                vnQuery += vbCrLf & "'" & DstCompany.SelectedValue & "',"
                vnQuery += vbCrLf & DstWhs.SelectedValue & ","
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_SubWarehouse_MA Where SubWhsName='" & fbuFormatString(Trim(TxtWhsSubName.Text)) & "' and CompanyCode='" & DstCompany.SelectedValue & "' and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgSubWhsName.Text = "Nama Warehouse " & Trim(TxtWhsSubName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Warehouse untuk Company " & DstCompany.SelectedItem.Text
                    LblMsgSubWhsName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & vnDBMaster & "Sys_SubWarehouse_MA set "
                vnQuery += vbCrLf & "SubWhsCode='" & fbuFormatString(Trim(TxtSubWhsCode.Text)) & "',"
                vnQuery += vbCrLf & "SubWhsName='" & fbuFormatString(Trim(TxtWhsSubName.Text)) & "',"
                vnQuery += vbCrLf & "SubWhsDescription='" & fbuFormatString(Trim(TxtWhsSubDescr.Text)) & "',"
                vnQuery += vbCrLf & "CompanyCode='" & DstCompany.SelectedValue & "',"
                vnQuery += vbCrLf & "WarehouseOID=" & DstWhs.SelectedValue & ","
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
        TxtOID.Text = GrvList.SelectedRow.Cells(5).Text
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
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.SubWhsCode,PM.SubWhsName,PM.SubWhsDescription,PM.CompanyCode,CM.WarehouseName,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_SubWarehouse_MA PM"
        vnQuery += vbCrLf & "      left outer join " & vnDBMaster & "Sys_Warehouse_MA CM on CM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & " Where (PM.SubWhsCode like '%" & vnCriteria & "%' or PM.SubWhsName like '%" & vnCriteria & "%')"

        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "       and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & " Order by PM.SubWhsName"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class