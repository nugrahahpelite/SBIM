Imports GlobalUtil
Imports System.Data.SqlClient
Public Class WbfUserMs
    Inherits System.Web.UI.Page

    Enum ensColListEmp
        EmpNip = 0
        EmpName = 1
        Company = 2
        Branch = 3
        Division = 4
        JobTitle = 5
        Supervisor = 6
    End Enum

    Enum ensColCompany
        CompanyCode = 2
    End Enum
    Enum ensColWhs
        WarehouseOID = 3
    End Enum

    Private Sub psClearData()
        TxtUserID.Text = ""
        TxtUserNip.Text = ""
        TxtUserName.Text = ""
        TxtUserSSO.Text = ""
        TxtUserOID.Text = ""
        TxtUserPwd.Text = ""
        TxtUserPwdR.Text = ""
        ChkActive.Checked = False
        ChkUserAdmin.Checked = False
        ChkAllCompany.Checked = False
        ChkAllWarehouse.Checked = False

        For vn = 0 To RdlUserGroup.Items.Count - 1
            RdlUserGroup.Items(vn).Selected = False
        Next

        BtnPwdEdit.Enabled = False
        BtnEdit.Enabled = False
    End Sub

    Private Sub psClearMessage()
        LblMsgErrorNE.Visible = False
        LblMsgUserGroup.Visible = False
        LblMsgUserID.Visible = False
        LblMsgUserName.Visible = False
        LblMsgUserSSO.Visible = False
        LblMsgUserPwdR.Visible = False
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
        If TxtUserOID.Text = "" Then Exit Sub

        vnQuery = "Select * From Sys_SsoUser_MA Where OID=" & TxtUserOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
            psFillGrvCompany(False, 0, vnSQLConn)
            psFillGrvWhs(False, 0, vnSQLConn)
        Else
            TxtUserOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtUserID.Text = fbuValStr(vnDtb.Rows(0).Item("UserID"))
            TxtUserNip.Text = fbuValStr(vnDtb.Rows(0).Item("UserNip"))
            TxtUserName.Text = fbuValStr(vnDtb.Rows(0).Item("UserName"))
            TxtUserSSO.Text = fbuValStr(vnDtb.Rows(0).Item("UserSSO"))
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
            ChkUserAdmin.Checked = IIf(vnDtb.Rows(0).Item("UserAdmin") = "1", True, False)

            RdlUserGroup.SelectedValue = vnDtb.Rows(0).Item("UserGroupOID")
            'DstLocation.SelectedValue = fbuValStr(vnDtb.Rows(0).Item("UserLocationOID"))
            ChkAllWarehouse.Checked = (fbuValStr(vnDtb.Rows(0).Item("UserWarehouseCode")) = "")
            ChkAllCompany.Checked = (vnDtb.Rows(0).Item("UserCompanyCode") = "")

            BtnEdit.Enabled = True
            BtnPwdEdit.Enabled = BtnEdit.Enabled
            vnDtb.Dispose()

            psFillGrvCompany(False, TxtUserOID.Text, vnSQLConn)
            psFillGrvWhs(False, TxtUserOID.Text, vnSQLConn)
        End If
        vnDtb.Dispose()
        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtUserID.ReadOnly = Not vriBo
        TxtUserName.ReadOnly = Not vriBo
        If HdfActionStatus.Value = cbuActionNew Then
            TxtUserPwd.ReadOnly = Not vriBo
            TxtUserPwdR.ReadOnly = Not vriBo
        Else
            TxtUserPwd.ReadOnly = True
            TxtUserPwdR.ReadOnly = True
        End If
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo
        BtnDaftar.Enabled = Not vriBo
        BtnPwdEdit.Enabled = Not vriBo

        BtnEditCompany.Enabled = Not vriBo
        BtnEditWhs.Enabled = Not vriBo
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session("UserID") Is Nothing Then
            Response.Redirect("~/Default.aspx", False)
            Exit Sub
        End If
        If Session("UserAdmin") <> 1 Then
            Response.Redirect("~/Default.aspx")
        End If
        Session("CurrentFolder") = "UserAdmin"

        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            TxtUserOID.Text = Request.QueryString("vpUserOID")

            pbuFillRdlUserGroup(RdlUserGroup, vnSQLConn)
            'pbuFillDstLocation(DstLocation, True, vnSQLConn)
            psDisplayData()

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnDaftar_Click(sender As Object, e As EventArgs) Handles BtnDaftar.Click
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        Response.Redirect("~/UserAdmin/WbfUserList.aspx?vpShowDialog=0")
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtUserID.Text)) < 3 Then
            LblMsgUserID.Text = "Isikan ID User (min 4 digit)"
            LblMsgUserID.Visible = True
            TxtUserID.Focus()
            vnSave = False
        End If
        If Len(Trim(TxtUserName.Text)) < 3 Then
            LblMsgUserName.Text = "Isikan Nama User (min 3 digit)"
            LblMsgUserName.Visible = True
            TxtUserName.Focus()
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
            Dim vnUG As Boolean = False
            Dim vn As Byte
            For vn = 0 To RdlUserGroup.Items.Count - 1
                If RdlUserGroup.Items(vn).Selected Then
                    vnUG = True
                    Exit For
                End If
            Next
            If vnUG = False Then
                LblMsgUserGroup.Text = "Pilih User Group"
                LblMsgUserGroup.Visible = True
                vnSave = False
            End If

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            If HdfActionStatus.Value = cbuActionNew Then
                If Trim(TxtUserPwd.Text) = "" Or Trim(TxtUserPwd.Text) <> Trim(TxtUserPwdR.Text) Then
                    LblMsgUserPwdR.Text = "Isikan Password = Retype Password"
                    LblMsgUserPwdR.Visible = True
                    TxtUserPwd.Focus()
                    vnSave = False
                End If
                vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserID='" & Trim(TxtUserID.Text) & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUserID.Text = "User ID " & Trim(TxtUserID.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                    LblMsgUserID.Visible = True
                    TxtUserID.Focus()
                    vnSave = False
                End If
                vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserName='" & Trim(TxtUserName.Text) & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUserName.Text = "Nama User " & Trim(TxtUserName.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                    LblMsgUserName.Visible = True
                    TxtUserName.Focus()
                    vnSave = False
                End If
                If Trim(TxtUserSSO.Text) <> "" Then
                    vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserSSO='" & Trim(TxtUserSSO.Text) & "'"
                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                        LblMsgUserSSO.Text = "User SSO " & Trim(TxtUserSSO.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                        LblMsgUserSSO.Visible = True
                        TxtUserSSO.Focus()
                        vnSave = False
                    End If
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoUser_MA"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Insert into Sys_SsoUser_MA("
                vnQuery += vbCrLf & "OID,UserID,UserNip,UserName,UserSSO,"
                'vnQuery += vbCrLf & "UserLocationOID,"
                vnQuery += vbCrLf & "UserCompanyCode,"
                vnQuery += vbCrLf & "UserWarehouseCode,"
                vnQuery += vbCrLf & "UserPassword,UserAdmin,UserGroupOID,Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += vbCrLf & "values("
                vnQuery += vbCrLf & vnOID & ",'" & Trim(TxtUserID.Text) & "','" & Trim(TxtUserNip.Text) & "','" & fbuFormatString(Trim(TxtUserName.Text)) & "','" & fbuFormatString(Trim(TxtUserSSO.Text)) & "',"
                'vnQuery += vbCrLf & "'" & DstLocation.SelectedValue & "',"

                vnQuery += vbCrLf & IIf(ChkAllCompany.Checked, "''", "'0'") & ","
                vnQuery += vbCrLf & IIf(ChkAllWarehouse.Checked, "''", "'0'") & ","

                vnQuery += vbCrLf & "'" & GlobalUtil.EncryptDecrypt.Encrypt(Trim(TxtUserPwd.Text), "MyEncryptPassword") & "',"
                vnQuery += vbCrLf & IIf(ChkUserAdmin.Checked, 1, 0) & ","
                vnQuery += vbCrLf & RdlUserGroup.SelectedValue & ","
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                TxtUserOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserID='" & Trim(TxtUserID.Text) & "' and OID<>" & TxtUserOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUserID.Text = "User ID " & Trim(TxtUserID.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                    LblMsgUserID.Visible = True
                    TxtUserID.Focus()
                    vnSave = False
                End If
                vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserName='" & Trim(TxtUserName.Text) & "' and OID<>" & TxtUserOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUserName.Text = "Nama User " & Trim(TxtUserName.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                    LblMsgUserName.Visible = True
                    TxtUserName.Focus()
                    vnSave = False
                End If
                If Trim(TxtUserSSO.Text) <> "" Then
                    vnQuery = "Select count(*) from Sys_SsoUser_MA Where UserSSO='" & Trim(TxtUserSSO.Text) & "' and OID<>" & TxtUserOID.Text
                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                        LblMsgUserSSO.Text = "User SSO " & Trim(TxtUserSSO.Text) & " Sudah digunakan untuk User lain" & vbCrLf & "Silakan cek daftar User"
                        LblMsgUserSSO.Visible = True
                        TxtUserSSO.Focus()
                        vnSave = False
                    End If
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                Dim vnUserCompanyCode As String = ""
                If ChkAllCompany.Checked = False Then
                    vnQuery = "Select CompanyCode From Sys_SsoUserCompany_MA Where UserOID=" & TxtUserOID.Text
                    vnUserCompanyCode = fbuGetDataStrSQL(vnQuery, vnSQLConn)
                    If vnUserCompanyCode = "" Then
                        vnUserCompanyCode = "0"
                    End If
                End If

                Dim vnUserWarehouseCode As String = ""
                If ChkAllWarehouse.Checked = False Then
                    vnQuery = "Select WarehouseOID From Sys_SsoUserWarehouse_MA Where UserOID=" & TxtUserOID.Text
                    vnUserWarehouseCode = fbuGetDataStrSQL(vnQuery, vnSQLConn)
                    If vnUserWarehouseCode = "" Then
                        vnUserWarehouseCode = "0"
                    End If
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update Sys_SsoUser_MA set "
                vnQuery += vbCrLf & "UserID='" & Trim(TxtUserID.Text) & "',"
                vnQuery += vbCrLf & "UserNip='" & Trim(TxtUserNip.Text) & "',"
                vnQuery += vbCrLf & "UserName='" & fbuFormatString(Trim(TxtUserName.Text)) & "',"
                vnQuery += vbCrLf & "UserSSO='" & fbuFormatString(Trim(TxtUserSSO.Text)) & "',"
                vnQuery += vbCrLf & "UserCompanyCode='" & vnUserCompanyCode & "',"
                vnQuery += vbCrLf & "UserWarehouseCode='" & vnUserWarehouseCode & "',"
                vnQuery += vbCrLf & "UserAdmin=" & IIf(ChkUserAdmin.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "UserGroupOID=" & RdlUserGroup.SelectedValue & ","
                vnQuery += vbCrLf & "Status='" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & vnUserOID
                vnQuery += vbCrLf & " Where OID=" & TxtUserOID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psDisplayData()
            psEnableInput(False)
            psEnableSave(False)
            HdfActionStatus.Value = cbuActionNorm

        Catch ex As Exception
            LblMsgErrorNE.Text = "ERROR :" & ex.Message
            LblMsgErrorNE.Visible = True
        End Try
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvCompany(False, 0, vnSQLConn)
        psFillGrvWhs(False, 0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        HdfActionStatus.Value = cbuActionNew
        psClearData()
        psEnableInput(True)
        psEnableSave(True)
        ChkActive.Checked = True

        TxtUserID.Focus()
    End Sub

    Protected Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        HdfActionStatus.Value = cbuActionNorm
        psEnableInput(False)
        psEnableSave(False)
        BtnPwdEdit.Enabled = Not (Trim(TxtUserOID.Text) = "")
        If TxtUserOID.Text = "" Then
            psClearData()
        Else
            psDisplayData()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        If Trim(TxtUserOID.Text) = "" Then Exit Sub
        HdfActionStatus.Value = cbuActionEdit
        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Protected Sub BtnPwdEdit_Click(sender As Object, e As EventArgs) Handles BtnPwdEdit.Click
        TxtUserPwd.ReadOnly = False
        TxtUserPwdR.ReadOnly = False

        BtnPwdEdit.Visible = False
        BtnPwdSimpan.Visible = True
        BtnPwdBatal.Visible = True

        BtnDaftar.Enabled = False
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False
        TxtUserPwd.Focus()
    End Sub

    Protected Sub BtnPwdSimpan_Click(sender As Object, e As EventArgs) Handles BtnPwdSimpan.Click
        TxtUserPwd.Text = Trim(TxtUserPwd.Text)
        TxtUserPwdR.Text = Trim(TxtUserPwdR.Text)

        Dim vnSave As Boolean = True
        psClearMessage()
        If Len(TxtUserPwd.Text) < 4 Then
            LblMsgUserPwdR.Text = "Isikan Password = Min 4 Digit"
            LblMsgUserPwdR.Visible = True
            vnSave = False
        End If
        If TxtUserPwd.Text = "" Or TxtUserPwd.Text <> TxtUserPwdR.Text Then
            LblMsgUserPwdR.Text = "Isikan Password = Retype Password"
            LblMsgUserPwdR.Visible = True
            vnSave = False
        End If
        If Not vnSave Then
            TxtUserPwd.Focus()
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction

        Try
            Dim vnQuery As String
            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnQuery = "Update Sys_SsoUser_MA set "
            vnQuery += vbCrLf & "UserPassword='" & GlobalUtil.EncryptDecrypt.Encrypt(Trim(TxtUserPwd.Text), "MyEncryptPassword") & "',"
            vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & Session("UserOID")
            vnQuery += vbCrLf & " Where OID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Insert into Sys_SsoUserUserPwd_HS(UserOID,ChangePwdDatetime,ChangePwdUserOID) values(" & TxtUserOID.Text & ",getdate()," & TxtUserOID.Text & ")"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            TxtUserPwd.ReadOnly = True
            TxtUserPwdR.ReadOnly = True

            BtnPwdEdit.Visible = True
            BtnPwdSimpan.Visible = False
            BtnPwdBatal.Visible = False

            BtnDaftar.Enabled = True
            BtnBaru.Enabled = True
            BtnEdit.Enabled = True
        Catch ex As Exception
            vnSQLTrans = Nothing
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub BtnPwdBatal_Click(sender As Object, e As EventArgs) Handles BtnPwdBatal.Click
        TxtUserPwd.ReadOnly = True
        TxtUserPwdR.ReadOnly = True

        BtnPwdEdit.Visible = True
        BtnPwdSimpan.Visible = False
        BtnPwdBatal.Visible = False

        BtnDaftar.Enabled = True
        BtnBaru.Enabled = True
        BtnEdit.Enabled = True
    End Sub

    Protected Sub TxtUserPwd_TextChanged(sender As Object, e As EventArgs) Handles TxtUserPwd.TextChanged

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Sub BtnUserNip_Click(sender As Object, e As EventArgs) Handles BtnUserNip.Click
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        If BtnBaru.Visible Then Exit Sub
        psShowListEmp(True)
    End Sub

    Protected Sub BtnListEmpFind_Click(sender As Object, e As EventArgs) Handles BtnListEmpFind.Click
        Dim vnSQLConnHris As New SqlConnection
        If Not fbuConnectSQLHris(vnSQLConnHris) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListEmp(vnSQLConnHris)

        vnSQLConnHris.Close()
        vnSQLConnHris.Dispose()
        vnSQLConnHris = Nothing
    End Sub

    Private Sub psFillGrvListEmp(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select eh.employee_nip EmployeeNip,"
        vnQuery += vbCrLf & "       eh.Employee_First_Name + case when isnull(eh.Employee_Middle_Name,'')='' then '' else ' ' end + eh.Employee_Middle_Name + case when isnull(eh.Employee_Last_Name,'')='' then '' else ' ' end + eh.Employee_Last_Name vEmployeeName,"
        vnQuery += vbCrLf & "       es.Employee_First_Name + case when isnull(es.Employee_Middle_Name,'')='' then '' else ' ' end + es.Employee_Middle_Name + case when isnull(es.Employee_Last_Name,'')='' then '' else ' ' end + es.Employee_Last_Name vSupervisorName,"
        vnQuery += vbCrLf & "       cm.company_code,dv.div_id,bm.branch_id,jt.job_title"
        vnQuery += vbCrLf & "  From tbl_employee_h eh"
        vnQuery += vbCrLf & "       inner join tbl_employee_employment em on em.employee_nip=eh.employee_nip"
        vnQuery += vbCrLf & "       left outer join tbl_employee_h es on es.employee_nip=em.employee_supervisor_id"
        vnQuery += vbCrLf & "       inner join tbl_master_company cm on cm.uniq_no=em.employee_company_id"
        vnQuery += vbCrLf & "       inner join tbl_master_branch bm on bm.uniq_no=em.employee_branch_id"
        vnQuery += vbCrLf & "       inner join tbl_master_division dv on dv.uniq_no=em.employee_div_id"
        vnQuery += vbCrLf & "       inner join tbl_master_job_title jt on jt.uniq_no=em.employee_job_title_id"

        vnQuery += vbCrLf & "       left outer join Sys_AbsenceAttBranch ab on ab.AttBranch=em.employee_att_branch"
        vnQuery += vbCrLf & "Where 1=1"
        If Trim(TxtListEmpName.Text) <> "" Then
            vnQuery += vbCrLf & "and ("
            vnQuery += vbCrLf & "     eh.employee_nip like '%" & Trim(TxtListEmpName.Text) & "%'"
            vnQuery += vbCrLf & "     or "
            vnQuery += vbCrLf & "     eh.Employee_First_Name + case when isnull(eh.Employee_Middle_Name,'')='' then '' else ' ' end + eh.Employee_Middle_Name + case when isnull(eh.Employee_Last_Name,'')='' then '' else ' ' end + eh.Employee_Last_Name like '%" & Trim(TxtListEmpName.Text) & "%'"
            vnQuery += vbCrLf & "    )"
        End If

        vnQuery += vbCrLf & "Order by eh.Employee_First_Name,eh.Employee_Middle_Name,eh.Employee_Last_Name"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvListEmp.DataSource = vnDtb
        GrvListEmp.DataBind()

        TxtListEmpName.Focus()
    End Sub

    Protected Sub GrvListEmp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvListEmp.SelectedIndexChanged

    End Sub

    Private Sub GrvListEmp_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListEmp.RowCommand
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvListEmp.Rows(vnIdx)
        TxtUserNip.Text = DirectCast(vnGRow.Cells(ensColListEmp.EmpNip).Controls(0), LinkButton).Text
        TxtUserName.Text = vnGRow.Cells(ensColListEmp.EmpName).Text

        If Trim(TxtUserID.Text) = "" Then
            TxtUserID.Text = TxtUserNip.Text
        End If

        psShowListEmp(False)
    End Sub

    Private Sub psShowListEmp(vriBo As Boolean)
        If vriBo Then
            DivListEmp.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivListEmp.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psDefaultDisplay()
        DivListEmp.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListEmp.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub

    Protected Sub BtnEditCompany_Click(sender As Object, e As EventArgs) Handles BtnEditCompany.Click
        If Val(TxtUserOID.Text) = 0 Then Exit Sub
        If BtnSimpan.Visible Then Exit Sub
        If Not BtnEditCompany.Visible Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        psFillGrvCompany(True, TxtUserOID.Text, vnSQLConn)
        vnSQLConn.Close()
        vnSQLConn = Nothing
        psEnableSaveCompany(True)
    End Sub

    Protected Sub BtnBatalCompany_Click(sender As Object, e As EventArgs) Handles BtnBatalCompany.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        psFillGrvCompany(False, TxtUserOID.Text, vnSQLConn)
        vnSQLConn.Close()
        vnSQLConn = Nothing
        psEnableSaveCompany(False)
    End Sub

    Protected Sub BtnSimpanCompany_Click(sender As Object, e As EventArgs) Handles BtnSimpanCompany.Click
        If GrvCompany.Rows.Count > 0 Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnQuery As String

            Dim vn As Integer
            Dim vnChkCompany As CheckBox
            Dim vnGRow As GridViewRow

            Dim vnSeq As Integer
            vnQuery = "Select isnull(max(HistorySeq),0) From Sys_SsoUserCompany_HS Where UserOID=" & TxtUserOID.Text
            vnSeq = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            Dim vnSQLTrans As SqlTransaction

            vnSQLTrans = vnSQLConn.BeginTransaction("Company")

            vnQuery = "Insert into Sys_SsoUserCompany_HS (UserOID,CompanyCode,HistorySeq,HistoryDatetime,HistoryUserOID)"
            vnQuery += vbCrLf & "Select UserOID,CompanyCode," & vnSeq & ",getdate()," & Session("UserOID") & " From Sys_SsoUserCompany_MA Where UserOID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, 1, vnSQLConn, vnSQLTrans)

            vnQuery = "Delete Sys_SsoUserCompany_MA Where UserOID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, 3, vnSQLConn, vnSQLTrans)

            For vn = 0 To GrvCompany.Rows.Count - 1
                vnGRow = GrvCompany.Rows(vn)
                vnChkCompany = vnGRow.FindControl("ChkCompany")

                If vnChkCompany.Checked = True Then
                    vnQuery = "Insert into Sys_SsoUserCompany_MA(UserOID,CompanyCode)"
                    vnQuery += vbCrLf & "values(" & TxtUserOID.Text & ",'" & vnGRow.Cells(ensColCompany.CompanyCode).Text & "')"
                    pbuExecuteSQLTrans(vnQuery, 1, vnSQLConn, vnSQLTrans)
                End If
            Next

            Dim vnUserCompanyCode As String = ""
            If ChkAllCompany.Checked = False Then
                vnQuery = "Select CompanyCode From Sys_SsoUserCompany_MA Where UserOID=" & TxtUserOID.Text
                vnUserCompanyCode = fbuGetDataStrSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)
                If vnUserCompanyCode = "" Then
                    vnUserCompanyCode = "0"
                End If
            End If

            vnQuery = "Update Sys_SsoUser_MA set "
            vnQuery += vbCrLf & "UserCompanyCode='" & vnUserCompanyCode & "'"
            vnQuery += vbCrLf & " Where OID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()

            psFillGrvCompany(False, TxtUserOID.Text, vnSQLConn)
            vnSQLConn.Close()
            vnSQLConn = Nothing
        End If
        psEnableSaveCompany(False)
    End Sub

    Private Sub psEnableSaveCompany(vriBo As Boolean)
        BtnEditCompany.Visible = Not vriBo
        BtnSimpanCompany.Visible = vriBo
        BtnBatalCompany.Visible = vriBo

        BtnBaru.Enabled = Not vriBo
        BtnEdit.Enabled = Not vriBo
    End Sub

    Private Sub psEnableSaveWhs(vriBo As Boolean)
        BtnEditWhs.Visible = Not vriBo
        BtnSimpanWhs.Visible = vriBo
        BtnBatalWhs.Visible = vriBo

        BtnBaru.Enabled = Not vriBo
        BtnEdit.Enabled = Not vriBo
    End Sub

    Private Sub psFillGrvCompany(vriEdit As Boolean, vriUserOID As Integer, vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnDtbB As New DataTable

        Dim vnQuery As String = ""
        If vriEdit Then
            vnQuery = "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany"
            vnQuery += vbCrLf & "Order by CompanyName"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        Else
            vnQuery = "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany"
            vnQuery += vbCrLf & "Where CompanyCode in (Select b.CompanyCode From Sys_SsoUserCompany_MA b Where b.UserOID=" & vriUserOID & ")"
            vnQuery += vbCrLf & "Order by CompanyName"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        End If

        GrvCompany.DataSource = vnDtb
        GrvCompany.DataBind()

        Dim vn As Integer
        Dim vnChkCompany As CheckBox
        Dim vnRow As GridViewRow
        If vriEdit Then
            If GrvCompany.Rows.Count > 0 Then
                vnQuery = "Select CompanyCode From Sys_SsoUserCompany_MA Where UserOID=" & vriUserOID
                pbuFillDtbSQL(vnDtbB, vnQuery, vriSQLConn)
                If vnDtbB.Rows.Count > 0 Then
                    Dim vnB As Integer
                    For vnB = 0 To vnDtbB.Rows.Count - 1
                        For vn = 0 To GrvCompany.Rows.Count - 1
                            vnRow = GrvCompany.Rows(vn)
                            vnChkCompany = vnRow.FindControl("ChkCompany")
                            If vnDtbB.Rows(vnB).Item(0) = vnRow.Cells(2).Text Then
                                vnChkCompany.Checked = True
                            End If
                        Next
                    Next
                End If
            End If
        Else
            For vn = 0 To GrvCompany.Rows.Count - 1
                vnRow = GrvCompany.Rows(vn)
                vnChkCompany = vnRow.FindControl("ChkCompany")
                vnChkCompany.Checked = True
            Next
        End If
    End Sub

    Private Sub psFillGrvWhs(vriEdit As Boolean, vriUserOID As Integer, vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnDtbB As New DataTable

        Dim vnQuery As String = ""
        If vriEdit Then
            vnQuery = "Select OID,WarehouseCode,WarehouseName From " & fbuGetDBMaster() & "Sys_Warehouse_MA"
            vnQuery += vbCrLf & "Order by WarehouseName"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        Else
            vnQuery = "Select OID,WarehouseCode,WarehouseName From " & fbuGetDBMaster() & "Sys_Warehouse_MA"
            vnQuery += vbCrLf & "Where OID in (Select b.WarehouseOID From Sys_SsoUserWarehouse_MA b Where b.UserOID=" & vriUserOID & ")"
            vnQuery += vbCrLf & "Order by WarehouseName"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        End If

        GrvWhs.DataSource = vnDtb
        GrvWhs.DataBind()

        Dim vn As Integer
        Dim vnChkWhs As CheckBox
        Dim vnRow As GridViewRow
        If vriEdit Then
            If GrvWhs.Rows.Count > 0 Then
                vnQuery = "Select WarehouseOID From Sys_SsoUserWarehouse_MA Where UserOID=" & vriUserOID
                pbuFillDtbSQL(vnDtbB, vnQuery, vriSQLConn)
                If vnDtbB.Rows.Count > 0 Then
                    Dim vnB As Integer
                    For vnB = 0 To vnDtbB.Rows.Count - 1
                        For vn = 0 To GrvWhs.Rows.Count - 1
                            vnRow = GrvWhs.Rows(vn)
                            vnChkWhs = vnRow.FindControl("ChkWhs")
                            If vnDtbB.Rows(vnB).Item(0) = vnRow.Cells(ensColWhs.WarehouseOID).Text Then
                                vnChkWhs.Checked = True
                            End If
                        Next
                    Next
                End If
            End If
        Else
            For vn = 0 To GrvWhs.Rows.Count - 1
                vnRow = GrvWhs.Rows(vn)
                vnChkWhs = vnRow.FindControl("ChkWhs")
                vnChkWhs.Checked = True
            Next
        End If
    End Sub

    Protected Sub BtnListEmpClose_Click(sender As Object, e As EventArgs) Handles BtnListEmpClose.Click
        psShowListEmp(False)
    End Sub

    Protected Sub BtnEditWhs_Click(sender As Object, e As EventArgs) Handles BtnEditWhs.Click
        If Val(TxtUserOID.Text) = 0 Then Exit Sub
        If BtnSimpan.Visible Then Exit Sub
        If Not BtnEditWhs.Visible Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        psFillGrvWhs(True, TxtUserOID.Text, vnSQLConn)
        vnSQLConn.Close()
        vnSQLConn = Nothing
        psEnableSaveWhs(True)
    End Sub

    Protected Sub BtnBatalWhs_Click(sender As Object, e As EventArgs) Handles BtnBatalWhs.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        psFillGrvWhs(False, TxtUserOID.Text, vnSQLConn)
        vnSQLConn.Close()
        vnSQLConn = Nothing
        psEnableSaveWhs(False)
    End Sub

    Protected Sub BtnSimpanWhs_Click(sender As Object, e As EventArgs) Handles BtnSimpanWhs.Click
        If GrvWhs.Rows.Count > 0 Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnQuery As String

            Dim vn As Integer
            Dim vnChkWhs As CheckBox
            Dim vnRow As GridViewRow

            Dim vnSeq As Integer
            vnQuery = "Select isnull(max(HistorySeq),0) From Sys_SsoUserWarehouse_HS Where UserOID=" & TxtUserOID.Text
            vnSeq = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            Dim vnSQLTrans As SqlTransaction

            vnSQLTrans = vnSQLConn.BeginTransaction("Whs")

            vnQuery = "Insert into Sys_SsoUserWarehouse_HS(UserOID,WarehouseOID,HistorySeq,HistoryDatetime,HistoryUserOID)"
            vnQuery += vbCrLf & "Select UserOID,WarehouseOID," & vnSeq & ",getdate()," & Session("UserOID") & " From Sys_SsoUserWarehouse_MA Where UserOID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, 1, vnSQLConn, vnSQLTrans)

            vnQuery = "Delete Sys_SsoUserWarehouse_MA Where UserOID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, 3, vnSQLConn, vnSQLTrans)

            For vn = 0 To GrvWhs.Rows.Count - 1
                vnRow = GrvWhs.Rows(vn)
                vnChkWhs = vnRow.FindControl("ChkWhs")

                If vnChkWhs.Checked = True Then
                    vnQuery = "Insert into Sys_SsoUserWarehouse_MA(UserOID,WarehouseOID)"
                    vnQuery += vbCrLf & "values(" & TxtUserOID.Text & "," & vnRow.Cells(ensColWhs.WarehouseOID).Text & ")"
                    pbuExecuteSQLTrans(vnQuery, 1, vnSQLConn, vnSQLTrans)
                End If
            Next

            Dim vnUserWarehouseCode As String = ""
            If ChkAllWarehouse.Checked = False Then
                vnQuery = "Select WarehouseOID From Sys_SsoUserWarehouse_MA Where UserOID=" & TxtUserOID.Text
                vnUserWarehouseCode = fbuGetDataStrSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)
                If vnUserWarehouseCode = "" Then
                    vnUserWarehouseCode = "0"
                End If
            End If

            vnQuery = "Update Sys_SsoUser_MA set "
            vnQuery += vbCrLf & "UserWarehouseCode='" & vnUserWarehouseCode & "'"
            vnQuery += vbCrLf & " Where OID=" & TxtUserOID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()

            psFillGrvWhs(False, TxtUserOID.Text, vnSQLConn)
            vnSQLConn.Close()
            vnSQLConn = Nothing
        End If
        psEnableSaveWhs(False)
    End Sub
End Class