Imports System.Data.SqlClient
Public Class WbfUserGroupMs
    Inherits System.Web.UI.Page

    Private Sub psClearData()
        TxtUGName.Text = ""
        TxtUGDescr.Text = ""
        TxtOID.Text = ""
        ChkActive.Checked = False
    End Sub
    Private Sub psClearMessage()
        LblMsgUGName.Visible = False
        LblMsgUGDescr.Visible = False
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

        vnQuery = "Select * From Sys_SsoUserGroup_MA Where OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtUGName.Text = fbuValStr(vnDtb.Rows(0).Item("SsoUserGroupName"))
            TxtUGDescr.Text = fbuValStr(vnDtb.Rows(0).Item("SsoUserGroupDescr"))
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        psFillGrvAccess(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtUGName.ReadOnly = Not vriBo
        TxtUGDescr.ReadOnly = Not vriBo
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
        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Session("UserAdmin") <> 1 Then
            Response.Redirect("~/Default.aspx")
        End If

        Session("CurrentFolder") = "UserAdmin"
        If Not IsPostBack Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtUGName.Text)) = 0 Then
            LblMsgUGName.Text = "Isi Nama User Group"
            LblMsgUGName.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtUGDescr.Text)) = 0 Then
            LblMsgUGDescr.Text = "Isi Deskripsi User Group"
            LblMsgUGDescr.Visible = True
            vnSave = False
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False
        Try

            Dim vnUGName As String = fbuFormatString(Trim(TxtUGName.Text))
            Dim vnUGDescr As String = fbuFormatString(Trim(TxtUGDescr.Text))

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from Sys_SsoUserGroup_MA Where SsoUserGroupName='" & vnUGName & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUGDescr.Text = "Nama User Group " & Trim(TxtUGName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama User Group"
                    LblMsgUGDescr.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    Exit Sub
                End If

                Dim vnOID As Integer
                vnQuery = "Select isnull(max(OID),0)+1 From Sys_SsoUserGroup_MA"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoUserGroup_MA("
                vnQuery += vbCrLf & "OID,SsoUserGroupName,SsoUserGroupDescr,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values(" & vnOID & ","
                vnQuery += vbCrLf & "'" & vnUGName & "',"
                vnQuery += vbCrLf & "'" & vnUGDescr & "',"
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from Sys_SsoUserGroup_MA Where SsoUserGroupDescr='" & vnUGName & "' and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgUGDescr.Text = "Nama User Group " & Trim(TxtUGName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama User Group"
                    LblMsgUGDescr.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoUserGroup_MA set "
                vnQuery += vbCrLf & "SsoUserGroupName='" & vnUGName & "',"
                vnQuery += vbCrLf & "SsoUserGroupDescr='" & vnUGDescr & "',"
                vnQuery += vbCrLf & "Status='" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & vnUserOID
                vnQuery += vbCrLf & " Where OID=" & TxtOID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            psSaveUGAccess(vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psEnableInput(False)
            psEnableSave(False)
            BtnEdit.Enabled = True
            HdfActionStatus.Value = cbuActionNorm

        Catch ex As Exception
            LblMsgErrorNE.Text = ex.Message
            LblMsgErrorNE.Visible = True

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

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
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
        TxtOID.Text = GrvList.SelectedRow.Cells(0).Text
        psDisplayData()
        psEnableInput(False)
        psEnableSave(False)
        HdfActionStatus.Value = cbuActionNorm
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
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
        vnQuery = "Select PM.OID,PM.SsoUserGroupName,PM.SsoUserGroupDescr,"
        vnQuery += vbCrLf & "PM.Status,Convert(varchar(11),PM.CreationDatetime,106)+' '+Convert(varchar(5),PM.CreationDatetime,108)vCreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "Convert(varchar(11),PM.ModificationDatetime,106)+' '+Convert(varchar(5),PM.ModificationDatetime,108)vModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From Sys_SsoUserGroup_MA PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "      Where PM.SsoUserGroupName like '%" & vnCriteria & "%' or PM.SsoUserGroupDescr like '%" & vnCriteria & "%'"
        vnQuery += vbCrLf & " Order by PM.SsoUserGroupDescr"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvAccess(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.TransCode,PM.TransName"
        vnQuery += vbCrLf & " From Sys_SsoTransName_MA PM"
        vnQuery += vbCrLf & "Where abs(PM.IsTransMenu)=1"
        vnQuery += vbCrLf & " Order by PM.TransName"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvAccess.DataSource = vnDtb
        GrvAccess.DataBind()

        Dim vnUG As String = TxtOID.Text
        Dim vnTC As String
        Dim vnAcc As String

        Dim vnRowX As Integer
        Dim vnColX As Integer
        Dim vnCtrX As Integer
        Dim vnGRow As GridViewRow
        Dim vnChk As CheckBox

        For vnRowX = 0 To GrvAccess.Rows.Count - 1
            vnGRow = GrvAccess.Rows(vnRowX)
            vnTC = vnGRow.Cells(0).Text

            For vnColX = 2 To GrvAccess.Columns.Count - 1
                For vnCtrX = 0 To vnGRow.Cells(vnColX).Controls.Count - 1
                    If TypeOf vnGRow.Cells(vnColX).Controls(vnCtrX) Is CheckBox Then
                        vnChk = vnGRow.Cells(vnColX).Controls(vnCtrX)
                        vnAcc = vnChk.ToolTip

                        vnQuery = "Select 1 From Sys_SsoTransAccess_MA Where TransCode='" & vnTC & "' and TrAccessCode='" & vnAcc & "'"
                        If fbuGetDataNumSQL(vnQuery, vriSQLConn) = 0 Then
                            vnChk.Checked = False
                            vnChk.Visible = False
                        Else
                            vnQuery = "Select 1 From Sys_SsoUserGroupAccess_MA Where UserGroupOID=" & vnUG & " and TransCode='" & vnTC & "' and TrAccessCode='" & vnAcc & "'"
                            If fbuGetDataNumSQL(vnQuery, vriSQLConn) = 0 Then
                                vnChk.Checked = False
                            Else
                                vnChk.Checked = True
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End Sub

    Private Sub psSaveUGAccess(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        Dim vnUG As String = TxtOID.Text
        Dim vnTC As String
        Dim vnAcc As String

        Dim vnRowX As Integer
        Dim vnColX As Integer
        Dim vnCtrX As Integer
        Dim vnGRow As GridViewRow
        Dim vnChk As CheckBox

        vnQuery = "Delete Sys_SsoUserGroupAccess_MA Where UserGroupOID=" & vnUG
        pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vriSQLTrans)

        For vnRowX = 0 To GrvAccess.Rows.Count - 1
            vnGRow = GrvAccess.Rows(vnRowX)
            vnTC = vnGRow.Cells(0).Text

            For vnColX = 2 To GrvAccess.Columns.Count - 1
                For vnCtrX = 0 To vnGRow.Cells(vnColX).Controls.Count - 1
                    If TypeOf vnGRow.Cells(vnColX).Controls(vnCtrX) Is CheckBox Then
                        vnChk = vnGRow.Cells(vnColX).Controls(vnCtrX)
                        vnAcc = vnChk.ToolTip

                        If vnChk.Checked Then
                            vnQuery = "Insert into Sys_SsoUserGroupAccess_MA values(" & vnUG & ",'" & vnTC & "','" & vnAcc & "')"
                            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vriSQLTrans)
                        End If
                    End If
                Next
            Next
        Next
    End Sub
End Class