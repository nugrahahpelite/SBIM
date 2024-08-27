Imports System.Data.SqlClient

Public Class WbfUserList
    Inherits System.Web.UI.Page
    Dim vsDtb As New DataTable

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
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            pbuFillDstUserGroup(DstUserGroup, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            If Session("LstUsFind") = "Y" Then
                TxtKriteria.Text = Session("LstUsTxtKriteria")
                DstUserGroup.SelectedValue = Session("LstUsDstUserGroup")
                psFillGrvUser()
            End If
        End If
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        psFillGrvUser()
        Session("LstUsFind") = "Y"
        Session("LstUsTxtKriteria") = TxtKriteria.Text
        Session("LstUsDstUserGroup") = DstUserGroup.SelectedValue
    End Sub

    Private Sub psFillGrvUser()
        Dim vnUserOID As String = Session("UserOID")
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        vsDtb = New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        vnQuery = "Select PM.OID,PM.UserID,PM.UserSSO,PM.UserNip,"
        vnQuery += vbCrLf & "case when isnull(PM.UserCompanyCode,'')='' then 'ALL' else PM.UserCompanyCode end UserCompanyCode,"
        vnQuery += vbCrLf & "case when isnull(PM.UserWarehouseCode,'')='' then 'ALL'"
        vnQuery += vbCrLf & "     when isnull(PM.UserWarehouseCode,'')='0' then PM.UserWarehouseCode"
        vnQuery += vbCrLf & "     else WM.WarehouseName end UserWarehouseCode,"
        vnQuery += vbCrLf & "PM.UserName,"
        vnQuery += vbCrLf & "case when abs(PM.UserAdmin)=1 then 'Y' else 'N' end Admin,"
        vnQuery += vbCrLf & "RM.SsoUserGroupName,PM.status,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,SM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From Sys_SsoUser_MA PM"
        vnQuery += vbCrLf & "      left outer join Sys_SsoLocation_MA LM on LM.OID=PM.UserLocationOID"
        vnQuery += vbCrLf & "      left outer join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM on WM.OID=PM.UserWarehouseCode"
        vnQuery += vbCrLf & "      inner join Sys_SsoUserGroup_MA RM on RM.OID=PM.UserGroupOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.oid=PM.CreationUserOID"
        vnQuery += vbCrLf & "Where (PM.UserID like '%" & vnCriteria & "%' or PM.UserNip like '%" & vnCriteria & "%' or PM.UserName like '%" & vnCriteria & "%' or PM.UserSSO like '%" & vnCriteria & "%')"

        If DstUserGroup.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "      and PM.UserGroupOID=" & DstUserGroup.SelectedValue
        End If

        vnQuery += vbCrLf & " Order by PM.UserName,PM.UserID"
        pbuFillDtbSQL(vsDtb, vnQuery, vnSQLConn)

        GrvUser.DataSource = vsDtb
        GrvUser.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvUser_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvUser.PageIndexChanging
        GrvUser.PageIndex = e.NewPageIndex
        psFillGrvUser()
    End Sub

    Private Sub GrvUser_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvUser.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvUser.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(0).Controls(0), LinkButton).Text
        End If
    End Sub

    Private Sub GrvUser_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GrvUser.RowDataBound
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            'Dim LinkBtn As LinkButton
            'LinkBtn = e.Row.FindControl("LinkBtn")
            'LinkBtn.Attributes.Add("onClick", "GetSelectedData(" & e.Row.RowIndex & ");")
            'LinkBtn.Text = vsDtb.Rows(e.Row.RowIndex).Item("UserName")
        End If
    End Sub

    Protected Sub GrvUser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvUser.SelectedIndexChanged
        Response.Redirect("~/UserAdmin/WbfUserMs.aspx?vpUserOID=" + GrvUser.SelectedRow.Cells(9).Text)
    End Sub

    Protected Sub BtnBackMs_Click(sender As Object, e As EventArgs) Handles BtnBackMs.Click
        Response.Redirect("~/UserAdmin/WbfUserMs.aspx")
    End Sub
End Class