Imports System.Data.SqlClient
Public Class WbfSsoGudangMs
    Inherits System.Web.UI.Page

    Private Sub psClearData()
        TxtGdgCode.Text = ""
        TxtGdgName.Text = ""
        TxtGdgAddress.Text = ""
        TxtGdgCity.Text = ""
        TxtOID.Text = ""
        ChkActive.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgGdgCode.Visible = False
        LblMsgGdgName.Visible = False
        LblMsgGdgAddress.Visible = False
        LblMsgGdgCity.Visible = False
        LblMsgLocation.Visible = False
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

        vnQuery = "Select * From " & fbuGetDBMaster() & "Sys_MstGudang_MA Where OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtGdgCode.Text = fbuValStr(vnDtb.Rows(0).Item("GdgCode"))
            TxtGdgName.Text = fbuValStr(vnDtb.Rows(0).Item("GdgName"))
            TxtGdgAddress.Text = fbuValStr(vnDtb.Rows(0).Item("GdgAddress"))
            TxtGdgCity.Text = fbuValStr(vnDtb.Rows(0).Item("GdgCity"))
            DstLocation.SelectedValue = vnDtb.Rows(0).Item("LocationOID")
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtGdgCode.ReadOnly = Not vriBo
        TxtGdgName.ReadOnly = Not vriBo
        TxtGdgAddress.ReadOnly = Not vriBo
        TxtGdgCity.ReadOnly = Not vriBo
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

            pbuFillDstLocation(DstLocation, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtGdgCode.Text)) = 0 Then
            LblMsgGdgCode.Text = "Isi Nama Gudang"
            LblMsgGdgCode.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtGdgName.Text)) = 0 Then
            LblMsgGdgName.Text = "Isi Nama Gudang"
            LblMsgGdgName.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtGdgAddress.Text)) = 0 Then
            LblMsgGdgAddress.Text = "Isi Alamat Gudang"
            LblMsgGdgAddress.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtGdgCity.Text)) = 0 Then
            LblMsgGdgCity.Text = "Isi Kota Gudang"
            LblMsgGdgCity.Visible = True
            vnSave = False
        End If
        If DstLocation.SelectedValue = 0 Then
            LblMsgLocation.Text = "Pilih Lokasi"
            LblMsgLocation.Visible = True
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

            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & fbuGetDBMaster() & "Sys_MstGudang_MA Where GdgName='" & fbuFormatString(Trim(TxtGdgName.Text)) & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgGdgName.Text = "Nama Gudang " & Trim(TxtGdgName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Gudang"
                    LblMsgGdgName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
                Dim vnOID As Integer
                vnQuery = "Select isnull(max(OID),0)+1 From " & fbuGetDBMaster() & "Sys_MstGudang_MA"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstGudang_MA("
                vnQuery += vbCrLf & "OID,GdgCode,GdgName,"
                vnQuery += vbCrLf & "GdgAddress,GdgCity,"
                vnQuery += vbCrLf & "LocationOID,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values(" & vnOID & ","
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtGdgCode.Text)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtGdgName.Text)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtGdgAddress.Text)) & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtGdgCity.Text)) & "',"
                vnQuery += vbCrLf & DstLocation.SelectedValue & ","
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from " & fbuGetDBMaster() & "Sys_MstGudang_MA Where GdgName='" & fbuFormatString(Trim(TxtGdgName.Text)) & "' and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgGdgName.Text = "Nama Gudang " & Trim(TxtGdgName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Gudang"
                    LblMsgGdgName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & fbuGetDBMaster() & "Sys_MstGudang_MA set "
                vnQuery += vbCrLf & "GdgCode='" & fbuFormatString(Trim(TxtGdgCode.Text)) & "',"
                vnQuery += vbCrLf & "GdgName='" & fbuFormatString(Trim(TxtGdgName.Text)) & "',"
                vnQuery += vbCrLf & "GdgAddress='" & fbuFormatString(Trim(TxtGdgAddress.Text)) & "',"
                vnQuery += vbCrLf & "GdgCity='" & fbuFormatString(Trim(TxtGdgCity.Text)) & "',"
                vnQuery += vbCrLf & "LocationOID=" & DstLocation.SelectedValue & ","
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

        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.GdgCode,PM.GdgName,PM.GdgAddress,PM.GdgCity,LM.LocationName,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstGudang_MA PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoLocation_MA LM on LM.OID=PM.LocationOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "      Where PM.GdgCode like '%" & vnCriteria & "%' or PM.GdgName like '%" & vnCriteria & "%'"
        vnQuery += vbCrLf & " Order by PM.GdgName"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

End Class