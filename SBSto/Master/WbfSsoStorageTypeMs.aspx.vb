Imports System.Data.SqlClient
Public Class WbfSsoStorageTypeMs
    Inherits System.Web.UI.Page
    Enum ensColList
        OID = 7
    End Enum
    Private Sub psClearData()
        TxtStorageTypeName.Text = ""
        TxtOID.Text = ""
        ChkActive.Checked = False
        ChkIsCrossDock.Checked = False
        ChkIsMultiLevel.Checked = False
        ChkIsRack.Checked = False
        ChkIsStagging.Checked = False
        ChkIsKarantina.Checked = False
        ChkIsDOTitip.Checked = False
        ChkIsDamage.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgStorageTypeName.Visible = False
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

        vnQuery = "Select *,"
        vnQuery += vbCrLf & "abs(IsMultiLevel)vIsMultiLevel,abs(IsRack)vIsRack,abs(IsStagging)vIsStagging,abs(IsCrossDock)vIsCrossDock,abs(IsKarantina)vIsKarantina,abs(IsDOTitip)vIsDOTitip"
        vnQuery += vbCrLf & "From " & fbuGetDBMaster() & "Sys_StorageType_MA Where OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtStorageTypeName.Text = fbuValStr(vnDtb.Rows(0).Item("StorageTypeName"))
            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)

            ChkIsCrossDock.Checked = IIf(vnDtb.Rows(0).Item("vIsCrossDock") = "1", True, False)
            ChkIsMultiLevel.Checked = IIf(vnDtb.Rows(0).Item("vIsMultiLevel") = "1", True, False)
            ChkIsRack.Checked = IIf(vnDtb.Rows(0).Item("vIsRack") = "1", True, False)
            ChkIsStagging.Checked = IIf(vnDtb.Rows(0).Item("vIsStagging") = "1", True, False)
            ChkIsKarantina.Checked = IIf(vnDtb.Rows(0).Item("vIsKarantina") = "1", True, False)
            ChkIsDOTitip.Checked = IIf(vnDtb.Rows(0).Item("vIsDOTitip") = "1", True, False)
            ChkIsDamage.Checked = IIf(vnDtb.Rows(0).Item("vIsDamage") = "1", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtStorageTypeName.ReadOnly = Not vriBo
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

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtStorageTypeName.Text)) = 0 Then
            LblMsgStorageTypeName.Text = "Isi Nama Storage Type"
            LblMsgStorageTypeName.Visible = True
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

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            Dim vnLantaiDescr As String = fbuFormatString(Trim(TxtStorageTypeName.Text))
            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_StorageType_MA Where StorageTypeName='" & vnLantaiDescr & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgStorageTypeName.Text = "Nama Storage Type " & Trim(TxtStorageTypeName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Storage Type"
                    LblMsgStorageTypeName.Visible = True
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

                vnQuery = "Insert into " & vnDBMaster & "Sys_StorageType_MA("
                vnQuery += vbCrLf & "StorageTypeName,"
                vnQuery += vbCrLf & "IsMultiLevel,"
                vnQuery += vbCrLf & "IsRack,"
                vnQuery += vbCrLf & "IsStagging,"
                vnQuery += vbCrLf & "IsCrossDock,"
                vnQuery += vbCrLf & "IsKarantina,"
                vnQuery += vbCrLf & "IsDOTitip,"
                vnQuery += vbCrLf & "IsDamage,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values("
                vnQuery += vbCrLf & "'" & vnLantaiDescr & "',"
                vnQuery += vbCrLf & IIf(ChkIsMultiLevel.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsRack.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsStagging.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsCrossDock.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsKarantina.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsDOTitip.Checked, 1, 0) & ","
                vnQuery += vbCrLf & IIf(ChkIsDamage.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnQuery = "Select OID From " & vnDBMaster & "Sys_StorageType_MA Where StorageTypeName='" & vnLantaiDescr & "'"
                vnOID = fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans)

                TxtOID.Text = vnOID
            Else
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_StorageType_MA Where StorageTypeName='" & vnLantaiDescr & "' and OID<>" & TxtOID.Text
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgStorageTypeName.Text = "Nama Storage Type " & Trim(TxtStorageTypeName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Nama Storage Type"
                    LblMsgStorageTypeName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & vnDBMaster & "Sys_StorageType_MA set "
                vnQuery += vbCrLf & "StorageTypeName='" & vnLantaiDescr & "',"
                vnQuery += vbCrLf & "IsMultiLevel=" & IIf(ChkIsMultiLevel.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsRack=" & IIf(ChkIsRack.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsStagging=" & IIf(ChkIsStagging.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsCrossDock=" & IIf(ChkIsCrossDock.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsKarantina=" & IIf(ChkIsKarantina.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsDOTitip=" & IIf(ChkIsDOTitip.Checked, 1, 0) & ","
                vnQuery += vbCrLf & "IsDamage=" & IIf(ChkIsDamage.Checked, 1, 0) & ","
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
        TxtOID.Text = GrvList.SelectedRow.Cells(ensColList.OID).Text
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
        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.StorageTypeName,"
        vnQuery += vbCrLf & "case when abs(PM.IsMultiLevel)=1 then 'Y' else 'N' end vIsMultiLevel,"
        vnQuery += vbCrLf & "case when abs(PM.IsRack)=1 then 'Y' else 'N' end vIsRack,"
        vnQuery += vbCrLf & "case when abs(PM.IsStagging)=1 then 'Y' else 'N' end vIsStagging,"
        vnQuery += vbCrLf & "case when abs(PM.IsCrossDock)=1 then 'Y' else 'N' end vIsCrossDock,"
        vnQuery += vbCrLf & "case when abs(PM.IsKarantina)=1 then 'Y' else 'N' end vIsKarantina,"
        vnQuery += vbCrLf & "case when abs(PM.IsDOTitip)=1 then 'Y' else 'N' end vIsDOTitip,"
        vnQuery += vbCrLf & "case when abs(PM.IsDamage)=1 then 'Y' else 'N' end vIsDamage,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_StorageType_MA PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "      Where PM.StorageTypeName like '%" & vnCriteria & "%'"
        vnQuery += vbCrLf & " Order by PM.StorageTypeName"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkIsMultiLevel_CheckedChanged(sender As Object, e As EventArgs) Handles ChkIsMultiLevel.CheckedChanged

    End Sub
End Class