Imports System.Data.SqlClient
Public Class WbfSsoStorageCompanyMs
    Inherits System.Web.UI.Page


    Private Sub psClearMessage()
        LblMsgErrorNE.Visible = False
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Master"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            BtnEditCompany.Enabled = (Session("UserAdmin") = 1)
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsGudang, vnSQLConn)

            pbuFillDstStorageType(DstStorageType, False, vnSQLConn)
            pbuFillDstWarehouse(DstWarehouse, False, vnSQLConn)
            pbuFillDstLantai(DstLantai, False, vnSQLConn)
            pbuFillDstZona(DstZona, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
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

    Private Sub psEnableSaveCompany(vriBo As Boolean)
        BtnEditCompany.Visible = Not vriBo
        BtnSimpanCompany.Visible = vriBo
        BtnBatalCompany.Visible = vriBo

        BtnDisplay.Visible = Not vriBo
        GrvLsStorage.Enabled = Not vriBo
    End Sub

    Private Sub psFillGrvCompany(vriEdit As Boolean, vriStorageOID As Integer, vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnDtbB As New DataTable

        Dim vnQuery As String = ""
        If vriEdit Then
            vnQuery = "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany"
            vnQuery += vbCrLf & "Order by CompanyName"
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        Else
            vnQuery = "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany"
            vnQuery += vbCrLf & "Where CompanyCode in (Select b.CompanyCode From Sys_CompanyStorageRel_MA b Where b.StorageOID=" & vriStorageOID & ")"
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
                vnQuery = "Select CompanyCode From Sys_CompanyStorageRel_MA Where StorageOID=" & vriStorageOID
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

    Protected Sub BtnEditCompany_Click(sender As Object, e As EventArgs) Handles BtnEditCompany.Click
        If Val(HdfStorageOID.Value) = 0 Then Exit Sub
        If Not BtnEditCompany.Visible Then Exit Sub
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvCompany(True, HdfStorageOID.Value, vnSQLConn)

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

        psFillGrvCompany(False, HdfStorageOID.Value, vnSQLConn)

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
            Dim vnRow As GridViewRow

            Dim vnSQLTrans As SqlTransaction = Nothing
            Dim vnBeginTrans As Boolean
            Try
                vnSQLTrans = vnSQLConn.BeginTransaction("Company")
                vnBeginTrans = True

                vnQuery = "Delete Sys_CompanyStorageRel_MA Where StorageOID=" & HdfStorageOID.Value
                pbuExecuteSQLTrans(vnQuery, 3, vnSQLConn, vnSQLTrans)

                For vn = 0 To GrvCompany.Rows.Count - 1
                    vnRow = GrvCompany.Rows(vn)
                    vnChkCompany = vnRow.FindControl("ChkCompany")

                    Debug.Print(vn & " " & vnRow.Cells(2).Text & " " & vnRow.Cells(1).Text)

                    If vnChkCompany.Checked = True Then
                        vnQuery = "Insert into Sys_CompanyStorageRel_MA(CompanyCode,StorageOID,CreationDatetime,CreationUserOID)"
                        vnQuery += vbCrLf & "values('" & vnRow.Cells(2).Text & "'," & HdfStorageOID.Value & ",getdate()," & Session("UserOID") & ")"
                        pbuExecuteSQLTrans(vnQuery, 1, vnSQLConn, vnSQLTrans)
                    End If
                Next
                vnSQLTrans.Commit()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                vnBeginTrans = False

                psFillGrvCompany(False, HdfStorageOID.Value, vnSQLConn)
                vnSQLConn.Close()
                vnSQLConn = Nothing

            Catch ex As Exception
                LblMsgErrorNE.Text = ex.Message
                LblMsgErrorNE.Visible = True
                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing
                End If
            End Try
        End If
        psEnableSaveCompany(False)
    End Sub

    Protected Sub BtnDisplay_Click(sender As Object, e As EventArgs) Handles BtnDisplay.Click
        psFillGrvLsStorage()
    End Sub

    Private Sub psFillGrvLsStorage()
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
        vnQuery = "Select WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName,"
        vnQuery += vbCrLf & "OM.StorageTypeName,PM.StorageSequenceNumber,PM.StorageLevel,PM.StorageColumn,PM.StorageNumber,"
        vnQuery += vbCrLf & "PM.OID"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_Storage_MA PM"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_StorageType_MA OM on OM.OID=PM.StorageTypeOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA BLZ on BLZ.OID=PM.BuildingLantaiZonaRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Zona_MA ZM on ZM.OID=BLZ.ZonaOID"

        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA BL on BL.OID=BLZ.BuildingLantaiRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Building_MA BM on BM.OID=BL.BuildingOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Lantai_MA LM on LM.OID=BL.LantaiOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA WM on WM.OID=BM.WarehouseOID"
        vnQuery += vbCrLf & "      Where 1=1"
        If Val(DstWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and BM.WarehouseOID=" & DstWarehouse.SelectedValue
        End If
        If Val(DstBuilding.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and BL.Building=" & DstBuilding.SelectedValue
        End If
        If Val(DstLantai.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and BL.LantaiOID=" & DstLantai.SelectedValue
        End If
        If Val(DstZona.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and BLZ.ZonaOID=" & DstZona.SelectedValue
        End If
        If Val(DstStorageType.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and PM.StorageTypeOID=" & DstStorageType.SelectedValue
        End If

        vnQuery += vbCrLf & " Order by WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName,PM.StorageSequenceNumber"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsStorage.DataSource = vnDtb
        GrvLsStorage.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvLsStorage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsStorage.SelectedIndexChanged

    End Sub

    Private Sub GrvLsStorage_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsStorage.RowCommand
        If e.CommandName = "Select" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            HdfStorageOID.Value = GrvLsStorage.Rows(vnIdx).Cells(0).Text
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvCompany(False, HdfStorageOID.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub GrvLsStorage_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsStorage.PageIndexChanging
        GrvLsStorage.PageIndex = e.NewPageIndex
        psFillGrvLsStorage()
    End Sub
End Class