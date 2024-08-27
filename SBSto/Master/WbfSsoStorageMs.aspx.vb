Imports System.Data.SqlClient
Imports Spire.Barcode
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.IO
Public Class WbfSsoStorageMs
    Inherits System.Web.UI.Page

    Dim settings As BarcodeSettings

    Dim vsIOFileStream As System.IO.FileStream
    Dim vsFileLength As Long

    Dim vsQRDir As String

    Const csFileFormat = ".jpg"

    Enum ensColList
        OID = 12
    End Enum

    Private Sub psClearData()
        TxtOID.Text = ""
        TxtSeqNo.Text = ""
        TxtColumn.Text = ""
        TxtLevel.Text = ""
        TxtStorageNo.Text = ""
        TxtQRCodeID.Text = ""
        RdbStagging.SelectedIndex = -1
        ChkActive.Checked = False
    End Sub

    Private Sub psClearMessage()
        LblMsgBuilding.Visible = False
        LblMsgLantai.Visible = False
        LblMsgZona.Visible = False
        LblMsgStorageType.Visible = False
        LblMsgSeqNo.Visible = False
        LblMsgLevel.Visible = False
        LblMsgColumn.Visible = False
        LblMsgStorageNo.Visible = False
        LblMsgStagging.Visible = False
        LblMsgQRCodeID.Visible = False
        LblMsgErrorNE.Visible = False
    End Sub
    Private Sub psDefaultDisplay()
        DivPreview.Style(HtmlTextWriterStyle.MarginTop) = "-375px"
        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
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

        vnQuery = "Select BM.WarehouseOID,BL.BuildingOID,BL.LantaiOID,BLZ.ZonaOID,"
        vnQuery += vbCrLf & "PM.StorageTypeOID,PM.StorageSequenceNumber,PM.StorageLevel,PM.StorageColumn,PM.StorageNumber,PM.StorageStagIO,PM.StorageQRCodeID,PM.Status"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_Storage_MA PM with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA BLZ with(nolock) on BLZ.OID=PM.BuildingLantaiZonaRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Zona_MA ZM with(nolock) on ZM.OID=BLZ.ZonaOID"

        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA BL with(nolock) on BL.OID=BLZ.BuildingLantaiRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Building_MA BM with(nolock) on BM.OID=BL.BuildingOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Lantai_MA LM with(nolock) on LM.OID=BL.LantaiOID"
        vnQuery += vbCrLf & "Where PM.OID=" & TxtOID.Text

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            DstWarehouse.SelectedValue = vnDtb.Rows(0).Item("WarehouseOID")
            pbuFillDstBuilding_ByWarehouse(DstBuilding, False, DstWarehouse.SelectedValue, vnSQLConn)

            DstBuilding.SelectedValue = vnDtb.Rows(0).Item("BuildingOID")
            DstLantai.SelectedValue = vnDtb.Rows(0).Item("LantaiOID")
            DstZona.SelectedValue = vnDtb.Rows(0).Item("ZonaOID")
            DstStorageType.SelectedValue = vnDtb.Rows(0).Item("StorageTypeOID")
            TxtSeqNo.Text = fbuValStr(vnDtb.Rows(0).Item("StorageSequenceNumber"))
            TxtColumn.Text = fbuValStr(vnDtb.Rows(0).Item("StorageColumn"))
            TxtLevel.Text = fbuValStr(vnDtb.Rows(0).Item("StorageLevel"))
            TxtStorageNo.Text = fbuValStr(vnDtb.Rows(0).Item("StorageNumber"))
            TxtQRCodeID.Text = vnDtb.Rows(0).Item("StorageQRCodeID")

            If vnDtb.Rows(0).Item("StorageStagIO") = "0" Then
                RdbStagging.SelectedIndex = -1
            Else
                RdbStagging.SelectedValue = vnDtb.Rows(0).Item("StorageStagIO")
            End If

            ChkActive.Checked = IIf(vnDtb.Rows(0).Item("Status") = "ACTIVE", True, False)
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtOID.ReadOnly = Not vriBo
        TxtSeqNo.ReadOnly = Not vriBo
        TxtColumn.ReadOnly = Not vriBo
        TxtLevel.ReadOnly = Not vriBo
        TxtStorageNo.ReadOnly = Not vriBo

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
            psDefaultDisplay()

            BtnBaru.Enabled = (Session("UserAdmin") = 1)
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsGudang, vnSQLConn)

            pbuFillDstStorageType(DstStorageType, False, vnSQLConn)
            pbuFillDstStorageType(DstListStorageType, True, vnSQLConn)

            pbuFillDstWarehouse(DstWarehouse, False, vnSQLConn)
            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)

            pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)

            pbuFillDstLantai(DstLantai, False, vnSQLConn)
            pbuFillDstLantai(DstListLantai, True, vnSQLConn)

            pbuFillDstZona(DstZona, False, vnSQLConn)
            pbuFillDstZona(DstListZona, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            BtnEdit.Visible = (Session("UserAdmin") = 1)
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If DstStorageType.SelectedValue = 0 Then
            LblMsgStorageType.Text = "Pilih Storage Type"
            LblMsgStorageType.Visible = True
            vnSave = False
        End If

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
        If DstZona.SelectedValue = 0 Then
            LblMsgLantai.Text = "Pilih Zona"
            LblMsgLantai.Visible = True
            vnSave = False
        End If

        ChkIsRack.Checked = False
        ChkIsStagging.Checked = False
        ChkIsCrossDock.Checked = False

        If DstStorageType.SelectedValue = enuStorageType.Rack Then
            ChkIsRack.Checked = True
        ElseIf DstStorageType.SelectedValue = enuStorageType.Staging Then
            ChkIsStagging.Checked = True
        ElseIf DstStorageType.SelectedValue = enuStorageType.CrossDock Then
            ChkIsCrossDock.Checked = True
        ElseIf DstStorageType.SelectedValue = enuStorageType.Karantina Then
            ChkIsKarantina.Checked = True
        ElseIf DstStorageType.SelectedValue = enuStorageType.DO_Titip Then
            ChkIsDOTitip.Checked = True
        ElseIf DstStorageType.SelectedValue = enuStorageType.Damage Then
            ChkIsDamage.Checked = True
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
            Dim vnDBMaster As String = fbuGetDBMaster()

            Dim vnWarehouseOID As Integer = DstWarehouse.SelectedValue
            Dim vnBuildingOID As Integer = DstBuilding.SelectedValue
            Dim vnLantaiOID As Integer = DstLantai.SelectedValue
            Dim vnZonaOID As Integer = DstZona.SelectedValue

            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")

            Dim vnBuildingLantaiZonaRelOID = fbuGetBuildingLantaiZonaRelOID(vnBuildingOID, vnLantaiOID, vnZonaOID, vnSQLConn)
            If vnBuildingLantaiZonaRelOID = 0 Then
                LblMsgLantai.Text = "Simpan Gagal...Building " & DstBuilding.SelectedItem.Text & " - Lantai " & DstLantai.SelectedItem.Text & " - Zona " & DstZona.SelectedItem.Text & " - Belum Ada<br />Silakan Cek Daftar Building - Lantai - Zona"
                LblMsgLantai.Visible = True

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Exit Sub
            End If

            Dim vnStorageTypeOID As Integer = DstStorageType.SelectedValue
            'vnQuery = "Select abs(IsRack) From " & vnDBMaster & "Sys_StorageType_MA Where OID=" & vnStorageTypeOID

            'Dim vnIsRack As Byte = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If ChkIsRack.Checked Then
                If Trim(TxtSeqNo.Text) = "" Then
                    LblMsgSeqNo.Text = "Isi Sequence Number"
                    LblMsgSeqNo.Visible = True
                    vnSave = False
                End If
                If Trim(TxtLevel.Text) = "" Then
                    LblMsgLevel.Text = "Isi Level"
                    LblMsgLevel.Visible = True
                    vnSave = False
                End If
                If Trim(TxtColumn.Text) = "" Then
                    LblMsgColumn.Text = "Isi Column"
                    LblMsgColumn.Visible = True
                    vnSave = False
                End If

            ElseIf ChkIsStagging.Checked = True Then
                If RdbStagging.SelectedIndex = -1 Then
                    LblMsgStagging.Text = "Pilih IN / OUT"
                    LblMsgStagging.Visible = True
                    vnSave = False
                End If
            ElseIf ChkIsCrossDock.Checked = True Then
            ElseIf ChkIsKarantina.Checked = True Then
            Else
                If Trim(TxtStorageNo.Text) = "" Then
                    LblMsgStorageNo.Text = "Isi StorageNo "
                    LblMsgStorageNo.Visible = True
                    vnSave = False
                Else

                End If
            End If
            If Not vnSave Then
                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Exit Sub
            End If

            Dim vnSequenceNo As String = fbuFormatString(Trim(TxtSeqNo.Text))
            Dim vnLevel As String = fbuFormatString(Trim(TxtLevel.Text))
            Dim vnColumn As String = fbuFormatString(Trim(TxtColumn.Text))
            Dim vnStorageNo As String = fbuFormatString(Trim(TxtStorageNo.Text))

            Dim vnStorageKeyID As String = ""

            Dim vnStorageQRCodeID As String

            Dim vnCmd As SqlCommand

            Dim vnStaggingIO As String = "0"
            If HdfActionStatus.Value = cbuActionNew Then

                If ChkIsRack.Checked Then
                    vnQuery = "Select count(*) From " & vnDBMaster & "Sys_Storage_MA with(nolock)"
                    vnQuery += vbCrLf & "Where BuildingLantaiZonaRelOID='" & vnBuildingLantaiZonaRelOID & "' and"
                    vnQuery += vbCrLf & "      StorageSequenceNumber='" & vnSequenceNo & "' and StorageLevel='" & vnLevel & "' and StorageColumn='" & vnColumn & "'"
                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                        vnStorageKeyID = vnBuildingLantaiZonaRelOID & Trim(TxtSeqNo.Text) & Trim(TxtColumn.Text) & Trim(TxtLevel.Text)
                        TxtQRCodeID.Text = fbuGetHash(vnStorageKeyID)
                    Else
                        LblMsgSeqNo.Text = "Simpan Gagal...Seq-Level-Column " & TxtSeqNo.Text & "-" & TxtLevel.Text & "-" & TxtColumn.Text & " - Sudah Ada<br />Silakan Cek Daftar Storage"
                        LblMsgSeqNo.Visible = True
                        vnSave = False
                    End If

                ElseIf ChkIsStagging.Checked Then
                    vnStaggingIO = RdbStagging.SelectedValue

                    '<---18 Aug 2023 diganti, awalnya 1 staging per warehouse --> jadi 1 staging per warehouse per lantai
                    vnQuery = "Select count(*)"
                    vnQuery += vbCrLf & "From " & vnDBMaster & "Sys_Storage_MA sto with(nolock)"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA blz with(nolock) on blz.OID=sto.BuildingLantaiZonaRelOID"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA blt with(nolock) on blt.OID=blz.BuildingLantaiRelOID"
                    vnQuery += vbCrLf & "Where sto.StorageTypeOID=" & vnStorageTypeOID & " and sto.StorageStagIO=" & vnStaggingIO
                    vnQuery += vbCrLf & "     and blt.BuildingOID=" & DstBuilding.SelectedValue & " and blt.LantaiOID=" & vnLantaiOID

                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                        vnStorageKeyID = vnBuildingLantaiZonaRelOID & "STAGGING" & vnStaggingIO
                        'vnStorageKeyID = vnBuildingLantaiZonaRelOID & "xLantaiOID" & vnLantaiOID & "x" & "STAGGING" & vnStaggingIO
                        TxtQRCodeID.Text = fbuGetHash(vnStorageKeyID)
                    Else
                        LblMsgSeqNo.Text = "Simpan Gagal...Stagging " & RdbStagging.SelectedItem.Text & " - Sudah Ada di " & DstBuilding.SelectedItem.Text & " - " & DstLantai.SelectedItem.Text & "<br />Silakan Cek Daftar Storage"
                        LblMsgSeqNo.Visible = True
                        vnSave = False
                    End If

                ElseIf ChkIsCrossDock.Checked Then
                    vnQuery = "Select count(*)"
                    vnQuery += vbCrLf & "From " & vnDBMaster & "Sys_Storage_MA sto with(nolock)"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA blz with(nolock) on blz.OID=sto.BuildingLantaiZonaRelOID"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA blt with(nolock) on blt.OID=blz.BuildingLantaiRelOID"
                    vnQuery += vbCrLf & "Where sto.StorageTypeOID=" & vnStorageTypeOID & " and blt.BuildingOID=" & DstBuilding.SelectedValue
                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                        vnStorageKeyID = vnBuildingLantaiZonaRelOID & "CROSSDOCK"
                        TxtQRCodeID.Text = fbuGetHash(vnStorageKeyID)
                    Else
                        LblMsgSeqNo.Text = "Simpan Gagal...CrossDock Sudah Ada di Building " & DstBuilding.SelectedItem.Text & "<br />Silakan Cek Daftar Storage"
                        LblMsgSeqNo.Visible = True
                        vnSave = False
                    End If

                ElseIf ChkIsKarantina.Checked Then
                    vnQuery = "Select count(*)"
                    vnQuery += vbCrLf & "From " & vnDBMaster & "Sys_Storage_MA sto with(nolock)"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA blz with(nolock) on blz.OID=sto.BuildingLantaiZonaRelOID"
                    vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA blt with(nolock) on blt.OID=blz.BuildingLantaiRelOID"
                    vnQuery += vbCrLf & "Where sto.StorageTypeOID=" & vnStorageTypeOID & " and blt.BuildingOID=" & DstBuilding.SelectedValue
                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                        vnStorageKeyID = vnBuildingLantaiZonaRelOID & "KARANTINA"
                        TxtQRCodeID.Text = fbuGetHash(vnStorageKeyID)
                    Else
                        LblMsgSeqNo.Text = "Simpan Gagal...Karantina Sudah Ada di Building " & DstBuilding.SelectedItem.Text & "<br />Silakan Cek Daftar Storage"
                        LblMsgSeqNo.Visible = True
                        vnSave = False
                    End If

                Else
                    '<---16 Feb 2023 Pallet ga ada lagi --> diganti Floor
                    'vnQuery = "Select count(*) From " & vnDBMaster & "Sys_Storage_MA Where StorageNumber='" & vnStorageNo & "'"
                    vnQuery = "Select count(*) From " & vnDBMaster & "Sys_Storage_MA with(nolock)"
                    vnQuery += vbCrLf & "Where BuildingLantaiZonaRelOID='" & vnBuildingLantaiZonaRelOID & "' and"
                    vnQuery += vbCrLf & "      StorageNumber='" & vnStorageNo & "'"
                    '<<==16 Feb 2023 Pallet ga ada lagi --> diganti Floor

                    If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                        vnStorageKeyID = vnBuildingLantaiZonaRelOID & Trim(TxtStorageNo.Text)
                        TxtQRCodeID.Text = fbuGetHash(vnStorageKeyID)
                    Else
                        LblMsgStorageNo.Text = "Simpan Gagal...Storage Number " & TxtSeqNo.Text & " - Sudah Ada<br />Silakan Cek Daftar Storage"
                        LblMsgStorageNo.Visible = True
                        vnSave = False
                    End If
                End If

                vnStorageQRCodeID = fbuFormatString(Trim(TxtQRCodeID.Text))

                vnQuery = "Select count(*) From " & vnDBMaster & "Sys_Storage_MA with(nolock) Where StorageQRCodeID='" & vnStorageQRCodeID & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgQRCodeID.Text = "Simpan Gagal...QR Code ID " & TxtQRCodeID.Text & " - Sudah Ada<br />Silakan Cek Daftar Storage"
                    LblMsgQRCodeID.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into " & vnDBMaster & "Sys_Storage_MA("
                vnQuery += vbCrLf & "BuildingLantaiZonaRelOID,"
                vnQuery += vbCrLf & "StorageTypeOID,"
                vnQuery += vbCrLf & "StorageSequenceNumber,"
                vnQuery += vbCrLf & "StorageLevel,"
                vnQuery += vbCrLf & "StorageColumn,"
                vnQuery += vbCrLf & "StorageNumber,"
                vnQuery += vbCrLf & "StorageStagIO,"
                vnQuery += vbCrLf & "StorageQRCodeID,"
                vnQuery += vbCrLf & "StorageQRCodeIDImg,"
                vnQuery += vbCrLf & "Status,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values("
                vnQuery += vbCrLf & vnBuildingLantaiZonaRelOID & ","
                vnQuery += vbCrLf & vnStorageTypeOID & ","
                vnQuery += vbCrLf & "'" & vnSequenceNo & "',"
                vnQuery += vbCrLf & "'" & vnLevel & "',"
                vnQuery += vbCrLf & "'" & vnColumn & "',"
                vnQuery += vbCrLf & "'" & vnStorageNo & "',"
                vnQuery += vbCrLf & "" & vnStaggingIO & ","
                vnQuery += vbCrLf & "'" & vnStorageQRCodeID & "',"
                vnQuery += vbCrLf & "@vnStorageQRCodeIDImg,"
                vnQuery += vbCrLf & "'" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"

                '<---Generate Image
                Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")
                Dim vsIOFileStream As System.IO.FileStream
                Dim vsFileLength As Long
                Const csFileFormat = ".jpg"

                Dim vnFileName As String
                Dim vnFileByte() As Byte

                vnFileName = "Storage_" & vnStorageKeyID & "_" & Format(Date.Now, "yyyyMMdd_HHmmss") & "~sm" & csFileFormat

                Dim vnQRDir As String = ""

                pbuGenerateQRCode(vnFileName, vnStorageQRCodeID, vnQRDir)

                vsIOFileStream = System.IO.File.OpenRead(vnQRDir & vnFileName)

                vsFileLength = vsIOFileStream.Length
                ReDim vnFileByte(vsFileLength)

                vsIOFileStream.Read(vnFileByte, 0, vsFileLength)
                '<<==Generate Image

                vnCmd = New SqlClient.SqlCommand(vnQuery, vnSQLConn)
                vnCmd.Parameters.AddWithValue("@vnStorageQRCodeIDImg", vnFileByte)
                vnCmd.Transaction = vnSQLTrans
                vnCmd.ExecuteNonQuery()

                vnSQLTrans.Commit()
                vnSQLTrans = Nothing
                vnBeginTrans = False

                Dim vnOID As Integer
                vnQuery = "Select OID From " & vnDBMaster & "Sys_Storage_MA with(nolock)"
                vnQuery += vbCrLf & "Where BuildingLantaiZonaRelOID=" & vnBuildingLantaiZonaRelOID & " and StorageTypeOID=" & vnStorageTypeOID & " and StorageQRCodeID='" & vnStorageQRCodeID & "' and CreationUserOID=" & vnUserOID
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

                TxtOID.Text = vnOID
            Else
                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update " & vnDBMaster & "Sys_Storage_MA set"
                vnQuery += vbCrLf & "Status='" & IIf(ChkActive.Checked, "ACTIVE", "NOT ACTIVE") & "',"
                vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & vnUserOID
                vnQuery += vbCrLf & " Where OID=" & TxtOID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnSQLTrans.Commit()
                vnSQLTrans = Nothing
                vnBeginTrans = False
            End If

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

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
            End If
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName,"
        vnQuery += vbCrLf & "OM.StorageTypeName,abs(OM.IsRack)vIsRack,PM.StorageSequenceNumber,PM.StorageColumn,PM.StorageLevel,PM.StorageNumber,"
        vnQuery += vbCrLf & "case when PM.StorageStagIO=0 then ''"
        vnQuery += vbCrLf & "     when PM.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
        vnQuery += vbCrLf & "PM.StorageQRCodeID,"
        vnQuery += vbCrLf & "PM.OID,PM.Status,PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_Storage_MA PM with(nolock)"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_StorageType_MA OM with(nolock) on OM.OID=PM.StorageTypeOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiZonaRel_MA BLZ with(nolock) on BLZ.OID=PM.BuildingLantaiZonaRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Zona_MA ZM with(nolock) on ZM.OID=BLZ.ZonaOID"

        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_BuildingLantaiRel_MA BL with(nolock) on BL.OID=BLZ.BuildingLantaiRelOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Building_MA BM with(nolock) on BM.OID=BL.BuildingOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Lantai_MA LM with(nolock) on LM.OID=BL.LantaiOID"
        vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=BM.WarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM with(nolock) on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM with(nolock) on AM.OID=PM.ModificationUserOID"
        vnQuery += vbCrLf & "Where 1=1"

        If ChkListQRCodeID.Checked Then
            vnQuery += vbCrLf & "            and PM.StorageQRCodeID='" & Trim(TxtListQRCodeID.Text) & "'"
        Else
            If ChkListStorageOID.Checked Then
                vnQuery += vbCrLf & "            and PM.OID=" & Val(TxtListStorageOID.Text)
            Else
                If Val(DstListWarehouse.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and BM.WarehouseOID=" & DstListWarehouse.SelectedValue
                End If
                If Val(DstListBuilding.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and BL.BuildingOID=" & DstListBuilding.SelectedValue
                End If
                If Val(DstListLantai.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and BL.LantaiOID=" & DstListLantai.SelectedValue
                End If
                If Val(DstListZona.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and BLZ.ZonaOID=" & DstListZona.SelectedValue
                End If
                If Val(DstListStorageType.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and PM.StorageTypeOID=" & DstListStorageType.SelectedValue
                End If

                If ChkQRNull.Checked Then
                    vnQuery += vbCrLf & "            and isnull(PM.StorageQRCodeID,'')=''"
                End If

                If DstListStorageType.SelectedValue = enuStorageType.Rack Then
                    If Trim(TxtListRackY_SeqNo.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(PM.StorageSequenceNumber,'')='" & fbuFormatString(Trim(TxtListRackY_SeqNo.Text)) & "'"
                    End If
                    If Trim(TxtListRackY_Level.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(PM.StorageLevel,'')='" & fbuFormatString(Trim(TxtListRackY_Level.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName,PM.StorageSequenceNumber,PM.StorageColumn,PM.StorageLevel"

                ElseIf DstListStorageType.SelectedValue = enuStorageType.Floor Then
                    If Trim(TxtListRackN_Start.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(PM.StorageNumber,'')>='" & fbuFormatString(Trim(TxtListRackN_Start.Text)) & "'"
                    End If
                    If Trim(TxtListRackN_End.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(PM.StorageNumber,'')<='" & fbuFormatString(Trim(TxtListRackN_End.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName,PM.StorageNumber"

                Else
                    vnQuery += vbCrLf & " Order by WM.WarehouseName,BM.BuildingName,LM.LantaiDescription,ZM.ZonaName"
                End If
            End If
        End If

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

    Private Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub
    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psPreview()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        psClearMessage()
        Dim vnCrpFileName As String = ""
        psGenerateCrp(vnCrpFileName)

        Dim vnRootURL As String = ConfigurationManager.AppSettings("WebRootFolder")
        Dim vnParam As String
        vnParam = "vqCrpPreviewType=" & stuCrpPreviewType.ByQueryPopwin
        vnParam += "&vqCrpFileName=" & vnCrpFileName
        vnParam += "&vqCrpSubReport1="
        vnParam += "&vqCrpSubReport2="
        vnParam += "&vqCrpSubReport3="
        vnParam += "&vqCrpSubReport4="
        vnParam += "&vqCrpQuery=" & vbuCrpQuery
        vnParam += "&vqCrpQuery1="
        vnParam += "&vqCrpQuery2="
        vnParam += "&vqCrpQuery3="
        vnParam += "&vqCrpQuery4="
        vnParam += "&vqCrpPreview=Pdf"

        vbuPreviewOnClose = "0"

        ifrPreview.Src = vnRootURL & "Preview/WbfCrpViewer.aspx?" & vnParam
        psShowPreview(True)
    End Sub

    Private Sub psGenerateCrp(ByRef vriCrpFileName As String)
        Dim vn As Integer
        Dim vnChkPrint As CheckBox
        Dim vnGRow As GridViewRow
        Dim vnOIDList As String = "x"

        For vn = 0 To GrvList.Rows.Count - 1
            vnGRow = GrvList.Rows(vn)
            vnChkPrint = vnGRow.FindControl("ChkPrint")
            If vnChkPrint.Checked Then
                vnOIDList += vnGRow.Cells(ensColList.OID).Text & "x"
            End If
        Next

        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoStorage

        vbuCrpQuery = "Select Floor(vRowNum/5)vPageNum,*"
        vbuCrpQuery += "       From " & vnDBMaster & "fnTbl_SsoStorage('" & vnOIDList & "','" & Session("UserID") & "')"
        vbuCrpQuery += " order by vStorageOID"
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psPreview()
    End Sub

    Protected Sub DstStorageType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstStorageType.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuSetStorageTypeIs(DstStorageType.SelectedValue, ChkIsMultiLevel, ChkIsRack, ChkIsStagging, ChkIsCrossDock, ChkIsKarantina, ChkIsDOTitip, ChkIsDamage, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        TxtSeqNo.Text = ""
        TxtColumn.Text = ""
        TxtLevel.Text = ""
        TxtStorageNo.Text = ""

        RdbStagging.SelectedIndex = -1

        If ChkIsRack.Checked = True Then
            TxtSeqNo.Enabled = True
            TxtColumn.Enabled = True
            TxtLevel.Enabled = True
            TxtStorageNo.Enabled = False

            RdbStagging.Enabled = False

        ElseIf ChkIsStagging.Checked = True Then
            TxtSeqNo.Enabled = False
            TxtColumn.Enabled = False
            TxtLevel.Enabled = False
            TxtStorageNo.Enabled = False

            RdbStagging.Enabled = True

        ElseIf ChkIsCrossDock.Checked = True Or ChkIsKarantina.Checked = True Then
            TxtSeqNo.Enabled = False
            TxtColumn.Enabled = False
            TxtLevel.Enabled = False
            TxtStorageNo.Enabled = False

            RdbStagging.Enabled = False

        Else
            TxtSeqNo.Enabled = False
            TxtColumn.Enabled = False
            TxtLevel.Enabled = False
            TxtStorageNo.Enabled = True

            RdbStagging.Enabled = False
        End If
    End Sub

    Protected Sub BtnGenQR_Click(sender As Object, e As EventArgs) Handles BtnGenQR.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psClearMessage()

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Dim vnCmd As SqlCommand
        Try
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDtb As New DataTable
            Dim vnQuery As String
            vnQuery = "Select PM.OID,PM.BuildingLantaiZonaRelOID,"
            vnQuery += vbCrLf & "abs(OM.IsRack)vIsRack,PM.StorageSequenceNumber,PM.StorageColumn,PM.StorageLevel,PM.StorageNumber"
            vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_Storage_MA PM with(nolock)"
            vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_StorageType_MA OM with(nolock) on OM.OID=PM.StorageTypeOID"
            vnQuery += vbCrLf & " Where isnull(StorageQRCodeID,'')=''"

            Dim vnIsRack As Byte

            Dim vnStorageOID As Integer
            Dim vnBuildingLantaiZonaRelOID As Integer
            Dim vnSequenceNo As String
            Dim vnLevel As String
            Dim vnColumn As String
            Dim vnStorageNo As String
            Dim vnStorageKeyID As String
            Dim vnStorageQRCodeID As String

            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
            Dim vnDRow As DataRow
            For vn = 0 To vnDtb.Rows.Count - 1
                vnDRow = vnDtb.Rows(vn)
                vnStorageOID = vnDRow.Item("OID")
                vnBuildingLantaiZonaRelOID = vnDRow.Item("BuildingLantaiZonaRelOID")
                vnIsRack = vnDRow.Item("vIsRack")
                If vnIsRack = 0 Then
                    vnStorageNo = vnDRow.Item("StorageNumber")
                    vnStorageKeyID = fbuGetHash(vnBuildingLantaiZonaRelOID & vnStorageNo)
                Else
                    vnSequenceNo = vnDRow.Item("StorageSequenceNumber")
                    vnColumn = vnDRow.Item("StorageColumn")
                    vnLevel = vnDRow.Item("StorageLevel")
                    vnStorageKeyID = fbuGetHash(vnBuildingLantaiZonaRelOID & vnSequenceNo & vnColumn & vnLevel)
                End If
                vnStorageQRCodeID = fbuFormatString(Trim(vnStorageKeyID))

                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_Storage_MA with(nolock) Where StorageQRCodeID='" & vnStorageQRCodeID & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                    vnSQLTrans = vnSQLConn.BeginTransaction()
                    vnBeginTrans = True

                    vnQuery = "Update " & vnDBMaster & "Sys_Storage_MA set"
                    vnQuery += vbCrLf & "StorageQRCodeID='" & vnStorageQRCodeID & "',"
                    vnQuery += vbCrLf & "StorageQRCodeIDImg=@vnStorageQRCodeIDImg"
                    vnQuery += vbCrLf & " Where OID=" & vnStorageOID

                    '<---Generate Image
                    Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")
                    Dim vsIOFileStream As System.IO.FileStream
                    Dim vsFileLength As Long
                    Const csFileFormat = ".jpg"

                    Dim vnFileName As String
                    Dim vnFileByte() As Byte

                    vnFileName = "Storage_" & vnStorageKeyID & "_" & Format(Date.Now, "yyyyMMdd_HHmmss") & "~sm" & csFileFormat

                    Dim vnQRDir As String = ""

                    pbuGenerateQRCode(vnFileName, vnStorageQRCodeID, vnQRDir)

                    vsIOFileStream = System.IO.File.OpenRead(vnQRDir & vnFileName)

                    vsFileLength = vsIOFileStream.Length
                    ReDim vnFileByte(vsFileLength)

                    vsIOFileStream.Read(vnFileByte, 0, vsFileLength)
                    '<<==Generate Image

                    vnCmd = New SqlClient.SqlCommand(vnQuery, vnSQLConn)
                    vnCmd.Parameters.AddWithValue("@vnStorageQRCodeIDImg", vnFileByte)
                    vnCmd.Transaction = vnSQLTrans
                    vnCmd.ExecuteNonQuery()

                    vnSQLTrans.Commit()
                    vnSQLTrans = Nothing
                    vnBeginTrans = False
                End If
            Next

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

        Catch ex As Exception
            LblMsgErrorNE.Text = ex.Message
            LblMsgErrorNE.Visible = ex.Message
            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
                vnBeginTrans = False
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub DstListStorageType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstListStorageType.SelectedIndexChanged
        If DstListStorageType.SelectedValue = enuStorageType.Floor Then
            PanListRackN.Visible = True
            PanListRackY.Visible = False
        ElseIf DstListStorageType.SelectedValue = enuStorageType.Rack Then
            PanListRackN.Visible = False
            PanListRackY.Visible = True
        ElseIf DstListStorageType.SelectedValue = enuStorageType.Staging Then
            PanListRackN.Visible = False
            PanListRackY.Visible = False
        Else
            PanListRackN.Visible = False
            PanListRackY.Visible = False
        End If
    End Sub

    Protected Sub ChkListCheckAll_CheckedChanged(sender As Object, e As EventArgs) Handles ChkListCheckAll.CheckedChanged
        Dim vnChecked As Boolean
        vnChecked = ChkListCheckAll.Checked

        Dim vn As Integer
        Dim vnChkPrint As CheckBox
        Dim vnGRow As GridViewRow

        For vn = 0 To GrvList.Rows.Count - 1
            vnGRow = GrvList.Rows(vn)
            vnChkPrint = vnGRow.FindControl("ChkPrint")
            vnChkPrint.Checked = vnChecked
        Next
    End Sub

    Private Sub DstListWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstListWarehouse.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        If DstListWarehouse.SelectedValue = "0" Then
            pbuFillDstBuilding(DstListBuilding, True, vnSQLConn)
        Else
            pbuFillDstBuilding_ByWarehouse(DstListBuilding, True, DstListWarehouse.SelectedValue, vnSQLConn)
        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkListQRCodeID_CheckedChanged(sender As Object, e As EventArgs) Handles ChkListQRCodeID.CheckedChanged

    End Sub
End Class