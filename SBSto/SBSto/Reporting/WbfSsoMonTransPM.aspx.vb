Imports System.Data.SqlClient
Public Class WbfSsoMonTransPM
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoMonTrans"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Reporting"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            psDefaultDisplay()

            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -1, Date.Now), "dd MMM yyyy")
            TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoStorageStock, vnSQLConn)

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWarehouse, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
    End Sub

    Private Sub psDefaultDisplay()

    End Sub

    Protected Sub BtnRefreshAll_Click(sender As Object, e As EventArgs) Handles BtnRefreshAll.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvPM(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvPM(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnCrStatus As String = ""

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        If ChkPM_InProgress.Checked = False And ChkPM_Done.Checked = False And ChkPM_Batal.Checked = False Then
            ChkPM_InProgress.Checked = True
            ChkPM_Done.Checked = True
            ChkPM_Batal.Checked = True
        End If

        Dim vnCrStatus_1_Penerimaan_Pembelian As String = ""
        Dim vnCrStatus_2_Putaway As String = ""
        Dim vnCrStatus_3_Move As String = ""
        Dim vnCrStatus_4_PickList_Picking_Dispatch As String = ""
        Dim vnCrStatus_5_Picking As String = ""
        Dim vnCrStatus_6_Dispatch As String = ""
        Dim vnCrStatus_7_Dispatch_Receive As String = ""
        Dim vnCrStatus_8_Move_Antar_Staging_Out As String = ""
        Dim vnCrStatus_9_Putaway_DO_Titip As String = ""
        Dim vnCrStatus_10_Putaway_Void As String = ""
        Dim vnCrStatus_11_Putaway_Dispatch_Receive As String = ""

        If ChkPM_InProgress.Checked And ChkPM_Done.Checked And ChkPM_Batal.Checked Then
            vnCrStatus = ""
        Else
            '<---1 Kriteria Status Penerimaan Pembelian
            If ChkPM_InProgress.Checked Then
                vnCrStatus_1_Penerimaan_Pembelian = "PM.TransStatus=" & enuTCRCPO.On_Receive
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_1_Penerimaan_Pembelian = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCRCPO.On_Receive
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_1_Penerimaan_Pembelian = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCRCPO.On_Receive
            End If
            vnCrStatus_1_Penerimaan_Pembelian = "            and (" & vnCrStatus_1_Penerimaan_Pembelian & ")"

            '<---2 Kriteria Status Putaway Penerimaan Pembelian
            If ChkPM_InProgress.Checked Then
                vnCrStatus_2_Putaway = "PM.TransStatus in(" & enuTCPWAY.Baru & "," & enuTCPYAY.Baru & "," & enuTCPWAY.On_Putaway & "," & enuTCPYAY.On_Putaway & "," & enuTCPYAY.On_Delivery_Putaway & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_2_Putaway = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPWAY.Putaway_Done & " or PM.TransStatus>=" & enuTCPYAY.Putaway_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_2_Putaway = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPWAY.Baru
            End If
            vnCrStatus_2_Putaway = "            and (" & vnCrStatus_2_Putaway & ")"

            '<---3 Kriteria Status Pindah Lokasi
            If ChkPM_InProgress.Checked Then
                vnCrStatus_3_Move = "PM.TransStatus in(" & enuTCPDLK.Baru & "," & enuTCPDLW.Baru & "," & enuTCPDLK.On_Movement & "," & enuTCPDLW.On_Movement & "," & enuTCPDLW.On_Delivery_Movement & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_3_Move = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPDLK.Movement_Done & " or PM.TransStatus>=" & enuTCPDLW.Movement_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_3_Move = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPDLK.Baru
            End If
            vnCrStatus_3_Move = "            and (" & vnCrStatus_3_Move & ")"

            '<---4 Picklist - Picking - Dispatch
            If ChkPM_InProgress.Checked Then
                vnCrStatus_4_PickList_Picking_Dispatch = "PM.TransStatus in(" & enuTCPICK.Baru & "," & enuTCPICK.Prepared & "," & enuTCPICK.On_Picking & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_4_PickList_Picking_Dispatch = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPICK.Picking_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_4_PickList_Picking_Dispatch = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPICK.Baru
            End If
            vnCrStatus_4_PickList_Picking_Dispatch = "            and (" & vnCrStatus_4_PickList_Picking_Dispatch & ")"

            '<---5 Picking
            If ChkPM_InProgress.Checked Then
                vnCrStatus_5_Picking = "PM.TransStatus in(" & enuTCPCKG.None & "," & enuTCPCKG.On_Picking & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_5_Picking = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPCKG.Picking_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_5_Picking = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPCKG.None
            End If
            vnCrStatus_5_Picking = "            and (" & vnCrStatus_5_Picking & ")"

            '<---6 Dispatch
            If ChkPM_InProgress.Checked Then
                vnCrStatus_6_Dispatch = "PM.TransStatus in(" & enuTCDISP.On_Dispatch & "," & enuTCDISP.Dispatch_Done & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_6_Dispatch = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCDISP.Driver_Confirm
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_6_Dispatch = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCDISP.On_Dispatch
            End If
            vnCrStatus_6_Dispatch = "            and (" & vnCrStatus_6_Dispatch & ")"

            '<---7 Dispatch Receive
            If ChkPM_InProgress.Checked Then
                vnCrStatus_7_Dispatch_Receive = "PM.TransStatus in(" & enuTCDISR.On_Dispatch_Receive & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_7_Dispatch_Receive = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCDISR.Dispatch_Receive_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_7_Dispatch_Receive = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCDISR.On_Dispatch_Receive
            End If
            vnCrStatus_7_Dispatch_Receive = "            and (" & vnCrStatus_7_Dispatch_Receive & ")"

            '<---8 Moving Antar Staging Out
            If ChkPM_InProgress.Checked Then
                vnCrStatus_8_Move_Antar_Staging_Out = "PM.TransStatus in(" & enuTCDSGO.Staging_Out_1_Preparation & "," & enuTCDSGO.On_Delivery_To_Staging_Out_2 & "," & enuTCDSGO.Staging_Out_2_On_Receiving & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_8_Move_Antar_Staging_Out = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCDSGO.Staging_Out_2_Receive_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_8_Move_Antar_Staging_Out = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCDSGO.Staging_Out_1_Preparation
            End If
            vnCrStatus_8_Move_Antar_Staging_Out = "            and (" & vnCrStatus_8_Move_Antar_Staging_Out & ")"

            '<---9 Putaway DO Titip
            If ChkPM_InProgress.Checked Then
                vnCrStatus_9_Putaway_DO_Titip = "PM.TransStatus in(" & enuTCPDTW.Baru & "," & enuTCPDTY.Baru & "," & enuTCPDTW.On_Putaway & "," & enuTCPDTY.On_Putaway & "," & enuTCPDTY.On_Delivery_Putaway & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_9_Putaway_DO_Titip = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPDTW.Putaway_Done & " or PM.TransStatus>=" & enuTCPDTY.Putaway_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_9_Putaway_DO_Titip = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPDTW.Baru
            End If
            vnCrStatus_9_Putaway_DO_Titip = "            and (" & vnCrStatus_9_Putaway_DO_Titip & ")"

            '<---10 Putaway Picking yang dibatalkan
            If ChkPM_InProgress.Checked Then
                vnCrStatus_10_Putaway_Void = "PM.TransStatus in(" & enuTCPDTV.Baru & "," & enuTCPDTV.On_Putaway & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_10_Putaway_Void = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPDTV.Putaway_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_10_Putaway_Void = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPDTV.Baru
            End If
            vnCrStatus_10_Putaway_Void = "            and (" & vnCrStatus_10_Putaway_Void & ")"

            '<---11 Putaway Penerimaan Dispatch
            If ChkPM_InProgress.Checked Then
                vnCrStatus_11_Putaway_Dispatch_Receive = "PM.TransStatus in(" & enuTCPDSW.Baru & "," & enuTCPDSY.Baru & "," & enuTCPDSW.On_Putaway & "," & enuTCPDSY.On_Putaway & "," & enuTCPDSY.On_Delivery_Putaway & ")"
            End If
            If ChkPM_Done.Checked Then
                vnCrStatus_11_Putaway_Dispatch_Receive = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus>=" & enuTCPDSY.Putaway_Done & " or PM.TransStatus>=" & enuTCPDSY.Putaway_Done
            End If
            If ChkPM_Batal.Checked Then
                vnCrStatus_11_Putaway_Dispatch_Receive = IIf(vnCrStatus = "", "", " or ") & "PM.TransStatus<" & enuTCPDSW.Baru
            End If
            vnCrStatus_11_Putaway_Dispatch_Receive = "            and (" & vnCrStatus_11_Putaway_Dispatch_Receive & ")"

        End If

        '<---1 Penerimaan Pembelian-------------------------
        vnQuery = "Select PM.vTransOID,PM.vRefOID,SM.TransCode,SM.TransName,PM.vTransCompanyCode,PM.vTransNo,convert(varchar(11),PM.vTransDate,106)vTransDate,PM.vRefNo,"
        vnQuery += vbCrLf & "     PM.vTransType,PM.WarehouseName,PM.vWarehouseName_Dest,PM.vTransNote,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     PM.vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From(" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,PM.RcvPORefOID vRefOID,PM.TransCode,PM.RcvPOCompanyCode vTransCompanyCode,PM.RcvPONo vTransNo,PM.RcvPODate vTransDate,PM.RcvPORefNo vRefNo,"
        vnQuery += vbCrLf & "     RT.RcvPOTypeName vTransType,WM.WarehouseName,''vWarehouseName_Dest,''vTransNote,ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From Sys_SsoRcvPOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOType_MA RT with(nolock) on RT.OID=PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.RcvPOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1 and RcvRefTypeOID<>" & enuRcvType.Opening_Balance
        vnQuery += vbCrLf & vnCrStatus_1_Penerimaan_Pembelian

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvPODate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvPODate <= '" & TxtListEnd.Text & "'"
        End If

        '<---2 Putaway Penerimaan Pembelian-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.vPtwOID vTransOID,PM.RcvPOHOID vRefOID,PM.TransCode,PM.vPtwCompanyCode vTransCompanyCode,PM.vPtwNo vTransNo,PM.vPtwDate vTransDate,PM.RcvPONo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,''vTransNote,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPutaway() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vPtwCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_2_Putaway

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---3 Pindah Lokasi-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.vMoveOID vTransOID,Null vRefOID,PM.TransCode,PM.vMoveCompanyCode vTransCompanyCode,PM.vMoveNo vTransNo,PM.vMoveDate vTransDate,''vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,''vTransNote,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoMovement() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_3_Move

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.vMoveDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.vMoveDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---4 Picklist - Picking - Dispatch-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,PCLRefHOID vRefOID,PM.TransCode,PM.PCLCompanyCode vTransCompanyCode,PM.PCLNo vTransNo,PM.PCLDate vTransDate,PCLRefHNo vRefNo,"
        vnQuery += vbCrLf & "     PM.SchDTypeName vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,PM.vPCLRefHInfo_Html vTransNote,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPicklist_Picking_Dispatch() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_4_PickList_Picking_Dispatch

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---5 Picking-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,PM.PCLRefHOID vRefOID,PM.TransCode,PM.PCKCompanyCode vTransCompanyCode,PM.PCKNo vTransNo,PM.PCKDate vTransDate,PM.PCLRefHNo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,''vWarehouseName_Dest,PM.vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPCKHeader() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_5_Picking

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PCKDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PCKDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---6 Dispatch-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,Null vRefOID,PM.TransCode,PM.DSPCompanyCode vTransCompanyCode,PM.DSPNo vTransNo,PM.DSPDate vTransDate,Null vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,''vWarehouseName_Dest,'Driver = '+dm.DcmDriverName+'<br />Plat No = '+vm.VehicleNo vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoDSPHeader() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmDriver_MA dm with(nolock) on dm.OID=PM.DcmSchDriverOID"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmVehicle_MA vm with(nolock) on vm.OID=PM.DcmVehicleOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_6_Dispatch

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.DSPDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.DSPDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---7 Dispatch Receive-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,DSPHOID vRefOID,PM.TransCode,PM.DSRCompanyCode vTransCompanyCode,PM.DSRNo vTransNo,PM.DSRDate vTransDate,DSPNo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,''vWarehouseName_Dest,'Gudang Asal='+ WA.WarehouseName +'<br />Driver = '+dm.DcmDriverName+'<br />Plat No = '+vm.VehicleNo vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoDSRHeader() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WA with(nolock) on WA.OID=PM.vWarehouseOID_Asal"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmDriver_MA dm with(nolock) on dm.OID=PM.DcmSchDriverOID"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmVehicle_MA vm with(nolock) on vm.OID=PM.DcmVehicleOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_7_Dispatch_Receive

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.DSRDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.DSRDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---8 Moving Antar Staging Out-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.OID vTransOID,Null vRefOID,PM.TransCode,PM.SGOCompanyCode vTransCompanyCode,PM.SGONo vTransNo,PM.SGODate vTransDate,''vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,''vWarehouseName_Dest,'Staging Out Asal ='+ SA.vStorageInfo_Wh_Bd_Lt +'<br />Staging Out Destination = '+ SD.vStorageInfo_Wh_Bd_Lt vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoSGOHeader() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) SA on SA.vStorageOID=PM.StorageOID"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) SD on SD.vStorageOID=PM.StorageOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vMoveCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_8_Move_Antar_Staging_Out

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.SGODate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.SGODate <= '" & TxtListEnd.Text & "'"
        End If

        '<---9 Putaway DO Titip-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.vPtwOID vTransOID,PM.PCLRefHOID vRefOID,PM.TransCode,PM.vPtwCompanyCode vTransCompanyCode,PM.vPtwNo vTransNo,PM.vPtwDate vTransDate,PM.PCLRefHNo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,PM.vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPutaway_DT() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vPtwCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_9_Putaway_DO_Titip

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---10 Putaway Picking yang dibatalkan-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.vPtwOID vTransOID,PM.PCLRefHOID vRefOID,PM.TransCode,PM.vPtwCompanyCode vTransCompanyCode,PM.vPtwNo vTransNo,PM.vPtwDate vTransDate,PM.PCLRefHNo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,PM.vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPutaway_PTV() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vPtwCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_10_Putaway_Void

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        '<---11 Putaway Penerimaan Dispatch-------------------------
        vnQuery += vbCrLf & vbCrLf & "UNION ALL" & vbCrLf

        vnQuery += vbCrLf & "Select PM.vPtwOID vTransOID,PM.DSRHOID vRefOID,PM.TransCode,PM.vPtwCompanyCode vTransCompanyCode,PM.vPtwNo vTransNo,PM.vPtwDate vTransDate,PM.DSRNo vRefNo,"
        vnQuery += vbCrLf & "     ''vTransType,WM.WarehouseName,WD.WarehouseName vWarehouseName_Dest,''vPCLRefHInfo_Html,PM.TransStatusDescr,"
        vnQuery += vbCrLf & "     vCreation,PM.CreationDatetime"
        vnQuery += vbCrLf & "From fnTbl_SsoPutaway_DS() PM"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA WD with(nolock) on WD.OID=PM.WarehouseOID_Dest"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.vPtwCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus_11_Putaway_Dispatch_Receive

        If DstListWarehouse.SelectedValue > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWarehouse.SelectedValue
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & ")PM"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransName_MA SM with(nolock)on SM.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "Order by PM.CreationDatetime,PM.vTransNo"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvPM.DataSource = vnDtb
        GrvPM.DataBind()

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub
End Class