Imports System.Data.SqlClient
Public Class WbfSsoMonTrans
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

        psFillGrvGR(vnSQLConn)
        psFillGrvPtw(vnSQLConn)
        psFillGrvPck(vnSQLConn)
        psFillGrvMove(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvGR(vriSQLConn As SqlConnection)

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkGR_OnReceive.Checked = False And ChkGR_ReceiveDone.Checked = False And ChkGR_ReceiveApp.Checked = False And ChkGR_PutawayProcess.Checked = False And ChkGR_AllPutawayComplete.Checked = False Then
            ChkGR_OnReceive.Checked = True
            ChkGR_ReceiveDone.Checked = True
            ChkGR_ReceiveApp.Checked = True
            ChkGR_PutawayProcess.Checked = True
            ChkGR_AllPutawayComplete.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkGR_OnReceive.Checked = True Then
            vnCrStatus += enuTCRCPO.On_Receive & ","
        End If
        If ChkGR_ReceiveDone.Checked = True Then
            vnCrStatus += enuTCRCPO.Receive_Done & ","
        End If
        If ChkGR_ReceiveApp.Checked = True Then
            vnCrStatus += enuTCRCPO.Receive_Approved & ","
        End If
        If ChkGR_PutawayProcess.Checked = True Then
            vnCrStatus += enuTCRCPO.Putaway_Process & ","
        End If
        If ChkGR_AllPutawayComplete.Checked = True Then
            vnCrStatus += enuTCRCPO.All_Putaway_Clomplete & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,MT.RcvTypeName,PM.RcvPONo,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "     PM.RcvPORefNo,RT.RcvPOTypeName,"
        vnQuery += vbCrLf & "     PM.RcvPOCompanyCode,WM.WarehouseName,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ReceivedDoneDatetime,106)+' '+convert(varchar(5),PM.ReceivedDoneDatetime,108)+' '+ RD.UserName vReceivedDone,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ReceivedAppDatetime,106)+' '+convert(varchar(5),PM.ReceivedAppDatetime,108)+' '+ AP.UserName vReceivedApp"
        vnQuery += vbCrLf & "From Sys_SsoRcvPOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoRcvType_MA MT with(nolock) on MT.OID=PM.RcvRefTypeOID"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOType_MA RT with(nolock) on RT.OID=PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA RD with(nolock) on RD.OID=PM.ReceivedDoneUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA AP with(nolock) on AP.OID=PM.ReceivedAppUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.RcvPOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvPODate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.RcvPODate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvGR.DataSource = vnDtb
        GrvGR.DataBind()
    End Sub

    Private Sub psFillGrvPtw(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserOID As String = Session("UserOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkPtw_Process.Checked = False And ChkPtw_Done.Checked = False Then
            ChkPtw_Process.Checked = True
            ChkPtw_Done.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkPtw_Process.Checked = True Then
            vnCrStatus += enuTCPWAY.Baru & "," & enuTCPWAY.On_Putaway & "," & enuTCPYAY.Baru & "," & enuTCPYAY.On_Delivery_Putaway & "," & enuTCPYAY.On_Putaway & ","
        End If
        If ChkPtw_Done.Checked = True Then
            vnCrStatus += enuTCPWAY.Putaway_Done & "," & enuTCPYAY.Putaway_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnWarehouseOID As String = DstListWarehouse.SelectedValue
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select TransCode,TransName,vPtwOID,vPtwCompanyCode,vPtwNo,RcvPONo,convert(varchar(11),vPtwDate,106)vPtwDate,"
        vnQuery += vbCrLf & "       WarehouseName,vWarehouseName_Dest,TransStatusDescr,"
        vnQuery += vbCrLf & "       vCreation,vOnDelivery,vOnPutaway,vPutawayDone"
        vnQuery += vbCrLf & "From("
        vnQuery += vbCrLf & "Select trn.TransCode,trn.TransName,pwh.OID vPtwOID,pwh.PWCompanyCode vPtwCompanyCode,pwh.PWNo vPtwNo,rcv.RcvPONo,PWDate vPtwDate,"
        vnQuery += vbCrLf & "       str.WarehouseName,''vWarehouseName_Dest,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       ''vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnPutawayDatetime,106)+' '+convert(varchar(5),pwh.OnPutawayDatetime,108)+' '+ OP.UserName vOnPutaway,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.PutawayDoneDatetime,106)+' '+convert(varchar(5),pwh.PutawayDoneDatetime,108)+' '+ PD.UserName vPutawayDone"
        vnQuery += vbCrLf & "  From Sys_SsoPWHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransName_MA trn with(nolock)on trn.TransCode=pwh.TransCode"
        vnQuery += vbCrLf & "       inner join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=pwh.RcvPOHOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=rcv.StorageOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnPutawayUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.PutawayDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PWCompanyCode and uc.UserOID=" & vnUserOID
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & vnUserOID
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & vnWarehouseOID
        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PWDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PWDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "UNION ALL"

        vnQuery += vbCrLf & "Select trn.TransCode,trn.TransName,pwh.OID,pwh.PYCompanyCode,pwh.PYNo,rcv.RcvPONo,pwh.PYDate,"
        vnQuery += vbCrLf & "       str.WarehouseName,whd.WarehouseName vWarehouseName_Dest,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnDeliveryPtwDatetime,106)+' '+convert(varchar(5),pwh.OnDeliveryPtwDatetime,108)+' '+ OD.UserName vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnPutawayDatetime,106)+' '+convert(varchar(5),pwh.OnPutawayDatetime,108)+' '+ OP.UserName vOnPutaway,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.PutawayDoneDatetime,106)+' '+convert(varchar(5),pwh.PutawayDoneDatetime,108)+' '+ PD.UserName vPutawayDone"
        vnQuery += vbCrLf & "  From Sys_SsoPYHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransName_MA trn with(nolock)on trn.TransCode=pwh.TransCode"
        vnQuery += vbCrLf & "       inner join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=pwh.RcvPOHOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=rcv.StorageOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA whd on whd.OID=pwh.WarehouseOID_Dest"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OD with(nolock) on OD.OID=pwh.OnDeliveryPtwUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnPutawayUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.PutawayDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PYCompanyCode and uc.UserOID=" & vnUserOID
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & vnUserOID
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and (pwh.WarehouseOID = " & vnWarehouseOID & " or pwh.WarehouseOID_Dest = " & vnWarehouseOID & ")"

        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PYDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PYDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "UNION ALL"

        vnQuery += vbCrLf & "Select trn.TransCode,trn.TransName,pwh.OID,pwh.PTKCompanyCode,pwh.PTKNo,rcv.RcvPONo,PTKDate,"
        vnQuery += vbCrLf & "       str.WarehouseName,''vWarehouseName_Dest,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnPutawayDatetime,106)+' '+convert(varchar(5),pwh.OnPutawayDatetime,108)+' '+ OP.UserName vOnPutaway,"
        vnQuery += vbCrLf & "       ''vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.PutawayDoneDatetime,106)+' '+convert(varchar(5),pwh.PutawayDoneDatetime,108)+' '+ PD.UserName vPutawayDone"
        vnQuery += vbCrLf & "  From Sys_SsoPTKHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransName_MA trn with(nolock)on trn.TransCode=pwh.TransCode"
        vnQuery += vbCrLf & "       inner join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=pwh.RcvPOHOID"
        vnQuery += vbCrLf & "       inner join " & fbuGetDBMaster() & "fnTbl_SsoStorageInfo('') str on str.vStorageOID=rcv.StorageOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnPutawayUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.PutawayDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PTKCompanyCode and uc.UserOID=" & vnUserOID
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & vnUserOID
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & vnWarehouseOID

        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PTKDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PTKDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & ")vTb"
        vnQuery += vbCrLf & "Order by vPtwDate,vPtwNo"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvPtw.DataSource = vnDtb
        GrvPtw.DataBind()
    End Sub

    Private Sub psFillGrvMove(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserOID As String = Session("UserOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkMove_Process.Checked = False And ChkMove_Done.Checked = False Then
            ChkMove_Process.Checked = True
            ChkMove_Done.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkMove_Process.Checked = True Then
            vnCrStatus += enuTCPDLK.Baru & "," & enuTCPDLK.On_Movement & "," & enuTCPDLW.Baru & "," & enuTCPDLW.On_Delivery_Movement & "," & enuTCPDLW.On_Movement & ","
        End If
        If ChkMove_Done.Checked = True Then
            vnCrStatus += enuTCPDLK.Movement_Done & "," & enuTCPDLW.Movement_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and pwh.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnWarehouseOID As String = DstListWarehouse.SelectedValue
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select TransCode,TransName,vMoveOID,vMoveCompanyCode,vMoveNo,convert(varchar(11),vMoveDate,106)vMoveDate,"
        vnQuery += vbCrLf & "       WarehouseName,vWarehouseName_Dest,TransStatusDescr,"
        vnQuery += vbCrLf & "       vCreation,vOnDelivery,vOnMovement,vMovementDone"
        vnQuery += vbCrLf & "From("

        vnQuery += vbCrLf & "Select trn.TransCode,trn.TransName,pwh.OID vMoveOID,pwh.PDLCompanyCode vMoveCompanyCode,pwh.PDLNo vMoveNo,PDLDate vMoveDate,"
        vnQuery += vbCrLf & "       whs.WarehouseName,''vWarehouseName_Dest,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       ''vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnMovementDatetime,106)+' '+convert(varchar(5),pwh.OnMovementDatetime,108)+' '+ OP.UserName vOnMovement,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.MovementDoneDatetime,106)+' '+convert(varchar(5),pwh.MovementDoneDatetime,108)+' '+ PD.UserName vMovementDone"
        vnQuery += vbCrLf & "  From Sys_SsoPDLHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransName_MA trn with(nolock)on trn.TransCode=pwh.TransCode"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnMovementUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.MovementDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PDLCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and pwh.WarehouseOID = " & vnWarehouseOID
        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDLDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDLDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "UNION ALL"

        vnQuery += vbCrLf & "Select trn.TransCode,trn.TransName,pwh.OID,pwh.PDWCompanyCode,pwh.PDWNo,PDWDate,"
        vnQuery += vbCrLf & "       whs.WarehouseName,whd.WarehouseName vWarehouseName_Dest,stn.TransStatusDescr,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.CreationDatetime,106)+' '+convert(varchar(5),pwh.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnDeliveryMovementDatetime,106)+' '+convert(varchar(5),pwh.OnDeliveryMovementDatetime,108)+' '+ OD.UserName vOnDelivery,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.OnMovementDatetime,106)+' '+convert(varchar(5),pwh.OnMovementDatetime,108)+' '+ OP.UserName vOnMovement,"
        vnQuery += vbCrLf & "       convert(varchar(11),pwh.MovementDoneDatetime,106)+' '+convert(varchar(5),pwh.MovementDoneDatetime,108)+' '+ PD.UserName vMovementDone"
        vnQuery += vbCrLf & "  From Sys_SsoPDWHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransName_MA trn with(nolock)on trn.TransCode=pwh.TransCode"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID=pwh.WarehouseOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=pwh.WarehouseOID_Dest"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=pwh.CreationUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OD with(nolock) on OD.OID=pwh.OnDeliveryMovementUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA OP with(nolock) on OP.OID=pwh.OnMovementUserOID"
        vnQuery += vbCrLf & "       left outer join Sys_SsoUser_MA PD with(nolock) on PD.OID=pwh.MovementDoneUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=pwh.PDWCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=pwh.WarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and (pwh.WarehouseOID = " & vnWarehouseOID & " or pwh.WarehouseOID_Dest = " & vnWarehouseOID & ")"
        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDWDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.PDWDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & ")vTb"
        vnQuery += vbCrLf & "Order by vMoveNo,vMoveDate"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvMove.DataSource = vnDtb
        GrvMove.DataBind()
    End Sub

    Private Sub psFillGrvPck(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserOID As String = Session("UserOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkMove_Process.Checked = False And ChkMove_Done.Checked = False Then
            ChkMove_Process.Checked = True
            ChkMove_Done.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkPck_Baru.Checked = True Then
            vnCrStatus += enuTCPICK.Baru & ","
        End If
        If ChkPck_Prepared.Checked = True Then
            vnCrStatus += enuTCPICK.Prepared & ","
        End If
        If ChkPck_OnPicking.Checked = True Then
            vnCrStatus += enuTCPICK.On_Picking & ","
        End If
        If ChkPck_PickingDone.Checked = True Then
            vnCrStatus += enuTCPICK.Picking_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnWarehouseOID As String = DstListWarehouse.SelectedValue
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.PCLNo,convert(varchar(11),PM.PCLDate,106)vPCLDate,convert(varchar(11),PM.PCLScheduleDate,106)vPCLScheduleDate,"
        vnQuery += vbCrLf & "     PM.PCLCompanyCode,TP.SchDTypeName,PM.PCLRefHOID,PM.PCLRefHNo,replace(PM.PCLRefHInfo,char(10),'<br />')vPCLRefHInfo,"
        vnQuery += vbCrLf & "     pck.PCKNo,pdp.vDspPtwNo,WM.WarehouseName,whd.WarehouseName vWarehouseName_Dest,PM.PCLNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr vStatusPickList,"
        vnQuery += vbCrLf & "     stp.TransStatusDescr vStatusPicking,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared"

        vnQuery += vbCrLf & "From Sys_SsoPCLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=PM.WarehouseOID_Dest"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_DcmSchDType_MA TP with(nolock) on TP.OID=PM.SchDTypeOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     left outer join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=PM.OID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"
        vnQuery += vbCrLf & "     left outer join fnTbl_SsoPicking_Dispatch_Putaway() pdp on pdp.PCKHOID=pck.OID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA  stp with(nolock) on stp.TransStatus=pck.TransStatus and stp.TransCode=pck.TransCode"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and PM.WarehouseOID = " & vnWarehouseOID
        vnQuery += vbCrLf & vnCrStatus

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLScheduleDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PCLScheduleDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.PCLNo"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvPck.DataSource = vnDtb
        GrvPck.DataBind()
    End Sub

End Class