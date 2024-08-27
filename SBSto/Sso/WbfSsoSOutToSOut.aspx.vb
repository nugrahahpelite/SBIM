Imports System.Data.SqlClient
Public Class WbfSsoSOutToSOut
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoSOutToSOut"
    Const csTNoPrefix = "SGO"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String

    Enum ensColList
        OID = 0
    End Enum
    Enum ensSOTSO
        Cancelled = -2
        Preparation = 14
        On_Delivery = 16
        On_Receiving = 18
        Receive_Done = 19
        Closed = 20
    End Enum
    Enum ensColData1
        OID = 0
        RcvPONo = 1
        SGOScan1Qty = 2
        vSGOScan1Note = 3
        vSGOScan1User = 4
        vSGOScan1Time = 5
        vDelItem1 = 6
        vSGOScan1Deleted = 7
        SGOScan1DeletedNote = 8
        vSGOScan1DeletedTime = 9
    End Enum

    Enum ensColData2
        OID = 0
        vStorageInfoHtml = 1
        RcvPONo = 2
        PWScan2Qty = 3
        vPWScan2Note = 4
        vPWScan2User = 5
        vPWScan2Time = 6
        vDelItem2 = 7
        vSGOScan2Deleted = 8
        SGOScan2DeletedNote = 9
        vSGOScan2DeletedTime = 10
    End Enum

    Enum ensColSumm
        vPCKHOID = 0
        PCKNo = 1
        PCKDate = 2
        PCLRefHNo = 3
        SchDTypeName = 4
        vIsQtyConfirm = 5
        vIsQtyConfirm_Dest = 6
        PCLNo = 7
        vDelItem = 8
    End Enum


    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtStorageOID.Text = ""
        TxtvStgOut_Dest_InfoComplete.Text = ""
        TxtTransNo.Text = ""
        TxtTransWhsName.Text = ""
        TxtvStgOut_InfoComplete.Text = ""
        TxtCompany.Text = ""

        HdfTransStatus.Value = enuTCSSOH.Baru
    End Sub
    Enum ensColLsScan
        vRcvPOScanDeleted = 5
    End Enum
    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Sso"
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPutaway_Antar_Wh, vnSQLConn)

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        Dim vnCrList As String = fbuFormatString(Trim(TxtListFind.Text))

        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnCrStatus As String = ""
        If ChkSt_Preparation.Checked = False And ChkSt_OnDelivery.Checked = False And ChkSt_OnReceiving.Checked = False And ChkSt_Done.Checked = False And ChkSt_Closed.Checked = False And ChkSt_Cancelled.Checked = False Then
            ChkSt_OnDelivery.Checked = True
            ChkSt_Preparation.Checked = True
        End If

        If ChkSt_Preparation.Checked = True Then
            vnCrStatus += enuTCDSGO.Staging_Out_1_Preparation & ","
        End If
        If ChkSt_OnDelivery.Checked = True Then
            vnCrStatus += enuTCDSGO.On_Delivery_To_Staging_Out_2 & ","
        End If
        If ChkSt_OnReceiving.Checked = True Then
            vnCrStatus += enuTCDSGO.Staging_Out_2_On_Receiving & ","
        End If
        If ChkSt_Done.Checked = True Then
            vnCrStatus += enuTCDSGO.Staging_Out_2_Receive_Done & ","
        End If
        If ChkSt_Closed.Checked = True Then
            vnCrStatus += ensSOTSO.Closed & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += ensSOTSO.Cancelled & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and stn.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	Select DISTINCT pwh.OID,pwh.SGOCompanyCode,mc.CompanyName,pwh.SGONo,convert(varchar(11),pwh.SGODate,106)vSGODate,pwh.WarehouseOID, st1.WarehouseName,"
        vnQuery += vbCrLf & "	   pwh.StorageOID,st1.vStorageInfo_Wh_Bd_Lt vStgOut,st1.vStorageInfo,st1.vStorageInfo_Complete vStgOut_InfoComplete,pwh.StorageOID_Dest,st2.vStorageInfo_Wh_Bd_Lt vStgOut_Dest,"
        vnQuery += vbCrLf & "	   st2.vStorageInfo_Complete vStgOut_Dest_InfoComplete,pwh.SGOCancelNote,stn.TransStatusDescr"
        vnQuery += vbCrLf & "	  From Sys_SsoSGOHeader_TR pwh with(nolock)"
        vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) st1 on st1.vStorageOID=pwh.StorageOID"
        vnQuery += vbCrLf & "	       left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) st2 on st2.vStorageOID=pwh.StorageOID_Dest"
        vnQuery += vbCrLf & "	       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus"
        vnQuery += vbCrLf & "	       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.SGOCompanyCode"
        vnQuery += vbCrLf & "	       INNER JOIN " & vnDBMaster & "DimCompany mc  with(nolock) on mc.CompanyCode =pwh.SGOCompanyCode"
        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        vnQuery += vbCrLf & "and usc.UserOID=" & Session("UserOID")

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and pwh.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and ( pwh.SGONo like '%" & Trim(TxtListNo.Text) & "%' )"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and pwh.SGODate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and pwh.SGODate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and ( pwh.WarehouseOID = " & DstListWhs.SelectedValue & " ) "
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        'psShowBrg(False)
        'psShowSumm(False)
        psShowList(True)
        psFillGrvList()
    End Sub

    'Protected Sub BtnListClose_Click(sender As Object, e As EventArgs) Handles BtnListClose.Click
    '    psShowBrg(False)
    '    psShowSumm(False)
    '    psShowList(False)

    'End Sub


    Private Sub psShowList(vriBo As Boolean)
        If vriBo Then
            DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psFillGrvList()
            TxtListFind.Focus()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
        End If
    End Sub



    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        psClearData()
        psClearMessage()
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        If Not IsDate(TxtListStart.Text) Then
            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -14, Date.Now), "dd MMM yyyy")
        End If
        If Not IsDate(TxtListEnd.Text) Then
            TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")
        End If
        GrvBRG.Visible = False
        PanBRG.Visible = False
        GrvSumm.Visible = False
        PanBRG.Visible = False
        tbTrans.Visible = False

        psShowList(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblXlsProses.Text = ""

    End Sub

    Private Sub psButtonStatusDefault()
        BtnList.Enabled = True
        BtnCancelSGO.Enabled = True
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If e.CommandName = "SGONo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)

            HdfTransOID.Value = vnRow.Cells(ensColList.OID).Text

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgConfirm.Text = pbMsgError
                LblMsgConfirm.Visible = True
                pbMsgError = ""
                Exit Sub
            End If

            psDisplayData(HdfTransOID.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
            psShowList(False)

            tbTrans.Visible = True
        End If
    End Sub

    Protected Sub GrvSumm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSumm.SelectedIndexChanged

    End Sub

    Private Sub GrvSumm_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumm.RowCommand
        If e.CommandName = "PCKNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvSumm.Rows(vnIdx)
            Dim vnpckno As String = vnRow.Cells(1).Text
            Dim vnsgono As String = TxtTransNo.Text

            HdfTransOID2.Value = vnRow.Cells(ensColSumm.vPCKHOID).Text

            HdfPCKOID.Value = vnRow.Cells(ensColSumm.vPCKHOID).Text

            HdfSGONo2.Value = TxtTransNo.Text

            TxtPCKOID.Text = vnRow.Cells(ensColSumm.vPCKHOID).Text
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgConfirm.Text = pbMsgError
                LblMsgConfirm.Visible = True
                pbMsgError = ""
                Exit Sub
            End If
            PanBRG.Visible = True
            GrvBRG.Visible = True

            'psDisplayBRG(HdfTransOID2.Value, vnSQLConn)
            psFillGrvBRG(HdfSGONo2.Value, HdfTransOID2.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
            psShowList(False)
        End If

    End Sub

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriSGOHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,''PCKNo,''vPCKDate,''PCLNo,''PCLRefHNo,''SchDTypeName,''vIsQtyConfirm,''vIsQtyConfirm_Dest,''vDelItem Where 1=2"
        Else
            Dim vnCrList As String = fbuFormatString(Trim(TxtTransID.Text))
            Dim vnQuery1 As String
            vnQuery1 = "Select pwh.OID vPCKHOID,pwh.PCKNo,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
            vnQuery1 += vbCrLf & "       pch.PCLNo,pch.PCLRefHNo,msc.SchDTypeName,"
            vnQuery1 += vbCrLf & "       case when abs(pcp.IsQtyConfirm)=1 then 'Confirm' else 'Not Confirm' end vIsQtyConfirm,"
            vnQuery1 += vbCrLf & "       case when abs(pcp.IsQtyConfirm_Dest)=1 then 'Confirm' else 'Not Confirm' end vIsQtyConfirm_Dest,"

            If HdfTransStatus.Value = enuTCDSGO.Staging_Out_1_Preparation Then
                vnQuery1 += vbCrLf & "       'Hapus'vDelItem"
            Else
                vnQuery1 += vbCrLf & "       ''vDelItem"
            End If

            vnQuery1 += vbCrLf & "  From Sys_SsoPCKHeader_TR pwh with(nolock)"
            vnQuery1 += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
            vnQuery1 += vbCrLf & "       inner join " & fbuGetDBDcm() & "Sys_DcmSchDType_MA msc with(nolock) on msc.OID=pch.SchDTypeOID"
            vnQuery1 += vbCrLf & "       inner join Sys_SsoSGOPick_TR pcp with(nolock) on pcp.PCKHOID=pwh.OID"
            vnQuery1 += vbCrLf & " Where pwh.StorageOID=" & TxtStorageOID.Text & " and pcp.SGOHOID=" & TxtTransID.Text

            If HdfTransStatus.Value > enuTCDSGO.Staging_Out_1_Preparation Then
                vnQuery = vnQuery1 & vbCrLf & "     Order by pch.PCLNo"
            Else
                vnQuery = "Select * From ("
                vnQuery += vbCrLf & vnQuery1

                vnQuery += vbCrLf & "UNION"

                vnQuery += vbCrLf & "Select pwh.OID,pwh.PCKNo,convert(varchar(11),pwh.PCKDate,106)vPCKDate,"
                vnQuery += vbCrLf & "       pch.PCLNo,pch.PCLRefHNo,msc.SchDTypeName,''vIsQtyConfirm,''vIsQtyConfirm_Dest,''vDelItem"
                vnQuery += vbCrLf & "  From Sys_SsoPCKHeader_TR pwh with(nolock)"
                vnQuery += vbCrLf & "       inner join Sys_SsoPCLHeader_TR pch with(nolock) on pch.OID=pwh.PCLHOID"
                vnQuery += vbCrLf & "       inner join " & fbuGetDBDcm() & "Sys_DcmSchDType_MA msc with(nolock) on msc.OID=pch.SchDTypeOID"
                vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.PCKCompanyCode"
                vnQuery += vbCrLf & " Where pwh.StorageOID=" & TxtStorageOID.Text & " and isnull(pwh.StorageOID_Current,0)=0"
                vnQuery += vbCrLf & "       and pwh.TransStatus=" & enuTCPCKG.Picking_Done & " and usc.UserOID=" & Session("UserOID") & " and pwh.WarehouseOID='" & Session("LoginWhsOID") & "' and pwh.StorageOID=" & TxtStorageOID.Text
                vnQuery += vbCrLf & "       and pwh.PCKCompanyCode='" & HdfCompanyCode.Value & "'"
                vnQuery += vbCrLf & "       and (pwh.PCKNo like '%" & vnCrList & "%' or pch.PCLNo like '%" & vnCrList & "%' or pch.PCLRefHNo like '%" & vnCrList & "%')"
                vnQuery += vbCrLf & ")tb"
                vnQuery += vbCrLf & "Order by case when vIsQtyConfirm='' then 5 else 4 end,PCLNo"
            End If
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
        GrvSumm.Visible = True
        tbTrans.Visible = True
    End Sub


    Private Sub psDisplayData(vriOID As String, vriSQLConn As SqlClient.SqlConnection)
        psClearMessage()
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnQuery As String
        Dim vnDtb As New DataTable
        vnQuery = "	Select DISTINCT pwh.OID,pwh.SGOCompanyCode,mc.CompanyName,pwh.SGONo,convert(varchar(11),pwh.SGODate,106)vSGODate,pwh.WarehouseOID, st1.WarehouseName, 	"
        vnQuery += vbCrLf & "	   pwh.StorageOID,st1.vStorageInfo_Wh_Bd_Lt vStgOut,st1.vStorageInfo,st1.vStorageInfo_Complete vStgOut_InfoComplete, stn.TransCode, stn.TransStatus, stn.TransStatusDescr, 	"
        vnQuery += vbCrLf & "	   st2.vStorageInfo_Complete vStgOut_Dest_InfoComplete, pwh.StorageOID_Dest,st2.vStorageInfo_Wh_Bd_Lt vStgOut_Dest	"
        vnQuery += vbCrLf & "	  From Sys_SsoSGOHeader_TR pwh with(nolock)	"
        vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) st1 on st1.vStorageOID=pwh.StorageOID	"
        vnQuery += vbCrLf & "	       left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) st2 on st2.vStorageOID=pwh.StorageOID_Dest	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoTransStatus_MA stn with(nolock) on stn.TransCode=pwh.TransCode and stn.TransStatus=pwh.TransStatus	"
        vnQuery += vbCrLf & "	       inner join Sys_SsoUserCompany_MA usc with(nolock) on usc.CompanyCode=pwh.SGOCompanyCode	"
        vnQuery += vbCrLf & "	       INNER JOIN " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode =pwh.SGOCompanyCode	"
        vnQuery += vbCrLf & "Where pwh.OID=" & vriOID
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            HdfCompanyCode.Value = ""
            psClearData()

        Else
            Dim vnDRow As DataRow = vnDtb.Rows(0)


            HdfTransOID.Value = vriOID

            TxtTransID.Text = vnDtb.Rows(0).Item("OID")

            TxtListFind.Text = vnDtb.Rows(0).Item("OID")

            TxtTransDate.Text = vnDtb.Rows(0).Item("vSGODate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("SGONo")

            TxtStorageOID.Text = vnDtb.Rows(0).Item("StorageOID")
            TxtTransWhsName.Text = vnDtb.Rows(0).Item("WarehouseName")

            HdfCompanyCode.Value = vnDtb.Rows(0).Item("SGOCompanyCode")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")
            HdfWarehouseOID.Value = vnDtb.Rows(0).Item("WarehouseOID")

            HdfStorageOID.Value = vnDtb.Rows(0).Item("StorageOID")

            TxtCompany.Text = vnDtb.Rows(0).Item("SGOCompanyCode")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")
            TxtvStgOut_InfoComplete.Text = vnDtb.Rows(0).Item("vStgOut_InfoComplete")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")
            If (IsDBNull(vnDtb.Rows(0).Item("StorageOID_Dest")) = True) Then
                HdfStorageOID_Dest.Value = ""
                TxtvStgOut_Dest_InfoComplete.Text = ""
            Else
                HdfStorageOID_Dest.Value = vnDtb.Rows(0).Item("StorageOID_Dest")

                TxtvStgOut_Dest_InfoComplete.Text = vnDtb.Rows(0).Item("vStgOut_Dest_InfoComplete")
            End If

            psButtonStatus()
        End If

        psFillGrvSumm(0, Val(TxtTransID.Text), vriSQLConn)
        GrvSumm.Visible = True
        vnDtb.Dispose()

    End Sub

    Private Sub psFillGrvBRG(vriSGONo As String, vriPCKOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery1 As String
        Dim vnDtb As New DataTable
        Dim vnCrList As String = fbuFormatString(Trim(TxtTransID.Text))
        vnQuery1 = "	Select isnull(dps.OID,0)vSGOOID,sc.BrgCode,mbr.BrgName,rcv.RcvPONo,sc.RcvPOHOID,vSumPCKScanQty,"
        vnQuery1 += vbCrLf & "	isnull(SGOScanQty,0)SGOScanQty,isnull(SGOScanQty_Dest,0)SGOScanQty_Dest,"
        vnQuery1 += vbCrLf & "	 ''vConfirm,"
        vnQuery1 += vbCrLf & "	''vNotConfirm"
        vnQuery1 += vbCrLf & "	 FROM Sys_SsoPCKHeader_TR pck with(nolock)"
        vnQuery1 += vbCrLf & "	      inner join Sys_SsoSGOPick_TR sgopick with(nolock) ON sgopick.PCKHOID = pck.OID"
        vnQuery1 += vbCrLf & "	      inner join Sys_SsoSGOHeader_TR pwh with(nolock) ON sgopick.SGOHOID = pwh.OID AND pck.StorageOID = pwh.StorageOID"
        vnQuery1 += vbCrLf & "	      inner join (Select isc.BrgCode,isc.RcvPOHOID,ISC.PCKHOID,sum(isc.PCKScanQty)vSumPCKScanQty From Sys_SsoPCKScan_TR isc with(nolock) where abs(isc.PCKScanDeleted)=0 group by isc.BrgCode,isc.RcvPOHOID,ISC.PCKHOID)sc	ON sc.PCKHOID = pck.OID	"
        vnQuery1 += vbCrLf & "	      inner join Sys_SsoRcvPOHeader_TR rcv with(nolock)on rcv.OID=sc.RcvPOHOID"
        vnQuery1 += vbCrLf & "	      inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA mbr with(nolock)on mbr.BrgCode=sc.BrgCode and mbr.CompanyCode=rcv.RcvPOCompanyCode"
        vnQuery1 += vbCrLf & "	      left outer join (Select idps.OID,idps.BrgCode,idps.RcvPOHOID,idps.SGOHOID,idps.PCKHOID, SGOScanQty,SGOScanQty_Dest From Sys_SsoSGOScan_TR idps with(nolock))dps on dps.BRGCODE=sc.BRGCODE and dps.RcvPOHOID=sc.RcvPOHOID	AND dps.SGOHOID = pwh.OID AND dps.PCKHOID = pck.OID	"
        vnQuery1 += vbCrLf & "  Where 1=1 and pwh.SGONo = '" & Trim(HdfSGONo2.Value) & "'"
        vnQuery1 += vbCrLf & "	      and pck.OID = " & Trim(TxtPCKOID.Text) & " "
        vnQuery1 += vbCrLf & "	 Order by sc.BrgCode,rcv.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery1, vriSQLConn)
        GrvBRG.DataSource = vnDtb
        GrvBRG.DataBind()
    End Sub
    Private Sub psShowBrg(vriBo As Boolean)
        If vriBo Then
            GrvBRG.Style(HtmlTextWriterStyle.Visibility) = "visible"
            PanBRG.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TbBRG.Style(HtmlTextWriterStyle.Visibility) = "visible"
            PanBRG.Visible = True
            GrvBRG.Visible = True
            TbBRG.Visible = True
        Else
            GrvBRG.Style(HtmlTextWriterStyle.Visibility) = "hidden"

            PanBRG.Style(HtmlTextWriterStyle.Visibility) = "hidden"

            TbBRG.Style(HtmlTextWriterStyle.Visibility) = "visible"
            PanBRG.Visible = False
            GrvBRG.Visible = False
            TbBRG.Visible = False
        End If
    End Sub

    Private Sub psShowSumm(vriBo As Boolean)
        If vriBo Then
            GrvSumm.Style(HtmlTextWriterStyle.Visibility) = "visible"

            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            GrvSumm.Style(HtmlTextWriterStyle.Visibility) = "hidden"

            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psShowBrg(False)
        End If
    End Sub

    Protected Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub
    Private Sub psShowConfirm(vriBo As Boolean)
        If vriBo Then
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "visible"
            BtnConfirmYes.Visible = True
            BtnConfirmNo.Text = "NO"
        Else
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Protected Sub BtnCancelDSR_Click(sender As Object, e As EventArgs) Handles BtnCancelSGO.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        HdfProcessDataKey.Value = Format(Date.Now(), "yyyyMMdd_HHmmss_") & Session("UserOID") & "_" & csModuleName

        LblConfirmMessage.Text = "Anda Membatalkan Moving Antar Staging Out No. " & TxtTransID.Text & " ?<br />WARNING : Batal Moving Antar Staging Out Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelSGO"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelSGO" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelSGO()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnCancelSGO.Enabled = (HdfTransStatus.Value = enuTCDSGO.Staging_Out_1_Preparation)
            psButtonVisible()
        End If
    End Sub
    Private Sub psButtonVisible()
        BtnCancelSGO.Visible = BtnCancelSGO.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub

    Private Sub psCancelSGO()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True

            vsTextStream.WriteLine("Open SQL Connection Error....")
            vsTextStream.WriteLine(pbMsgError)
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnSGOHOID As String = TxtTransID.Text
            Dim vnTransStatus As Integer
            Dim vnCount1 As Integer
            Dim vnQuery As String
            vnQuery = "Select TransStatus From Sys_SsoSGOHeader_TR with(nolock) Where OID=" & vnSGOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnTransStatus = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnTransStatus <> enuTCDSGO.Staging_Out_1_Preparation Then
                LblMsgError.Text = "Status Sudah Moving"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                psDisplayData(vnSGOHOID, vnSQLConn)
                Exit Sub
            End If

            vnQuery = "Select count(OID) FROM Sys_SsoSGOPick_TR Where SGOHOID=" & vnSGOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0.1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnCount1 = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            If vnCount1 > 0 Then
                LblMsgError.Text = "Sudah ada picklist dipilih"
                LblMsgError.Visible = True

                vsTextStream.WriteLine(LblMsgError.Text)
                vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("---------------EOF-------------------------")
                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing

                psDisplayData(vnSGOHOID, vnSQLConn)
                Exit Sub
            End If

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            pbuSsoProcessDataKey(HdfProcessDataKey.Value, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoSGOHeader_TR set TransStatus=" & enuTCDISR.Cancelled & ",SGOCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & vnSGOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("pbuInsertStatusSGO...Start")
            pbuInsertStatusSGO(vnSGOHOID, enuTCDISR.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSGO...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psDisplayData(vnSGOHOID, vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

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
End Class