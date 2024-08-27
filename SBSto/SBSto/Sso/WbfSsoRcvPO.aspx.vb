Imports System.Data.SqlClient
Public Class WbfSsoRcvPO
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoRcvPO"
    Const csTNoPrefix = "GR"

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

    Enum ensColSumm
        POHOID = 0
        PO_NO = 1
        BRG = 2
        NAMA_BARANG = 3
        vSumPLQty = 4
        vSumPOQty = 5
        vSumRetDRealQty = 6
        vSumRcvPOScanQty = 7
        vRcvPOQty_Total = 8
        vQtyVarian_Import = 9
        vQtyVarian_Local = 10
    End Enum
    Enum ensColData
        OID = 0
        RcvPOScanQty = 1
        vRcvPOScanNote = 2
        vRcvPOScanUser = 3
        vRcvPOScanTime = 4
        vDelItem = 5
        vRcvPOScanDeleted = 6
        vRcvPOScanDeletedUser = 7
        RcvPOScanDeletedNote = 8
        vRcvPOScanDeletedTime = 9
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtTransDate.Text = ""
        TxtTransRefNo.Text = ""
        TxtTransNo.Text = ""
        TxtCompany.Text = ""
        TxtWarehouse.Text = ""

        HdfTransStatus.Value = "0"
        HdfRcvPORefOID.Value = "0"
        HdfRcvType.Value = "0"
        HdfTransStatus.Value = enuTCSSOH.Baru
    End Sub
    Enum ensColLsScan
        vRcvPOScanDeleted = 5
    End Enum
    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPenerimaanPembelian, vnSQLConn)

            'If Session("UserCompanyCode") = "" Then
            '    pbuFillDstCompany(DstCompany, False, vnSQLConn)
            'Else
            '    pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            'End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, False, vnSQLConn)
            pbuFillDstRcvType(DstListRcvType, True, vnSQLConn)
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        If ChkSt_OnReceive.Checked = False And ChkSt_ReceiveDone.Checked = False And ChkSt_ReceiveApproved.Checked = False And ChkSt_ReceivePtwProcess.Checked = False And ChkSt_ReceivePtwComplete.Checked = False Then
            ChkSt_OnReceive.Checked = True
            ChkSt_ReceiveDone.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_OnReceive.Checked = True Then
            vnCrStatus += enuTCRCPO.On_Receive & ","
        End If
        If ChkSt_ReceiveDone.Checked = True Then
            vnCrStatus += enuTCRCPO.Receive_Done & ","
        End If
        If ChkSt_ReceiveApproved.Checked = True Then
            vnCrStatus += enuTCRCPO.Receive_Approved & ","
        End If
        If ChkSt_ReceivePtwProcess.Checked = True Then
            vnCrStatus += enuTCRCPO.Putaway_Process & ","
        End If
        If ChkSt_ReceivePtwComplete.Checked = True Then
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
        vnQuery += vbCrLf & "     PM.RcvPOCompanyCode,WM.WarehouseName,PM.RcvPODoneNote,"
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

        vnQuery += vbCrLf & "Where 1=1 and PM.RcvRefTypeOID in(" & enuRcvType.Pembelian & "," & enuRcvType.Retur & "," & enuRcvType.Lain_lain & "," & enuRcvType.Karantina & ")"
        If DstListRcvType.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "      and PM.RcvRefTypeOID=" & DstListRcvType.SelectedValue
        End If
        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.RcvPONo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and RcvPODate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and RcvPODate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and WarehouseOID = " & DstListWhs.SelectedValue
        End If

        vnQuery += vbCrLf & "Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        psFillGrvList()
    End Sub

    Protected Sub BtnListClose_Click(sender As Object, e As EventArgs) Handles BtnListClose.Click
        psShowList(False)
    End Sub

    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowList(vriBo As Boolean)
        If vriBo Then
            DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psFillGrvList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            tbTrans.Style(HtmlTextWriterStyle.Visibility) = "visible"
        End If
    End Sub

    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
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
        psShowList(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblXlsProses.Text = ""
    End Sub

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoPenerimaanPembelian
            vnParam += "&vqTrNo=" & TxtTransNo.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)

            'Dim vnWinOpen As String
            'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
            'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
            'vnClientScript = Nothing
        End If
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        vnQuery = "Select PM.*,convert(varchar(11),PM.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "wh.WarehouseName,ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoRcvPOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA wh with(nolock) on wh.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoPenerimaanPembelian & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
            psButtonStatusDefault()
        Else
            TxtTransDate.Text = vnDtb.Rows(0).Item("vRcvPODate")
            TxtTransNo.Text = vnDtb.Rows(0).Item("RcvPONo")
            TxtTransRefNo.Text = vnDtb.Rows(0).Item("RcvPORefNo")

            TxtCompany.Text = vnDtb.Rows(0).Item("RcvPOCompanyCode")
            TxtWarehouse.Text = vnDtb.Rows(0).Item("WarehouseName")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            HdfRcvPORefOID.Value = vnDtb.Rows(0).Item("RcvPORefOID")

            HdfRcvType.Value = vnDtb.Rows(0).Item("RcvRefTypeOID")
            RdlRcvType.SelectedValue = HdfRcvType.Value

            If HdfRcvType.Value = enuRcvType.Pembelian Then
                HdfRcvPOType.Value = vnDtb.Rows(0).Item("RcvPORefTypeOID")
                RdlRcvPOType.SelectedValue = HdfRcvPOType.Value
                RdlRcvPOType.Visible = True
            Else
                RdlRcvPOType.Visible = False
            End If

            psButtonStatus()
        End If

        psFillGrvSumm(0, Val(TxtTransID.Text), vriSQLConn)
        PanData.Visible = False
        vnDtb.Dispose()
    End Sub

    Private Sub psButtonStatusDefault()
        BtnApprove.Enabled = False
        BtnList.Enabled = True
    End Sub
    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnApprove.Enabled = (HdfTransStatus.Value = enuTCRCPO.Receive_Done)
            BtnApprove.Visible = BtnApprove.Enabled
        End If
    End Sub
    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "Select" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)
            TxtTransID.Text = vnRow.Cells(ensColList.OID).Text

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

            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowList(False)
        End If
    End Sub
    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvSumm(0, TxtTransID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvSumm(vriEmpty As Byte, vriRcvPOHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String = ""

        If vriEmpty = 1 Then
            vnQuery = "Select 0 POHOID,''PO_NO,''BRG,''NAMA_BARANG,0 vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,0 vSumRcvPOScanQty Where 1=2"
        Else
            If HdfRcvType.Value = enuRcvType.Pembelian Then
                If HdfTransStatus.Value >= enuTCRCPO.Receive_Done Then
                    vnQuery = "Select rcm.POHOID,PO_NO,rcm.BRGCODE BRG,msb.BRGNAME NAMA_BARANG,rcm.SumPLQty vSumPLQty,rcm.SumPOQty vSumPOQty,0 vSumRetDRealQty,rcm.SumRcvPOScanQty vSumRcvPOScanQty,rcm.RcvPOQty_Total vRcvPOQty_Total,"
                    If HdfRcvPOType.Value = enuRcvPOType.Import Then
                        vnQuery += vbCrLf & "    (rcm.SumPLQty - rcm.SumRcvPOScanQty)vQtyVarian"
                    Else
                        vnQuery += vbCrLf & "    (rcm.SumPOQty - rcm.RcvPOQty_Total)vQtyVarian"
                    End If
                    vnQuery += vbCrLf & "From Sys_SsoRcvPOSummaryDone_TR rcm with(nolock)"
                    vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=rcm.RcvPOHOID"
                    vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=rcm.BRGCODE and msb.CompanyCode=rch.RcvPOCompanyCode"
                    vnQuery += vbCrLf & "     left outer join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=rcm.POHOID"
                    vnQuery += vbCrLf & "Where rch.OID=" & vriRcvPOHOID
                    vnQuery += vbCrLf & "order by case when isnull(rcm.SumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                Else
                    If HdfRcvPOType.Value = enuRcvPOType.Import Then
                        vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,0 vRcvPOQty_Total,(vSumRcvPOScanQty - vSumPLQty)vQtyVarian"
                        vnQuery += vbCrLf & "From fnTbl_SsoRcvPOImport_SummaryWithOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                    Else
                        vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,0 vSumPLQty,vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,vRcvPOQty_Total,(vRcvPOQty_Total - vSumPOQty)vQtyVarian"
                        vnQuery += vbCrLf & "From fnTbl_SsoRcvPOLocal_SummaryNonOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                    End If

                    vnQuery += vbCrLf & " Where (PO_NO like '%" & vnCriteria & "%' or BRG like '%" & vnCriteria & "%' or NAMA_BARANG like '%" & vnCriteria & "%')"
                    vnQuery += vbCrLf & "order by case when isnull(vSumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                End If

            ElseIf HdfRcvType.Value = enuRcvType.Retur Then
                vnQuery = "Select rcs.POHOID,rcs.PO_NO,rtd.BRG,rtd.NAMA_BARANG,0 vSumPLQty,0 vSumPOQty,rtd.vSumRetDRealQty,rcs.vSumRcvPOScanQty,0 vRcvPOQty_Total,0 vQtyVarian"
                vnQuery += vbCrLf & "From fnTbl_SsoRcvPOReturnScan(" & vriRcvPOHOID & "," & Session("UserOID") & ")rcs"
                vnQuery += vbCrLf & "     right outer join " & fbuGetDBDcm() & "fnTbl_DcmReturnDetailSumm() rtd on rtd.DcmRetHOID=rcs.RcvPORefOID and rtd.BRG=rcs.BRGCODE"
                vnQuery += vbCrLf & "  Where rtd.DcmRetHOID=" & HdfRcvPORefOID.Value & " and"
                vnQuery += vbCrLf & "        (rtd.BRG like '%" & vnCriteria & "%' or rtd.NAMA_BARANG like '%" & vnCriteria & "%')"
                vnQuery += vbCrLf & "order by case when isnull(rcs.vSumRcvPOScanQty,0)=0 then 5 else 1 end,rtd.BRG,rtd.NAMA_BARANG"

            ElseIf HdfRcvType.Value = enuRcvType.Lain_lain Then
                If HdfTransStatus.Value >= enuTCRCPO.Receive_Done Then
                    vnQuery = "Select rcm.POHOID,PO_NO,rcm.BRGCODE BRG,msb.BRGNAME NAMA_BARANG,rcm.SumPLQty vSumPLQty,rcm.SumPOQty vSumPOQty,0 vSumRetDRealQty,rcm.SumRcvPOScanQty vSumRcvPOScanQty,rcm.RcvPOQty_Total vRcvPOQty_Total,"

                    vnQuery += vbCrLf & "    (rcm.SumRcvPOScanQty - rcm.SumPLQty)vQtyVarian"

                    vnQuery += vbCrLf & "From Sys_SsoRcvPOSummaryDone_TR rcm with(nolock)"
                    vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=rcm.RcvPOHOID"
                    vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=rcm.BRGCODE and msb.CompanyCode=rch.RcvPOCompanyCode"
                    vnQuery += vbCrLf & "     left outer join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=rcm.POHOID"
                    vnQuery += vbCrLf & "Where rch.OID=" & vriRcvPOHOID
                    vnQuery += vbCrLf & "order by case when isnull(rcm.SumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                Else
                    vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,vSumRcvMscQty vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,0 vRcvPOQty_Total,(vSumRcvPOScanQty - vSumRcvMscQty)vQtyVarian"
                    vnQuery += vbCrLf & "From fnTbl_SsoRcvMsc_SummaryWithOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"

                    vnQuery += vbCrLf & " Where (PO_NO like '%" & vnCriteria & "%' or BRG like '%" & vnCriteria & "%' or NAMA_BARANG like '%" & vnCriteria & "%')"
                    vnQuery += vbCrLf & "order by case when isnull(vSumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                End If

            ElseIf HdfRcvType.Value = enuRcvType.Karantina Then
                If HdfTransStatus.Value >= enuTCRCPO.Receive_Done Then
                    vnQuery = "Select rcm.POHOID,PO_NO,rcm.BRGCODE BRG,msb.BRGNAME NAMA_BARANG,rcm.SumPLQty vSumPLQty,rcm.SumPOQty vSumPOQty,0 vSumRetDRealQty,rcm.SumRcvPOScanQty vSumRcvPOScanQty,rcm.RcvPOQty_Total vRcvPOQty_Total,"

                    vnQuery += vbCrLf & "    (rcm.SumRcvPOScanQty - rcm.SumPLQty)vQtyVarian"

                    vnQuery += vbCrLf & "From Sys_SsoRcvPOSummaryDone_TR rcm with(nolock)"
                    vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=rcm.RcvPOHOID"
                    vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=rcm.BRGCODE and msb.CompanyCode=rch.RcvPOCompanyCode"
                    vnQuery += vbCrLf & "     left outer join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=rcm.POHOID"
                    vnQuery += vbCrLf & "Where rch.OID=" & vriRcvPOHOID
                    vnQuery += vbCrLf & "order by case when isnull(rcm.SumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                Else
                    vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,vSumRcvKRQty vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,0 vRcvPOQty_Total,(vSumRcvPOScanQty - vSumRcvKRQty)vQtyVarian"
                    vnQuery += vbCrLf & "From fnTbl_SsoRcvKR_SummaryWithOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"

                    vnQuery += vbCrLf & " Where (PO_NO like '%" & vnCriteria & "%' or BRG like '%" & vnCriteria & "%' or NAMA_BARANG like '%" & vnCriteria & "%')"
                    vnQuery += vbCrLf & "order by case when isnull(vSumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
                End If
            End If
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
        If HdfRcvType.Value = enuRcvType.Pembelian Then
            If HdfRcvPOType.Value = enuRcvPOType.Import Then
                GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vQtyVarian_Import).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vQtyVarian_Import).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vQtyVarian_Local).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vQtyVarian_Local).ItemStyle.CssClass = "myDisplayNone"
            Else
                GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vQtyVarian_Import).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vQtyVarian_Import).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vQtyVarian_Local).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vQtyVarian_Local).ItemStyle.CssClass = ""
            End If
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).HeaderStyle.CssClass = ""
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).ItemStyle.CssClass = ""

            GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vRcvPOQty_Total).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vQtyVarian_Import).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vQtyVarian_Import).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vQtyVarian_Local).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vQtyVarian_Local).ItemStyle.CssClass = "myDisplayNone"
        End If
    End Sub
    Private Sub psFillGrvSumm_20230609_Orig(vriEmpty As Byte, vriRcvPOHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnCriteria As String = fbuFormatString(TxtFind.Text)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 POHOID,''PO_NO,''BRG,''NAMA_BARANG,0 vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,0 vSumRcvPOScanQty Where 1=2"
        Else
            If HdfRcvType.Value = enuRcvType.Pembelian Then
                If HdfRcvPOType.Value = enuRcvPOType.Import Then
                    vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty"
                    vnQuery += vbCrLf & "From fnTbl_SsoRcvPOImport_SummaryWithOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                Else
                    vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,0 vSumPLQty,vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty"
                    vnQuery += vbCrLf & "From fnTbl_SsoRcvPOLocal_SummaryNonOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                End If

                vnQuery += vbCrLf & " Where (PO_NO like '%" & vnCriteria & "%' or BRG like '%" & vnCriteria & "%' or NAMA_BARANG like '%" & vnCriteria & "%')"
                vnQuery += vbCrLf & "order by case when isnull(vSumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
            Else
                vnQuery = "Select rcs.POHOID,rcs.PO_NO,rtd.BRG,rtd.NAMA_BARANG,0 vSumPLQty,0 vSumPOQty,rtd.vSumRetDRealQty,rcs.vSumRcvPOScanQty"
                vnQuery += vbCrLf & "From fnTbl_SsoRcvPOReturnScan(" & vriRcvPOHOID & "," & Session("UserOID") & ")rcs"
                vnQuery += vbCrLf & "     right outer join " & fbuGetDBDcm() & "fnTbl_DcmReturnDetailSumm() rtd on rtd.DcmRetHOID=rcs.RcvPORefOID and rtd.BRG=rcs.BRGCODE"
                vnQuery += vbCrLf & "  Where rtd.DcmRetHOID=" & HdfRcvPORefOID.Value & " and"
                vnQuery += vbCrLf & "        (rtd.BRG like '%" & vnCriteria & "%' or rtd.NAMA_BARANG like '%" & vnCriteria & "%')"
                vnQuery += vbCrLf & "order by case when isnull(rcs.vSumRcvPOScanQty,0)=0 then 5 else 1 end,rtd.BRG,rtd.NAMA_BARANG"
            End If
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvSumm.DataSource = vnDtb
        GrvSumm.DataBind()
        If HdfRcvType.Value = enuRcvType.Pembelian Then
            If HdfRcvPOType.Value = enuRcvPOType.Import Then
                GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = "myDisplayNone"

                GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = ""
            Else
                GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = ""
                GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = ""

                GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = "myDisplayNone"
                GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = "myDisplayNone"
            End If
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvSumm.Columns(ensColSumm.PO_NO).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.PO_NO).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumPOQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumPOQty).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumPLQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvSumm.Columns(ensColSumm.vSumPLQty).ItemStyle.CssClass = "myDisplayNone"

            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).HeaderStyle.CssClass = ""
            GrvSumm.Columns(ensColSumm.vSumRetDRealQty).ItemStyle.CssClass = ""
        End If
    End Sub

    Protected Sub GrvSumm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSumm.SelectedIndexChanged

    End Sub

    Private Sub GrvSumm_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSumm.RowCommand
        If e.CommandName = "BRG" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvSumm.Rows(vnIdx)

            HdfPOHOID.Value = Val(vnGRow.Cells(ensColSumm.POHOID).Text)
            HdfStoKB.Value = DirectCast(vnGRow.Cells(ensColSumm.BRG).Controls(0), LinkButton).Text

            If HdfRcvPOType.Value = enuRcvPOType.Local Then
                LblDataTitle.Text = vnGRow.Cells(ensColSumm.PO_NO).Text & " - " & HdfStoKB.Value & " " & vnGRow.Cells(ensColSumm.NAMA_BARANG).Text
            Else
                LblDataTitle.Text = HdfStoKB.Value & " " & vnGRow.Cells(ensColSumm.NAMA_BARANG).Text
            End If

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvData(0, HdfPOHOID.Value, HdfStoKB.Value, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanData.Visible = True
        End If
    End Sub


    Private Sub psFillGrvData(vriEmpty As Byte, vriPOHOID As String, vriBrgCode As String, vriSQLConn As SqlClient.SqlConnection)
        If ChkSt_DelNo.Checked = False And ChkSt_DelYes.Checked = False Then
            ChkSt_DelNo.Checked = True
            ChkSt_DelYes.Checked = True
        End If

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "Select 0 OID,0 RcvPOScanQty,"
            vnQuery += vbCrLf & "       ''vRcvPOScanNote,"
            vnQuery += vbCrLf & "       ''vRcvPOScanUser,"
            vnQuery += vbCrLf & "	    ''vRcvPOScanTime,"
            vnQuery += vbCrLf & "	    ''vRcvPOScanDeleted,''vRcvPOScanDeletedUser,''RcvPOScanDeletedNote,"
            vnQuery += vbCrLf & "	    ''vRcvPOScanDeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem Where 1=2"
        Else
            vnQuery = "Select sc.OID,sc.RcvPOScanQty,"
            vnQuery += vbCrLf & "       sc.RcvPOScanNote vRcvPOScanNote,"
            vnQuery += vbCrLf & "       mu.UserName vRcvPOScanUser,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.RcvPOScanDatetime,106) + ' ' + convert(varchar(5),sc.RcvPOScanDatetime,108)vRcvPOScanTime,"
            vnQuery += vbCrLf & "	    case when abs(RcvPOScanDeleted)=1 then 'Y' else 'N' end vRcvPOScanDeleted,"
            vnQuery += vbCrLf & "       md.UserName vRcvPOScanDeletedUser,"
            vnQuery += vbCrLf & "	    RcvPOScanDeletedNote,"
            vnQuery += vbCrLf & "	    convert(varchar(11),sc.RcvPOScanDeletedDatetime,106) + ' ' + convert(varchar(5),sc.RcvPOScanDeletedDatetime,108)vRcvPOScanDeletedTime,"
            vnQuery += vbCrLf & "	    ''vDelItem"
            vnQuery += vbCrLf & "  From Sys_SsoRcvPOScan_TR sc with(nolock)"
            vnQuery += vbCrLf & "	    inner join Sys_SsoUser_MA mu with(nolock) on mu.OID=sc.RcvPOScanUserOID"
            vnQuery += vbCrLf & "	    left outer join Sys_SsoUser_MA md with(nolock) on md.OID=sc.RcvPOScanDeletedUserOID"
            vnQuery += vbCrLf & " Where sc.RcvPOHOID=" & TxtTransID.Text & " and sc.POHOID=" & vriPOHOID & " and sc.BrgCode='" & fbuFormatString(vriBrgCode) & "'"

            If Not (ChkSt_DelNo.Checked = True And ChkSt_DelYes.Checked = True) Then
                If ChkSt_DelNo.Checked = True Then
                    vnQuery += vbCrLf & "       and abs(RcvPOScanDeleted)=0"
                Else
                    vnQuery += vbCrLf & "       and abs(RcvPOScanDeleted)=1"
                End If
            End If
            vnQuery += vbCrLf & " Order by sc.OID"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        GrvData.DataSource = vnDtb
        GrvData.DataBind()

        If ChkSt_DelYes.Checked = True Then
            Dim vn As Integer
            For vn = 0 To GrvData.Rows.Count - 1
                If GrvData.Rows(vn).Cells(ensColData.vRcvPOScanDeleted).Text = "Y" Then
                    GrvData.Rows(vn).ForeColor = Drawing.Color.Red
                End If
            Next
        End If
    End Sub

    Protected Sub ChkSt_DelYes_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelYes.CheckedChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvData(0, HdfPOHOID.Value, HdfStoKB.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub ChkSt_DelNo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSt_DelNo.CheckedChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvData(0, HdfPOHOID.Value, HdfStoKB.Value, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub BtnApprove_Click(sender As Object, e As EventArgs) Handles BtnApprove.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Approve) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Approve Penerimaan No. " & TxtTransNo.Text & " ?<br/>WARNING : Approve Tidak Dapat Dibatalkan"
        HdfProcess.Value = "ApproveRcvPO"
        LblConfirmWarning.Text = ""

        psShowConfirm(True)
    End Sub
    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "ApproveRcvPO" Then
            psApproveRcvPO()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub psApproveRcvPO()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psApproveRcvPO", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnQuery As String
            Dim vnUserOID As String = Session("UserOID")
            Dim vnRcvPOHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            '1 Update Status Penerimaan PO
            vnQuery = "Update Sys_SsoRcvPOHeader_TR Set TransStatus=" & enuTCRCPO.Receive_Approved & ",ReceivedAppDatetime=Getdate(),ReceivedAppUserOID=" & vnUserOID & " Where OID=" & vnRcvPOHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1 Update Status Penerimaan PO")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            '2 Insert History Status Penerimaan
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2 Insert History Status Penerimaan...Start")
            pbuInsertStatusRcvPO(vnRcvPOHOID, enuTCRCPO.Receive_Approved, vnUserOID, vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("2 Insert History Status Penerimaan...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Approve Penerimaan Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            psDisplayData(vnSQLConn)

        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-99")
            vsTextStream.WriteLine("ERROR RAISED")
            vsTextStream.WriteLine(ex.Message)

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("--------------------------------- EOF ---------------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
        Finally
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub
End Class