Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Public Class WbfDmPerintahDOTDisplay
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfDmPerintahDOTDisplay"
    Const csTNoPrefix = "PKDO"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String
    Const csMaxByte = 1048576

    Enum ensColList
        OID = 0
    End Enum
    Enum ensColDetail
        OID = 0
        vAddItem = 1
        BRGCODE = 2
        BRGNAME = 3
        AvailableDOTQty = 4
        RequestQty = 5
        TxtRequestQty = 6
        PKDOTDQty = 7
        QtyOnPickList = 8
        vMessageItem = 9
        vDelItem = 10
    End Enum

    Enum ensColAttach
        OID = 0
        PKDOTImgNote = 1
        vUploadDatetime = 2
        vUploadDel = 3
    End Enum
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "DMgm"
        If Not IsPostBack Then
            psDefaultDisplay()
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPickList, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            End If

            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstWhs, False, vnSQLConn)
            pbuFillDstWarehouse(DstWhsDest, False, vnSQLConn)
            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstListWhs, True, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPKDOTDate.Text = ""
        TxtPKDOTScheduleDate.Text = ""
        TxtPKDOTNo.Text = ""
        TxtPKDOTNote.Text = ""
        TxtPKDOTCust.Text = ""
        TxtPKDOTShipToName.Text = ""
        TxtPKDOTShipToAddress.Text = ""

        ChkExpedition.Checked = False

        TxtPKDOTScheduleDate.Text = ""
        HdfTransStatus.Value = enuTCPerintahKirimDOT.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivAttach.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanAttach.Style(HtmlTextWriterStyle.Position) = "absolute"
        PanAttachImg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListItem.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListItem.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListCustomer.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListCustomer.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_InPicklist.Checked = False And ChkSt_Prepared.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_Prepared.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCPerintahKirimDOT.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCPerintahKirimDOT.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCPerintahKirimDOT.Prepared & ","
        End If
        If ChkSt_InPicklist.Checked = True Then
            vnCrStatus += enuTCPerintahKirimDOT.Dalam_Picklist & ","
        End If
        If ChkSt_PicklistDone.Checked = True Then
            vnCrStatus += enuTCPerintahKirimDOT.Picklist_Done & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,PM.PKDOTNo,convert(varchar(11),PM.PKDOTDate,106)vPKDOTDate,convert(varchar(11),PM.PKDOTScheduleDate,106)vPKDOTScheduleDate,"
        vnQuery += vbCrLf & "     PM.PKDOTCompanyCode,PM.CustCode+' '+mc.CUSTNAME vCustomer,PKDOTShipToName,"
        vnQuery += vbCrLf & "     case when abs(PM.IsExpedition)=0 then 'N' else 'Y' end vIsExpedition,"
        vnQuery += vbCrLf & "     case when abs(PM.IsPickListClosed)=1 then 'Y' else 'N' end vIsPickListClosed,"
        vnQuery += vbCrLf & "     PM.PKDOTShipToName,"
        vnQuery += vbCrLf & "     WM.WarehouseName,whd.WarehouseName vWarehouseName_Dest,"
        vnQuery += vbCrLf & "     PM.PKDOTNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared"

        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join (Select distinct CompanyCode,CUSTSUB,CUSTNAME From " & vnDBMaster & "Sys_MstCustomer_MA sa with(nolock))mc on mc.CUSTSUB=PM.CustCode and mc.CompanyCode=PM.PKDOTCompanyCode"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBMaster & "Sys_Warehouse_MA whd with(nolock) on whd.OID=PM.WarehouseOID_Dest"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     left outer join " & vnDBDcm & "Sys_DcmUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join " & vnDBDcm & "Sys_DcmUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.PKDOTCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If
        If Trim(TxtListNo.Text) <> "" Then
            vnQuery += vbCrLf & " and PM.PKDOTNo like '%" & Trim(TxtListNo.Text) & "%'"
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.PKDOTDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.PKDOTDate <= '" & TxtListEnd.Text & "'"
        End If
        If DstListWhs.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.WarehouseOID = " & DstListWhs.SelectedValue
        End If
        vnQuery += vbCrLf & "Order by PM.PKDOTNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()

        vnQuery = "Select dt.OID,'...'vAddItem,"
        vnQuery += vbCrLf & "       dt.BRGCODE,mb.BRGNAME,dt.AvailableDOTQty,dt.RequestQty,dt.PKDOTDQty,dt.QtyOnPickList,''vMessageItem,'Hapus Item'vDelItem"
        vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmPKDOTDetail_TR dt with(nolock)"
        vnQuery += vbCrLf & "            inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=dt.BRGCODE and mb.CompanyCode='" & DstCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "      Where dt.PKDOTHOID=" & vriHOID
        vnQuery += vbCrLf & "Order by dt.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        Dim vn As Integer
        GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
        GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

        GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
        GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

        GrvDetail.Columns(ensColDetail.RequestQty).HeaderStyle.CssClass = ""
        GrvDetail.Columns(ensColDetail.RequestQty).ItemStyle.CssClass = ""

        GrvDetail.Columns(ensColDetail.TxtRequestQty).HeaderStyle.CssClass = "myDisplayNone"
        GrvDetail.Columns(ensColDetail.TxtRequestQty).ItemStyle.CssClass = "myDisplayNone"

        If HdfTransStatus.Value <= enuTCPerintahKirimDOT.Baru Then
            GrvDetail.Columns(ensColDetail.QtyOnPickList).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.QtyOnPickList).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvDetail.Columns(ensColDetail.QtyOnPickList).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.QtyOnPickList).ItemStyle.CssClass = ""
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()
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
            psFillGrvList()
            psButtonShowList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psButtonShowList()
        BtnList.Enabled = True
    End Sub
    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psShowList(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgPKDOTDate.Text = ""
        LblMsgPKDOTScheduleDate.Text = ""
        LblMsgWhs.Text = ""
        LblMsgCust.Text = ""
        LblMsgWhsDest.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,abs(IsExpedition)vIsExpedition,"
        vnQuery += vbCrLf & "convert(varchar(11),PM.PKDOTDate,106)vPLDate,convert(varchar(11),PM.PKDOTScheduleDate,106)vPCLScheduleDate,"
        vnQuery += vbCrLf & "PM.CustCode+' '+mc.CUSTNAME vCustomer,ST.TransStatusDescr"
        vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join " & vnDBMaster & "Sys_MstCustomer_MA mc with(nolock) on mc.CUSTSUB=PM.CustCode and mc.CompanyCode=PM.PKDOTCompanyCode"
        vnQuery += vbCrLf & "     inner join " & vnDBDcm & "Sys_DcmTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.DcmPDOT & "'"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
            psFillGrvDetail(0, vriSQLConn)
        Else
            TxtPKDOTDate.Text = vnDtb.Rows(0).Item("vPLDate")
            TxtPKDOTNo.Text = vnDtb.Rows(0).Item("PKDOTNo")

            TxtPKDOTNote.Text = vnDtb.Rows(0).Item("PKDOTNote")

            TxtPKDOTScheduleDate.Text = vnDtb.Rows(0).Item("vPCLScheduleDate")
            TxtPKDOTCust.Text = vnDtb.Rows(0).Item("vCustomer")
            HdfCustCode.Value = vnDtb.Rows(0).Item("CustCode")

            TxtPKDOTShipToName.Text = vnDtb.Rows(0).Item("PKDOTShipToName")
            TxtPKDOTShipToAddress.Text = vnDtb.Rows(0).Item("PKDOTShipToAddress")

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            ChkExpedition.Checked = (vnDtb.Rows(0).Item("vIsExpedition") = "1")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("PKDOTCompanyCode"))
            HdfCompanyCode.Value = DstCompany.SelectedValue

            DstWhs.SelectedValue = Trim(vnDtb.Rows(0).Item("WarehouseOID"))
            HdfWhs.Value = DstWhs.SelectedValue

            DstWhsDest.SelectedValue = Trim(vnDtb.Rows(0).Item("WarehouseOID_Dest"))
            HdfWhsDest.Value = DstWhsDest.SelectedValue

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psFillGrvDetail(Val(TxtTransID.Text), vriSQLConn)
            psFillGrvAttach(Val(TxtTransID.Text), vriSQLConn)
            psFillGrvReserve(Val(TxtTransID.Text), "", "", vriSQLConn)
        End If
        vnDtb.Dispose()
    End Sub
    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
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

    Private Sub psShowListCustomer(vriBo As Boolean)
        If vriBo Then
            DivListCustomer.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListCustomer.Focus()
        Else
            DivListCustomer.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
        If e.CommandName = "BRGCODE" Then
            HdfDetailRowIdx.Value = vnIdx
            Dim vnBrgCode = DirectCast(vnGRow.Cells(ensColDetail.BRGCODE).Controls(0), LinkButton).Text
            LblMsgReserved.Text = "DETAIL QTY " & vnBrgCode & " " & vnGRow.Cells(ensColDetail.BRGNAME).Text
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvReserve(Val(TxtTransID.Text), vnBrgCode, vnGRow.Cells(ensColDetail.BRGNAME).Text, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnAttachClose_Click(sender As Object, e As EventArgs) Handles BtnAttachClose.Click
        psShowAttach(False)
    End Sub

    Private Sub GrvAttach_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvAttach.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvAttach.Rows(vnIdx)
        Dim vnAttachOID As String = vnGRow.Cells(ensColAttach.OID).Text
        If e.CommandName = "PKDOTImgNote" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            TxtAttachNote.Text = DirectCast(vnGRow.Cells(ensColAttach.PKDOTImgNote).Controls(0), LinkButton).Text
            psDisplayPhoto(vnAttachOID, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            psShowAttach(True)
        End If
    End Sub

    Private Sub psDisplayPhoto(vriOID As String, vriSQLConn As SqlConnection)
        Dim vnFileName As String
        Dim vnCmd = New SqlCommand("Select PKDOTImg From Sys_DcmPKDOTImg_TR Where OID=" & vriOID, vriSQLConn)
        Dim vnImageData As Byte() = DirectCast(vnCmd.ExecuteScalar(), Byte())
        If Not vnImageData Is Nothing Then
            Using ms As New MemoryStream(vnImageData, 0, vnImageData.Length)
                ms.Write(vnImageData, 0, vnImageData.Length)
                vnFileName = Format(Date.Now, "yyyyMMdd_HHmmss") & "_" & vriOID & ".jpg"
                Image.FromStream(ms, True).Save(Server.MapPath("~") & "\FileTemp\" & vnFileName)
                ImgAttachImg.ImageUrl = "~/FileTemp/" & vnFileName
                ImgAttachImg.ResolveUrl("~/FileTemp/" & vnFileName)
            End Using
        End If
        PanAttachImg.Visible = True
    End Sub

    Private Sub psShowAttach(vriBo As Boolean)
        If vriBo Then
            DivAttach.Style(HtmlTextWriterStyle.Visibility) = "visible"

            TxtAttachNote.Focus()
        Else
            DivAttach.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psFillGrvAttach(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriHOID = 0 Then
            vnQuery = "Select 0 OID,''PKDOTImgNote,''vUploadDatetime,''vUploadDel Where 1=2"
        Else
            Dim vnDBDcm As String = fbuGetDBDcm()
            If HdfTransStatus.Value = enuTCPerintahKirimDOT.Baru Then
                vnQuery = "Select * From("
                vnQuery += vbCrLf & "Select OID,PKDOTImgNote,convert(varchar(11),UploadDatetime,106)+' '+convert(varchar(11),UploadDatetime,108)vUploadDatetime,'Hapus'vUploadDel"
                vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTImg_TR"
                vnQuery += vbCrLf & "Where PKDOTHOID=" & vriHOID
                vnQuery += vbCrLf & "UNION"
                vnQuery += vbCrLf & "Select 0 OID,''PKDOTImgNote,''vUploadDatetime,'Attach'vUploadDel"
                vnQuery += vbCrLf & ")tb Order by case when OID=0 then 5 else 4 end,OID"
            Else
                vnQuery = "Select OID,PKDOTImgNote,convert(varchar(11),UploadDatetime,106)+' '+convert(varchar(11),UploadDatetime,108)vUploadDatetime,''vUploadDel"
                vnQuery += vbCrLf & "From " & vnDBDcm & "Sys_DcmPKDOTImg_TR"
                vnQuery += vbCrLf & "Where PKDOTHOID=" & vriHOID
                vnQuery += vbCrLf & "Order by case when OID=0 then 5 else 4 end,OID"
            End If
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvAttach.DataSource = vnDtb
        GrvAttach.DataBind()
    End Sub
    Private Sub psFillGrvReserve(vriHOID As String, vriBrgCode As String, vriBrgName As String, vriSQLConn As SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If vriHOID = 0 Then
            vnQuery = "Select ''BRGCODE,''BRGNAME,0 NotaHOID,''NotaNo,''vNotaDate,0 PKDOTNQty where 1=2"
        Else
            vnQuery = "Select pn.BRGCODE,mb.BRGNAME,pn.NotaHOID,nh.NotaNo,convert(varchar(11),nh.NotaDate,106)vNotaDate,pn.PKDOTNQty"
            vnQuery += vbCrLf & "       From " & vnDBDcm & "Sys_DcmPKDOTNota_TR pn with(nolock)"
            vnQuery += vbCrLf & "	         inner join " & vnDBDcm & "Sys_DcmNotaHeader_TR nh with(nolock) on nh.OID=pn.NotaHOID"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.BRGCODE=pn.BRGCODE and mb.CompanyCode=nh.CompanyCode"
            vnQuery += vbCrLf & "Where pn.PKDOTHOID=" & vriHOID
            If vriBrgCode <> "" Then
                vnQuery += vbCrLf & "      and pn.BRGCODE='" & vriBrgCode & "'"
            End If
            vnQuery += vbCrLf & "Order by pn.BRGCODE,nh.NotaDate,nh.NotaNo"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvReserved.DataSource = vnDtb
        GrvReserved.DataBind()
        If vriBrgCode = "" Then
            LblMsgReserved.Text = "DETAIL QTY"
        Else
            LblMsgReserved.Text = "DETAIL QTY " & vriBrgCode & " " & vriBrgName
        End If
    End Sub

    Protected Sub BtnSummary_Click(sender As Object, e As EventArgs) Handles BtnSummary.Click
        Response.Redirect("~/Reporting/WbfDmDOTSummary.aspx")
    End Sub

End Class