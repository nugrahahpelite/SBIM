Imports System.Data.SqlClient
Public Class WbfSsoCSKU
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoCSKU"
    Const csTNoPrefix = "CSKU"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        OID = 0
        vChangeStartDate = 1
        CompanyCode = 2
        BRGCODE_LAMA = 3
        BRGCODE_BARU = 4
        TransStatusDescr = 5
        vCreation = 6
        vPrepared = 7
        vApproved = 8
    End Enum

    Enum ensColListBrg
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
    End Enum
    Enum ensColDetail
        OID = 0
        vStorageOID = 1
        vStorageInfo_Complete = 2
        RcvPOHOID = 3
        RCVPONo = 4
        vRcvPODate = 5
        QtyOnHand = 6
        vPrintItem = 7
    End Enum
    Public Enum enuTCCSKU 'SsoChangeSKU
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Approved = 4
    End Enum
    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtPergantianDate.Text = ""
        LblDetail.Text = ""
        LblOSBrg.Text = ""
        HdfTransStatus.Value = enuTCCSKU.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivBrg2.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanBrg2.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivOutstanding.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        DivOutstanding.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrint.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrint.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoChangeSKU, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If




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
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        If ChkSt_Baru.Checked = False And ChkSt_Cancelled.Checked = False And ChkSt_Prepared.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_Prepared.Checked = True
            ChkSt_Approved.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCCSKU.Baru & ","
        End If
        If ChkSt_Cancelled.Checked = True Then
            vnCrStatus += enuTCCSKU.Cancelled & ","
        End If
        If ChkSt_Prepared.Checked = True Then
            vnCrStatus += enuTCCSKU.Prepared & ","
        End If

        If ChkSt_Approved.Checked = True Then
            vnCrStatus += enuTCCSKU.Approved & ","
        End If

        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and PM.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If
        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "	SELECT PM.OID,convert(varchar(11),PM.ChangeStartDate,106)vChangeStartDate,	"
        vnQuery += vbCrLf & "	  PM.CompanyCode, PM.BRGCODE_LAMA, PM.BRGCODE_BARU,PM.TransStatus, PM.TransCode,	"
        vnQuery += vbCrLf & "	  ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CR.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.ApprovedDatetime,106)+' '+convert(varchar(5),PM.ApprovedDatetime,108)+' '+ PR.UserName vApproved"
        vnQuery += vbCrLf & " FROM Sys_SsoCSKUHeader_TR PM with(nolock)	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode	"
        vnQuery += vbCrLf & "	  inner join Sys_SsoUser_MA CR with(nolock) on CR.OID=PM.CreationUserOID	"
        vnQuery += vbCrLf & "	  left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID	"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus

        If Val(TxtListTransID.Text) > 0 Then
            vnQuery += vbCrLf & " and PM.OID=" & Val(TxtListTransID.Text)
        End If


        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.ChangeStartDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.ChangeStartDate <= '" & TxtListEnd.Text & "'"
        End If

        If DstListCompany.SelectedIndex > 0 Then
            vnQuery += vbCrLf & "            and PM.CompanyCode = '" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & "Order by PM.CreationDatetime Desc"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvDetail(vriHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If (vriHOID = "0") Then
            vnQuery = "	Select 0 OID,0 vStorageOID,'' vStorageInfo_Complete,0 RcvPOHOID,'' RcvPONo,''vRcvPODate,0 QtyOnHand,''vPrintItem	"
            vnQuery += vbCrLf & "	 Where 1=2"
        Else
            Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

            vnQuery = "	Select pm.OID , pm.StorageOID vStorageOID, sm.vStorageInfo_Complete,pm.RcvPOHOID,rc.RCVPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,pm.QtyOnHand, 'Print'vPrintItem	"
            vnQuery += vbCrLf & "	 From Sys_SsoCSKUStorageStock_TR pm	"
            vnQuery += vbCrLf & "	      inner join " & vnDBMaster & "fnTbl_SsoStorageInfo('') sm on sm.vStorageOID =pm.StorageOID"
            vnQuery += vbCrLf & "	      inner join Sys_SsoRcvPOHeader_TR rc on rc.OID=pm.RcvPOHOID"

            If vnUserWarehouseCode <> "" Then
                vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=sm.WarehouseOID and uw.UserOID=" & Session("UserOID")
            End If

            vnQuery += vbCrLf & "	Where pm.CSKUHOID = " & TxtTransID.Text & "	"
            vnQuery += vbCrLf & "	Order by sm.vStorageOID"

        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If HdfTransStatus.Value = enuTCCSKU.Baru Then
            GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = "myDisplayNone"
        ElseIf (HdfTransStatus.Value = enuTCCSKU.Prepared) Then
            GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = "myDisplayNone"
        ElseIf (HdfTransStatus.Value = enuTCCSKU.Approved) Then
            GrvDetail.Columns(ensColDetail.vPrintItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vPrintItem).ItemStyle.CssClass = ""
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()
    End Sub

    Private Sub psFillGrvOSBrg(vriEmpty As Byte, vriCompanyCode As String, vriBrgCode_Lama As String, vriSQLConn As SqlConnection)
        psClearMessage()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriEmpty = 1 Then
            vnQuery = "	Select 0 vTransOID,''TransCode,'' vTransType,''vTransNo where 1=2"
        Else
            vnQuery = "	Select pm.vTransOID,pm.TransCode,pm.vTransType,pm.vTransNo"
            vnQuery += vbCrLf & "	 From fnTbl_SsoValApproveCSKU('" & vriCompanyCode & "','" & vriBrgCode_Lama & "') pm"
            vnQuery += vbCrLf & "	Order by pm.vTransNo"
        End If

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvOSBrg.DataSource = vnDtb
        GrvOSBrg.DataBind()
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

    Private Sub psShowPrint(vriBo As Boolean)
        If vriBo Then
            DivPrint.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPrint.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowList(vriBo As Boolean)
        If vriBo Then
            DivList.Style(HtmlTextWriterStyle.Visibility) = "visible"
            psFillGrvList()
            psButtonShowList()
        Else
            DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
            psButtonStatus()
        End If
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnCancelPCL.Enabled = False
        BtnPrepare.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True
    End Sub
    Protected Sub BtnList_Click(sender As Object, e As EventArgs) Handles BtnList.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psShowList(True)
        psButtonVisible()
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        psClearText()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Session(csModuleName & stuSession.Simpan) = ""

        psClearData()

        If DstCompany.Items.Count > 0 Then
            DstCompany.SelectedIndex = 0
        End If

        TxtPergantianDate.Text = fbuGetDateTodaySQL(vnSQLConn)

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail("0", vnSQLConn)
        psFillGrvOSBrg(1, "", "", vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing


        psEnableSave(True)


    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgPergantianDate.Text = ""
        LblMsgError.Text = ""
        LblMsgListBrg.Text = ""
        LblMsgListBrg2.Text = ""
        LblMsgXlsProsesError.Text = ""
        LblConfirmMessage.Text = ""
        LblConfirmNote.Text = ""
        LblConfirmProgress.Text = ""
        LblConfirmWarning.Text = ""
        LblFindProgress.Text = ""
        LblProgressBrg.Text = ""
        LblProgressBrg2.Text = ""
        LblProgressFind.Text = ""
        LblProgressSave.Text = ""

    End Sub

    Private Sub psClearText()
        TxtListBrg.Text = ""
        TxtListBrgCode.Text = ""
        TxtListBrgCode2.Text = ""
        TxtListBrgName.Text = ""
        TxtListBrgName2.Text = ""
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtListTransID.Text = ""

        TxtListEnd.Text = ""
        TxtListStart.Text = ""
        TxtPergantianDate.Text = ""
        TxtPergantianNo.Text = ""
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelPCL.Visible = Not vriBo
        BtnPrepare.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo

        BtnList.Visible = Not vriBo
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
    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoChangeSKU


            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)

            'Dim vnWinOpen As String
            'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
            'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
            'vnClientScript = Nothing
        End If
    End Sub

    Private Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearText()
        psClearMessage()

        psEnableSave(False)
        psButtonVisible()

        HdfActionStatus.Value = cbuActionNorm
        If TxtTransID.Text = "" Then
            psClearData()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvDetail("0", vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        Else
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            psFillGrvDetail("2", vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psDisplayData(vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtTransID.Text = "" Then
            psClearData()
            Exit Sub
        End If

        HdfActionStatus.Value = cbuActionNorm

        vnQuery = "Select PM.*,convert(varchar(11),PM.ChangeStartDate,106)vChangeStartDate,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoCSKUHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode='" & stuTransCode.SsoChangeSKU & "'"

        vnQuery += vbCrLf & "     Where PM.OID =" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
            psFillGrvOSBrg(1, "", "", vriSQLConn)
            psFillGrvDetail(0, vriSQLConn)
        Else

            TxtPergantianNo.Text = vnDtb.Rows(0).Item("OID")
            TxtPergantianDate.Text = vnDtb.Rows(0).Item("vChangeStartDate")

            TxtListBrgCode.Text = vnDtb.Rows(0).Item("BRGCODE_LAMA")
            TxtListBrgCode2.Text = vnDtb.Rows(0).Item("BRGCODE_BARU")

            LblDetail.Text = "DATA STOCK " & TxtListBrgCode.Text
            LblOSBrg.Text = "OUTSTANDING TRANSAKSI " & TxtListBrgCode.Text
            psFillGrvDetail("2", vriSQLConn)

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("CompanyCode"))

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")
            TxtListBrgName.Text = fbuGetNamaBarang(Trim(vnDtb.Rows(0).Item("CompanyCode")), vnDtb.Rows(0).Item("BRGCODE_LAMA"), vriSQLConn)
            TxtListBrgName2.Text = fbuGetNamaBarang(Trim(vnDtb.Rows(0).Item("CompanyCode")), vnDtb.Rows(0).Item("BRGCODE_BARU"), vriSQLConn)

            If HdfTransStatus.Value = enuTCCSKU.Cancelled Or HdfTransStatus.Value = enuTCCSKU.Approved Then
                psFillGrvOSBrg(1, "", "", vriSQLConn)
            Else
                psFillGrvOSBrg(0, DstCompany.SelectedValue, TxtListBrgCode.Text, vriSQLConn)
            End If

            psFillGrvDetail("2", vriSQLConn)
            psButtonStatus()
        End If
        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelPCL.Visible = BtnCancelPCL.Enabled
        BtnPrepare.Visible = BtnPrepare.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelPCL.Enabled = False
        BtnPrepare.Enabled = False

        BtnPreview.Enabled = False
        BtnList.Enabled = True

        psButtonVisible()
    End Sub

    Private Sub psButtonStatus()
        If TxtTransID.Text = "" Then
            psButtonStatusDefault()
        Else
            BtnBaru.Enabled = True
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCCSKU.Baru)

            BtnCancelPCL.Enabled = (HdfTransStatus.Value = enuTCCSKU.Baru Or HdfTransStatus.Value = enuTCCSKU.Prepared)
            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCCSKU.Baru)
            BtnApprove.Enabled = (HdfTransStatus.Value = enuTCCSKU.Prepared)
            BtnPreview.Enabled = (HdfTransStatus.Value = enuTCCSKU.Approved)

            psButtonVisible()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Trim(TxtTransID.Text) = "" Then Exit Sub
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Session(csModuleName & stuSession.Simpan) = ""

        HdfActionStatus.Value = cbuActionEdit
        psEnableSave(True)
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        psSaveBaru()
    End Sub
    Private Sub psSaveBaru()
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Session(csModuleName & stuSession.Simpan) <> "" Then
            Exit Sub
        End If

        Dim vnSave As Boolean = True
        psClearMessage()


        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            vnSave = False
        End If

        If Not IsDate(Trim(TxtPergantianDate.Text)) Then
            LblMsgPergantianDate.Text = "Isi Tanggal"
            vnSave = False
        End If

        If Not vnSave Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        Dim vnSQLTrans As SqlTransaction
        vnSQLTrans = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnQuery As String
            Dim vnUserNIP As String = Session("UserName")

            Dim vnBrgCode_Lama As String = Trim(TxtListBrgCode.Text)
            Dim vnBrgCode_Baru As String = Trim(TxtListBrgCode2.Text)

            If HdfActionStatus.Value = cbuActionNew Then
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)
                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoCSKUHeader_TR with(nolock)"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoCSKUHeader_TR(OID,ChangeStartDate,"
                vnQuery += vbCrLf & "CompanyCode,BRGCODE_LAMA,BRGCODE_BARU,"

                vnQuery += vbCrLf & "TransCode,TransStatus,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & ",'" & TxtPergantianDate.Text & "',"
                vnQuery += vbCrLf & "'" & Trim(vnCompanyCode) & "','" & vnBrgCode_Lama & "','" & vnBrgCode_Baru & "',"

                vnQuery += vbCrLf & "'" & stuTransCode.SsoChangeSKU & "',0," & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                psSaveDetail(vnOID, vnBrgCode_Lama, vnSQLConn, vnSQLTrans)

                pbuInsertStatusCSKU(vnOID, enuTCCSKU.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCCSKU.Baru

                Session(csModuleName & stuSession.Simpan) = "Done"

                psButtonStatus()
                BtnPrepare.Enabled = True
                BtnPrepare.Visible = True

            Else
                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoCSKUHeader_TR set"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                psSaveDetail(TxtTransID.Text, vnBrgCode_Lama, vnSQLConn, vnSQLTrans)

                pbuInsertStatusCSKU(TxtTransID.Text, enuTCCSKU.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"
                psButtonStatus()
                BtnPrepare.Enabled = True
                BtnPrepare.Visible = True

            End If

            psEnableSave(False)

            HdfActionStatus.Value = cbuActionNorm
            psDisplayData(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        Catch ex As Exception
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub
    Private Sub psSaveDetail(vriCSKUHOID As String, vriBrgCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Delete Sys_SsoCSKUStorageStock_TR Where CSKUHOID=" & vriCSKUHOID
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoCSKUStorageStock_TR(CSKUHOID,StorageOID,RcvPOHOID,BRGCODE,QtyOnHand)"
        vnQuery += vbCrLf & "Select " & vriCSKUHOID & ",StorageOID,RcvPOHOID,BRGCODE,QtyOnHand"
        vnQuery += vbCrLf & "  From Sys_SsoStorageStock_MA with(nolock) Where BRGCODE='" & vriBrgCode & "' and QtyOnHand<>0"
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
    End Sub

    Private Sub psSaveDetail_20231007_nugraha(vriCSKUHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        'Dim vnQuery As String
        'Dim vn As Integer
        'Dim vnGRow As GridViewRow

        'Dim vnTxtQtyOnHand As TextBox
        'Dim vnBrgCode As String
        'Dim vnQtyOnHand As Integer
        'For vn = 0 To GrvDetail.Rows.Count - 1
        '    vnGRow = GrvDetail.Rows(vn)
        '    vnTxtQtyOnHand = vnGRow.FindControl("QtyOnHand")

        '    If fbuValStrHtml(vnGRow.Cells(ensColDetail.OID).Text) = "0" Then
        '        If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
        '            vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

        '            vnQuery = "Insert into Sys_SsoCSKUStorageStock_TR"
        '            vnQuery += vbCrLf & "(CSKUHOID,"
        '            vnQuery += vbCrLf & "StorageOID,"
        '            vnQuery += vbCrLf & "RcvPOHOID,"
        '            vnQuery += vbCrLf & "BRGCODE,"
        '            vnQuery += vbCrLf & "QtyOnHand"
        '            vnQuery += vbCrLf & ")"
        '            vnQuery += vbCrLf & "values(" & vriCSKUHOID & ","
        '            vnQuery += vbCrLf & "" & vnGRow.Cells(1).Text & ","
        '            vnQuery += vbCrLf & "" & vnGRow.Cells(3).Text & ","
        '            vnQuery += vbCrLf & "'" & TxtListBrgCode.Text & "',"
        '            vnQuery += vbCrLf & "" & vnGRow.Cells(6).Text & ")"

        '            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        '        End If
        '    Else
        '        vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

        '        vnQuery = "Update Sys_SsoCSKUStorageStock_TR SET"
        '        vnQuery += vbCrLf & "QtyOnHand=" & Val(Replace(vnQtyOnHand, "", "")) & ""

        '        vnQuery += vbCrLf & "Where OID=" & vnGRow.Cells(ensColDetail.OID).Text
        '        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
        '    End If
        'Next
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "Select" Then
            Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvList.Rows(vnRowIdx)
            TxtTransID.Text = vnGRow.Cells(0).Text

            HdfCompanyCode.Value = vnGRow.Cells(ensColList.CompanyCode).Text
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

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelCSKU" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelCSKU()
        ElseIf HdfProcess.Value = "PrepareCSKU" Then
            psPrepareCSKU()
        ElseIf HdfProcess.Value = "ApproveCSKU" Then
            psApproveCSKU()
        End If
        psButtonStatus()
        psShowConfirm(False)
    End Sub

    Private Sub BtnPrepare_Click(sender As Object, e As EventArgs) Handles BtnPrepare.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Prepare) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Prepare Pergantian No. " & TxtTransID.Text & " ?<br />WARNING : Prepare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "PrepareCSKU"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub psCancelCSKU()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Try
            Dim vnQuery As String

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoCSKUHeader_TR set TransStatus=" & enuTCCSKU.Cancelled & ",CSKUCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            pbuInsertStatusCSKU(TxtTransID.Text, enuTCCSKU.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psDisplayData(vnSQLConn)

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

    Private Sub psPrepareCSKU()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psPrepareCSKU", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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

        psFillGrvOSBrg(0, DstCompany.SelectedValue, TxtListBrgCode.Text, vnSQLConn)
        If GrvOSBrg.Rows.Count > 0 Then
            LblMsgError.Text = "Masih Ada Outstanding Transaksi...Prepare Gagal."
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(LblMsgError.Text)
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
            Dim vnCSKUHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoCSKUHeader_TR set TransStatus=" & enuTCCSKU.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnCSKUHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusCSKU...Start")
            pbuInsertStatusCSKU(vnCSKUHOID, enuTCCSKU.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusCSKU...End")
            HdfTransStatus.Value = enuTCCSKU.Prepared

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Prepare Sukses")
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

    Protected Sub BtnCancelPCL_Click(sender As Object, e As EventArgs) Handles BtnCancelPCL.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Membatalkan Change SKU No. " & TxtTransID.Text & " ?<br />WARNING : Batal Change SKU Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelCSKU"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True

        psShowConfirm(True)
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        'psClearMessage()
        'If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)

        'Dim vnCrpFileName As String = ""
        'psGenerateCrp(vnCrpFileName)

        'Dim vnRootURL As String = ConfigurationManager.AppSettings("WebRootFolder")
        'Dim vnParam As String
        'vnParam = "vqCrpPreviewType=" & stuCrpPreviewType.ByQueryPopwin
        'vnParam += "&vqCrpFileName=" & vnCrpFileName
        'vnParam += "&vqCrpSubReport1="
        'vnParam += "&vqCrpSubReport2="
        'vnParam += "&vqCrpSubReport3="
        'vnParam += "&vqCrpSubReport4="
        'vnParam += "&vqCrpQuery=" & vbuCrpQuery
        'vnParam += "&vqCrpQuery1="
        'vnParam += "&vqCrpQuery2="
        'vnParam += "&vqCrpQuery3="
        'vnParam += "&vqCrpQuery4="
        'vnParam += "&vqCrpPreview=Pdf"

        'vbuPreviewOnClose = "0"

        'ifrPreview.Src = vnRootURL & "Preview/WbfCrpViewer.aspx?" & vnParam
        'psShowPreview(True)
    End Sub

    Private Sub psShowListBrg(vriBo As Boolean)
        If vriBo Then
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListBrg.Focus()
        Else
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub psShowListBrg2(vriBo As Boolean)
        If vriBo Then
            DivBrg2.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtBrg2.Focus()
        Else
            DivBrg2.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub

    Private Sub psShowBrg2(vriBo As Boolean)
        If vriBo Then
            DivBrg2.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListBrg.Focus()
        Else
            DivBrg2.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If

    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnListBrgFind_Click(sender As Object, e As EventArgs) Handles BtnListBrgFind.Click
        LblMsgListBrg.Text = ""
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If
        If Trim(TxtListBrg.Text) = "" Then
            LblMsgListBrg.Text = "Pilih Barang"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListBrg(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvListBrg(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        LblMsgListBrg.Text = ""

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String

        Dim vnBrg As String = fbuFormatString(Trim(TxtListBrg.Text))

        vnCriteria = "      Where CompanyCode='" & DstCompany.SelectedValue & "'"
        vnCriteria += vbCrLf & "            and (BRGCODE like '%" & vnBrg & "%' or BRGNAME like '%" & vnBrg & "%')"

        vnQuery = "SELECT BRGCODE,BRGNAME, BRGUNIT FROM " & vnDBMaster & "Sys_MstBarang_MA"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvListBrg.DataSource = vnDtb
        GrvListBrg.DataBind()

        TxtListBrg.Focus()
    End Sub

    Protected Sub BtnListBrgClose_Click(sender As Object, e As EventArgs) Handles BtnListBrgClose.Click
        psShowListBrg(False)
    End Sub

    Private Sub GrvListBrg_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvListBrg.PageIndexChanging
        GrvListBrg.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvListBrg(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvListBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListBrg.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRowList As GridViewRow = GrvListBrg.Rows(vnIdx)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        If (Val(TxtTransID.Text) > 0) Then
            Dim vnCSKUHOID As Integer = Convert.ToInt32(e.CommandArgument)

        Else
            If e.CommandName = "BRGCODE" Then
                Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListBrg.BRGCODE).Controls(0), LinkButton).Text

                Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

                vnGRowDetail.Cells(ensColListBrg.BRGCODE).Text = vnKodeBarang
                vnGRowDetail.Cells(ensColListBrg.BRGNAME).Text = vnGRowList.Cells(ensColListBrg.BRGNAME).Text

                psShowListBrg(False)
            End If
        End If

    End Sub



    Protected Sub BtnApprove_Click(sender As Object, e As EventArgs) Handles BtnApprove.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Prepare) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        LblConfirmMessage.Text = "Anda Akan Approve Pergantian SKU No. " & TxtTransID.Text & " ?<br />WARNING : Approval Tidak Dapat Dibatalkan"
        HdfProcess.Value = "ApproveCSKU"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False

        psShowConfirm(True)
    End Sub

    Private Sub psApproveCSKU()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psApproveCSKU", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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

        psFillGrvOSBrg(0, DstCompany.SelectedValue, TxtListBrgCode.Text, vnSQLConn)
        If GrvOSBrg.Rows.Count > 0 Then
            LblMsgError.Text = "Masih Ada Outstanding Transaksi...Approve Gagal."
            LblMsgError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(LblMsgError.Text)
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
            Dim vnCSKUHOID As String = TxtTransID.Text

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Execute spSsoApprove_CSKU " & vnCSKUHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoCSKUHeader_TR set TransStatus=" & enuTCCSKU.Approved & ",ApprovedUserOID=" & Session("UserOID") & ",ApprovedDatetime=getdate() Where OID=" & vnCSKUHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusCSKU...Start")
            pbuInsertStatusCSKU(vnCSKUHOID, enuTCCSKU.Approved, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusCSKU...End")
            HdfTransStatus.Value = enuTCCSKU.Approved

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Approve Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing
            HdfTransStatus.Value = enuTCCSKU.Approved

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

    Private Sub psFillGrvBrg2(vriSQLConn As SqlConnection)
        Dim vnDBMaster As String = fbuGetDBMaster()
        LblMsgListBrg.Text = ""

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String

        Dim vnBrg As String = fbuFormatString(Trim(TxtBrg2.Text))

        vnCriteria = "      Where CompanyCode='" & DstCompany.SelectedValue & "'"
        vnCriteria += vbCrLf & "            and (BRGCODE like '%" & vnBrg & "%' or BRGNAME like '%" & vnBrg & "%')"

        vnQuery = "SELECT BRGCODE,BRGNAME, BRGUNIT FROM " & vnDBMaster & "Sys_MstBarang_MA"
        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvBrg2.DataSource = vnDtb
        GrvBrg2.DataBind()

        TxtBrg2.Focus()
    End Sub



    Protected Sub BtnListBrgFind2_Click(sender As Object, e As EventArgs) Handles BtnListBrgFind2.Click
        LblMsgListBrg.Text = ""
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If
        If Trim(TxtListBrg.Text) = "" Then
            LblMsgListBrg.Text = "Pilih Barang"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvBrg2(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListBrgClose2_Click(sender As Object, e As EventArgs) Handles BtnListBrgClose2.Click
        psShowListBrg2(False)
    End Sub

    Protected Sub GrvListBrg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvListBrg.SelectedIndexChanged

    End Sub

    Protected Sub GrvBrg2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvBrg2.SelectedIndexChanged

    End Sub

    Private Sub psFillGrvLsBrg2()
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstBarang_MA PM"
        vnQuery += vbCrLf & "Where CompanyCode='" & DstListCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtBrg2.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtBrg2.Text) & "%')"
        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvBrg2.DataSource = vnDtb
        GrvBrg2.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvLsBrg()
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstBarang_MA PM"
        vnQuery += vbCrLf & "Where CompanyCode='" & DstListCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtBrg2.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtBrg2.Text) & "%')"
        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvListBrg.DataSource = vnDtb
        GrvListBrg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnListBrgCode_Click(sender As Object, e As EventArgs) Handles BtnListBrgCode.Click
        If Val(DstCompany.SelectedValue) > 0 Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psShowLsBrg(True)
    End Sub

    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub GrvLsBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvListBrg.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvListBrg.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(0).Controls(0), LinkButton).Text
            TxtListBrgCode.Text = vnValue
            TxtListBrgName.Text = vnRow.Cells(1).Text
            psShowLsBrg(False)
        End If
    End Sub
    Private Sub GrvBrg2_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvBrg2.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvBrg2.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(0).Controls(0), LinkButton).Text
            If (vnValue = TxtListBrgCode.Text) Then
                LblProgressBrg2.Text = "Tidak dapat memilih barang pengganti sama dengan barang sebelumnya"
            Else

                TxtListBrgCode2.Text = vnValue
                TxtListBrgName2.Text = vnRow.Cells(1).Text
                psShowBrg2(False)
            End If
        End If

    End Sub

    Protected Sub BtnListBrgCode2_Click(sender As Object, e As EventArgs) Handles BtnListBrgCode2.Click
        If Val(DstCompany.SelectedValue) > 0 Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psShowBrg2(True)
    End Sub

    Protected Sub GrvDetail_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvDetail.SelectedIndexChanged

    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
        HdfRcvPONo.Value = vnGRow.Cells(4).Text
        HdfBrgCode.Value = TxtListBrgCode.Text
        If e.CommandName = "vPrintItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If

            HdfPrintRcvPOHOID.Value = vnGRow.Cells(ensColDetail.RcvPOHOID).Text
            HdfPrintRcvPONo.Value = vnGRow.Cells(ensColDetail.RCVPONo).Text
            HdfPrintRcvPODate.Value = vnGRow.Cells(ensColDetail.vRcvPODate).Text

            LblPrintBrgCode.Text = TxtListBrgCode2.Text & " " & TxtListBrgName2.Text
            LblPrintRcvPONo.Text = HdfPrintRcvPONo.Value
            LblPrintRcvPODate.Text = HdfPrintRcvPODate.Value

            TxtPrintCount.Text = 1

            psShowPrint(True)
        End If
    End Sub

    Private Sub psPrintBrg()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnQuery As String
        Dim vnRcvPOHOID As String = HdfPrintRcvPOHOID.Value
        Dim vnRcvPONo As String = HdfPrintRcvPONo.Value
        Dim vnRcvPODate As String = HdfPrintRcvPODate.Value

        Dim vnBrgCode As String = TxtListBrgCode2.Text
        Dim vnBrgName As String = TxtListBrgName2.Text

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean
        Try
            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Select count(1) From Sys_SsoRcvPOBarangQRCode_TR Where RcvPOHOID=" & vnRcvPOHOID & " and BRGCODE='" & vnBrgCode & "'"
            If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                If fsGenBrgQRCode(vnRcvPOHOID, vnBrgCode, vnBrgCode & Space(5) & Chr(10) & vnBrgName & Chr(10) & "No.Terima:" & vnRcvPONo & Chr(10) & "Tgl Terima:" & vnRcvPODate & Chr(10) & cbuQR_IDTerima & vnRcvPOHOID, vnSQLConn, vnSQLTrans) = True Then
                    vnBeginTrans = False
                    vnSQLTrans.Commit()
                    vnSQLTrans = Nothing

                    psPreview(vnRcvPOHOID, vnBrgCode, vnBrgCode & vbCrLf & vnRcvPONo & vbCrLf & vnRcvPODate, vnSQLConn)
                Else
                    LblMsgError.Text = "Print Gagal..." & vbCrLf & pbMsgError
                    LblMsgError.Visible = True

                    If vnBeginTrans Then
                        vnSQLTrans.Rollback()
                        vnSQLTrans = Nothing
                    End If
                End If
            Else
                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                psPreview(vnRcvPOHOID, vnBrgCode, vnBrgCode & vbCrLf & vnRcvPONo & vbCrLf & vnRcvPODate, vnSQLConn)
            End If

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

    Private Function fsGenBrgQRCode(vriRcvPOHOID As String, vriBarangCode As String, vriQRData As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

        Dim vnReturn As Boolean
        Try
            Dim vnQuery As String

            Dim vsIOFileStream As System.IO.FileStream
            Dim vsFileLength As Long
            Const csFileFormat = ".jpg"

            Dim vnCmd As SqlCommand
            Dim vnFileName As String
            Dim vnFileByte() As Byte

            vnFileName = vriBarangCode & "_" & Format(Date.Now, "yyyyMMdd_HHmmss") & "~sm" & csFileFormat

            Dim vnQRDir As String = ""

            pbuGenerateQRCode(vnFileName, vriQRData, vnQRDir)

            vsIOFileStream = System.IO.File.OpenRead(vnQRDir & vnFileName)

            vsFileLength = vsIOFileStream.Length
            ReDim vnFileByte(vsFileLength)

            vsIOFileStream.Read(vnFileByte, 0, vsFileLength)

            vnQuery = "Insert into Sys_SsoRcvPOBarangQRCode_TR"
            vnQuery += vbCrLf & "(RcvPOHOID,BRGCODE,BRGCODEQRCodeImg)"
            vnQuery += vbCrLf & "Values("
            vnQuery += vbCrLf & vriRcvPOHOID & ",'" & vriBarangCode & "',@vnBRGQRCodeImg"
            vnQuery += vbCrLf & ")"

            vnCmd = New SqlClient.SqlCommand(vnQuery, vriSQLConn, vriSQLTrans)
            vnCmd.Parameters.AddWithValue("@vnBRGQRCodeImg", vnFileByte)
            vnCmd.Transaction = vriSQLTrans
            vnCmd.ExecuteNonQuery()

            vnReturn = True

            Return vnReturn
        Catch ex As Exception
            pbMsgError = ex.Message
            Return False
        End Try
    End Function

    Private Sub psFillGrvOutstanding()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        GrvOutstanding.Visible = True
        GrvOutstanding.Enabled = True

        Dim vnQuery As String
        Dim vnDtb As New DataTable
        vnQuery = "	DECLARE @vriCompanyCode as varchar(20)	"
        vnQuery += vbCrLf & "	set @vriCompanyCode = '" & DstCompany.SelectedValue & "'	"
        vnQuery += vbCrLf & "	DECLARE @vriBrgCode_Lama as varchar(10)	"
        vnQuery += vbCrLf & "	set @vriBrgCode_Lama = '" & TxtListBrgCode.Text & "'	"
        vnQuery += vbCrLf & "	SELECT * FROM (	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Picklist'vTransType,trh.PCLNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPCLHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPCLReserve_TR trd with(nolock) on trd.PCLHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PCLCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Picking'vTransType,trh.PCKNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPCKHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPCKScan_TR trd with(nolock) on trd.RcvPOHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PCKCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.PCKScanDeleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Penerimaan'vTransType,trh.RcvPONo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoRcvPOHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoRcvPOScan_TR trd with(nolock) on trd.RcvPOHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.RcvPOCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.RcvPOScanDeleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway'vTransType,trh.PWNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPWHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPWScan1_TR trd with(nolock) on trd.PWHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PWCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.PWScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway Antar Wh'vTransType,trh.PYNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPYHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPYScan1_TR trd with(nolock) on trd.PYHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PYCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.PYScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway DO Titip'vTransType,trh.DTWNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoDTWHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoDTWScan1_TR trd with(nolock) on trd.DTWHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.DTWCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.DTWScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway DO Titip Antar Wh'vTransType,trh.DTYNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoDTYHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoDTYScan1_TR trd with(nolock) on trd.DTYHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.DTYCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.DTYScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway Void'vTransType,trh.PTVNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPTVHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPTVScan1_TR trd with(nolock) on trd.PTVHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PTVCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.PTVScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	UNION	"
        vnQuery += vbCrLf & "	Select trh.OID vTransOID,trh.TransCode,'Putaway Karantina'vTransType,trh.PTKNo vTransNo	"
        vnQuery += vbCrLf & "	       From Sys_SsoPTKHeader_TR trh with(nolock)	"
        vnQuery += vbCrLf & "	        inner join Sys_SsoPTKScan1_TR trd with(nolock) on trd.PTKHOID=trh.OID	"
        vnQuery += vbCrLf & "	  Where trh.PTKCompanyCode=@vriCompanyCode and trd.BRGCODE=@vriBrgCode_Lama and trd.PTKScan1Deleted=0 and	"
        vnQuery += vbCrLf & "	        trh.TransStatus>=0 and trh.TransStatus<18	"
        vnQuery += vbCrLf & "	) TX 	"


        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvOutstanding.DataSource = vnDtb
        GrvOutstanding.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psPreview(vriRcvPOHOID As String, vriBarangCode As String, vriLabel As String, vriSQLConn As SqlConnection)
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("~/Default.aspx", False)
        Dim vnCrpFileName As String = ""
        psGenerateCrp(vnCrpFileName, vriRcvPOHOID, vriBarangCode, vriLabel, vriSQLConn)

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
    Private Sub psGenerateCrp(ByRef vriCrpFileName As String, vriRcvPOHOID As String, vriBarangCode As String, vriLabel As String, vriSQLConn As SqlConnection)
        'Barcode print 2x (printer Barcode)
        Dim vnUserOID As String = Session("UserOID")
        Dim vnQuery As String
        Dim vn As Integer

        Dim vnCount As Integer = Val(TxtPrintCount.Text)
        vnCount = Math.Ceiling(Val(TxtPrintCount.Text) / 2)

        Dim vnSQLTrans As SqlTransaction = Nothing

        Try
            vnSQLTrans = vriSQLConn.BeginTransaction("inp")
            vnQuery = "Delete Sys_SsoPrintQRBarang_Temp Where UserOID=" & vnUserOID
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)

            For vn = 0 To vnCount - 2
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriLabel & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       1,'" & vriLabel & "',BRGCODEQRCodeImg"

                vnQuery += vbCrLf & "  From Sys_SsoRcvPOBarangQRCode_TR Where RcvPOHOID=" & vriRcvPOHOID & " and BRGCODE='" & vriBarangCode & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            Next

            vn = vn + 1

            If Val(vnCount) Mod 2 = 0 Then
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriLabel & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       1,'" & vriLabel & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "  From Sys_SsoRcvPOBarangQRCode_TR Where RcvPOHOID=" & vriRcvPOHOID & " and BRGCODE='" & vriBarangCode & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            Else
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,"
                vnQuery += vbCrLf & "vVisible011,BcProductGenCode011,BcProductGenCodeImg011,"
                vnQuery += vbCrLf & "vVisible012,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ","
                vnQuery += vbCrLf & "       1,'" & vriLabel & "',BRGCODEQRCodeImg,"
                vnQuery += vbCrLf & "       0,'" & vriLabel & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "  From Sys_SsoRcvPOBarangQRCode_TR Where RcvPOHOID=" & vriRcvPOHOID & " and BRGCODE='" & vriBarangCode & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            End If

            vriCrpFileName = stuSsoCrp.CrpBnsrphBarcodeSelectionQR

            vbuCrpQuery = "Select * From Sys_SsoPrintQRBarang_Temp with(nolock) Where UserOID=" & vnUserOID

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing

        Catch ex As Exception

            vnSQLTrans.Rollback()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
        End Try
    End Sub

    Protected Sub BtnPrintYes_Click(sender As Object, e As EventArgs) Handles BtnPrintYes.Click
        If Val(TxtPrintCount.Text) = 0 Then Exit Sub
        psPrintBrg()
        psShowPrint(False)
    End Sub

    Protected Sub BtnPrintNo_Click(sender As Object, e As EventArgs) Handles BtnPrintNo.Click
        psShowPrint(False)
    End Sub

    Private Sub GrvDetail_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles GrvDetail.RowCreated

    End Sub
End Class