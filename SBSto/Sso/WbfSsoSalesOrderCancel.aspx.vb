Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Public Class WbfSsoSalesOrderCancel
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSalesOrderCancel"
    Const csTNoPrefix = "VSO"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String
    Const csMaxByte = 1048576
    Enum ensColList
        OID = 0
    End Enum

    Enum ensColAttach
        OID = 0
        NotaPRIOImgNote = 1
        vUploadDatetime = 2
        vUploadDel = 3
    End Enum
    Enum ensColSO
        OID = 0
        CompanyCode = 1
        SalesOrderNo = 2
        vSalesOrderDate = 3
        vSUB = 4
        NAMA_CUSTOMER = 5
        ALAMAT = 6
        NAMA_KOTA = 7
    End Enum

    Private Sub psClearData()
        TxtTransID.Text = ""
        TxtTransStatus.Text = ""
        TxtSODate.Text = ""
        TxtSOVoidNo.Text = ""
        TxtSOVoidNote.Text = ""
        TxtSOCustCode.Text = ""
        TxtSOCustName.Text = ""
        TxtSONo.Text = ""
        TxtInvPRIOCancelNote.Text = ""

        TxtSOCustAddress.Text = ""
        TxtSOCustCity.Text = ""

        HdfTransStatus.Value = enuTCVOSO.Baru
    End Sub

    Private Sub psDefaultDisplay()
        DivList.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanList.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanConfirm.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsSO.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoVoidSO, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
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

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        vnQuery = "Select PM.OID,soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,"
        vnQuery += vbCrLf & "     soh.SUB +' '+ soh.NAMA_CUSTOMER vCustomer,"
        vnQuery += vbCrLf & "     PM.SOVoidNo,convert(varchar(11),PM.SOVoidDate,106)vSOVoidDate,"
        vnQuery += vbCrLf & "     PM.SOVoidNote,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.CreationDatetime,106)+' '+convert(varchar(5),PM.CreationDatetime,108)+' '+ CS.UserName vCreation,"
        vnQuery += vbCrLf & "     convert(varchar(11),PM.PreparedDatetime,106)+' '+convert(varchar(5),PM.PreparedDatetime,108)+' '+ PR.UserName vPrepared"

        vnQuery += vbCrLf & "From Sys_SsoSOrderVoidHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoSalesOrderHeader_TR soh with(nolock) on soh.OID=PM.SalesOrderHOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA CS with(nolock) on CS.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "     left outer join Sys_SsoUser_MA PR with(nolock) on PR.OID=PM.PreparedUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=soh.CompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"

        If Trim(TxtListSalesOrderNo.Text) <> "" Then
            vnQuery += vbCrLf & " and soh.SalesOrderNo like '%" & Trim(TxtListSalesOrderNo.Text) & "%'"
        End If
        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and PM.SOVoidDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and PM.SOVoidDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & "Order by soh.SalesOrderNo"
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

        If vriHOID = 0 Then
            vnQuery = "Select ''BRG,''NAMA_BARANG,0 Qty Where 1=2"
        Else
            vnQuery = "Select BRG,NAMA_BARANG,Qty"
            vnQuery += vbCrLf & " From Sys_SsoSalesOrderDetail_TR"
            vnQuery += vbCrLf & "Where SalesOrderHOID=" & vriHOID & " Order by BRG"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

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
            psButtonStatus()
        End If
    End Sub

    Private Sub psButtonShowList()
        BtnBaru.Enabled = False
        BtnEdit.Enabled = False

        BtnCancelVoidSO.Enabled = False
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
        If Not IsDate(TxtListStart.Text) Then
            TxtListStart.Text = Format(DateAdd(DateInterval.Day, -1, Date.Now), "dd MMM yyyy")
        End If
        If Not IsDate(TxtListEnd.Text) Then
            TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")
        End If
        psShowList(True)
    End Sub

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
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

        HdfActionStatus.Value = cbuActionNew
        psFillGrvDetail(0, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgSOVoidDate.Text = ""
        LblMsgSOVoidNote.Text = ""
        LblMsgSONo.Text = ""
        LblMsgError.Text = ""
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSOVoidDate.ReadOnly = Not vriBo
        TxtSOVoidNote.ReadOnly = Not vriBo
        psClearMessage()
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
        BtnEdit.Visible = Not vriBo
        BtnBaru.Visible = Not vriBo

        BtnCancelVoidSO.Visible = Not vriBo
        BtnPrepare.Visible = Not vriBo

        BtnPreview.Visible = Not vriBo
        BtnSONo.Visible = vriBo

        BtnList.Visible = Not vriBo
    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvDetail.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)
        If e.CommandName = "vAddItem" Then
        End If
    End Sub

    Private Sub psAddItem()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean = False

        Dim vnQuery As String
        Dim vnCriteria As String

        vnCriteria = " Where CompanyCode='" & DstCompany.SelectedValue & "' and NO_NOTA='" & TxtSONo.Text & "'"

        Dim vnDtb As New DataTable
        vnQuery = "Select KODE_BARANG,NAMA_BARANG,HARGA,QTY,QTYBONUS,SATUAN,NO_REF,SALESMAN"
        vnQuery += vbCrLf & "  From Sys_DcmJUAL"

        vnQuery += vbCrLf & vnCriteria
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        Try
            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            Dim vnDRow As DataRow
            For vn = 0 To vnDtb.Rows.Count - 1
                vnDRow = vnDtb.Rows(vn)
                vnQuery = "Select count(1) From Sys_DcmSJDetail_TR with(nolock) where DcmSJHOID=" & TxtTransID.Text & " and KodeBarang='" & vnDRow.Item("KODE_BARANG") & "' and NamaBarang='" & fbuFormatString(vnDRow.Item("NAMA_BARANG")) & "' and Price=" & vnDRow.Item("HARGA")
                If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                    vnQuery = "Insert into Sys_DcmSJDetail_TR"
                    vnQuery += vbCrLf & "(DcmSJHOID,"
                    vnQuery += vbCrLf & "KodeBarang,NamaBarang,"
                    vnQuery += vbCrLf & "Price,Qty,QtyBonus,"
                    vnQuery += vbCrLf & "Satuan,NoRef,Salesman,"
                    vnQuery += vbCrLf & "DcmSJDQty,"
                    vnQuery += vbCrLf & "DcmSJDNote)"
                    vnQuery += vbCrLf & "values(" & TxtTransID.Text & ","

                    vnQuery += vbCrLf & "'" & vnDRow.Item("KODE_BARANG") & "','" & fbuFormatString(vnDRow.Item("NAMA_BARANG")) & "',"
                    vnQuery += vbCrLf & "" & vnDRow.Item("HARGA") & "," & vnDRow.Item("QTY") & "," & vnDRow.Item("QTYBONUS") & ","
                    vnQuery += vbCrLf & "'" & vnDRow.Item("SATUAN") & "','" & fbuFormatString(vnDRow.Item("NO_REF")) & "','" & fbuFormatString(vnDRow.Item("SALESMAN")) & "',"

                    vnQuery += vbCrLf & "0,"
                    vnQuery += vbCrLf & "'')"
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                End If
            Next

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psFillGrvDetail(TxtSOOID.Text, vnSQLConn)

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

    Private Sub psShowConfirm(vriBo As Boolean)
        If vriBo Then
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivConfirm.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
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

            psFillGrvDetail(0, vnSQLConn)

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

            psDisplayData(vnSQLConn)

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

        vnQuery = "Select PM.*,convert(varchar(11),PM.SOVoidDate,106)vSOVoidDate,"
        vnQuery += vbCrLf & "soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,"
        vnQuery += vbCrLf & "soh.SUB,soh.NAMA_CUSTOMER,soh.ALAMAT,soh.NAMA_KOTA,"
        vnQuery += vbCrLf & "ST.TransStatusDescr"
        vnQuery += vbCrLf & "From Sys_SsoSOrderVoidHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoSalesOrderHeader_TR soh with(nolock) on soh.OID=PM.SalesOrderHOID"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     Where PM.OID=" & TxtTransID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtSOOID.Text = vnDtb.Rows(0).Item("SalesOrderHOID")
            TxtSODate.Text = vnDtb.Rows(0).Item("vSalesOrderDate")
            TxtSONo.Text = vnDtb.Rows(0).Item("SalesOrderNo")
            TxtSOVoidNo.Text = vnDtb.Rows(0).Item("SOVoidNo")

            TxtSOCustCode.Text = vnDtb.Rows(0).Item("SUB")
            TxtSOCustName.Text = vnDtb.Rows(0).Item("NAMA_CUSTOMER")
            TxtSOCustAddress.Text = vnDtb.Rows(0).Item("ALAMAT")
            TxtSOCustCity.Text = vnDtb.Rows(0).Item("NAMA_KOTA")

            TxtSOVoidDate.Text = vnDtb.Rows(0).Item("vSOVoidDate")
            TxtSOVoidNote.Text = vnDtb.Rows(0).Item("SOVoidNote")
            TxtInvPRIOCancelNote.Text = fbuValStr(vnDtb.Rows(0).Item("SOVoidCancelNote"))

            TxtTransStatus.Text = vnDtb.Rows(0).Item("TransStatusDescr")

            DstCompany.SelectedValue = Trim(vnDtb.Rows(0).Item("CompanyCode"))

            HdfTransStatus.Value = vnDtb.Rows(0).Item("TransStatus")

            psButtonStatus()
        End If

        psFillGrvDetail(Val(TxtSOOID.Text), vriSQLConn)

        vnDtb.Dispose()
    End Sub

    Private Sub psButtonVisible()
        BtnBaru.Visible = BtnBaru.Enabled
        BtnEdit.Visible = BtnEdit.Enabled

        BtnCancelVoidSO.Visible = BtnCancelVoidSO.Enabled
        BtnPrepare.Visible = BtnPrepare.Enabled

        BtnPreview.Visible = BtnPreview.Enabled
        BtnList.Visible = BtnList.Enabled
    End Sub
    Private Sub psButtonStatusDefault()
        BtnBaru.Enabled = True
        BtnEdit.Enabled = False

        BtnCancelVoidSO.Enabled = False
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
            BtnEdit.Enabled = (HdfTransStatus.Value = enuTCVOSO.Baru)

            BtnCancelVoidSO.Enabled = (HdfTransStatus.Value = enuTCVOSO.Baru Or HdfTransStatus.Value = enuTCVOSO.Prepared)

            BtnPrepare.Enabled = (HdfTransStatus.Value = enuTCVOSO.Baru)
            BtnPreview.Enabled = (HdfTransStatus.Value >= enuTCVOSO.Prepared)

            If HdfTransStatus.Value = enuTCVOSO.Baru Then
                BtnPrepare.Text = "Prepare"
            ElseIf HdfTransStatus.Value = enuTCVOSO.Prepared Then
                BtnPrepare.Text = "Prepared"
            End If

            psButtonVisible()
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If Trim(TxtTransID.Text) = "" Then Exit Sub

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Session(csModuleName & stuSession.Simpan) = ""
        HdfActionStatus.Value = cbuActionEdit
        psFillGrvDetail(TxtSOOID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        psEnableInput(True)
        psEnableSave(True)

        BtnSONo.Visible = False
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        If HdfTransStatus.Value = enuTCVOSO.Baru Then
            psSaveBaru()
        End If
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
        If Trim(TxtSONo.Text) = "" Then
            LblMsgSONo.Text = "Isi Nomor SO"
            vnSave = False
        End If
        If Trim(TxtSOVoidNo.Text) = "" Then
            LblMsgSONo.Text = "Isi Nomor Void SO"
            vnSave = False
        End If
        If Not IsDate(Trim(TxtSOVoidDate.Text)) Then
            LblMsgSOVoidDate.Text = "Isi Tanggal"
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
            Dim vnUserNIP As String = Session("EmpNIP")

            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(1) From Sys_SsoSOrderVoidHeader_TR Where SalesOrderHOID=" & TxtSOOID.Text & " and TransStatus<>" & enuTCVOSO.Cancelled
                If fbuGetDataStrSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgSONo.Text = "SO " & TxtSONo.Text & " Sudah di Void"
                    vnSave = False
                End If
                vnQuery = "Select count(1) From Sys_SsoSOrderVoidHeader_TR Where SOVoidNo='" & Trim(TxtSOVoidNo.Text) & "' and TransStatus<>" & enuTCVOSO.Cancelled
                If fbuGetDataStrSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgSONo.Text = "Nomor Void " & TxtSOVoidNo.Text & " Sudah Ada"
                    vnSave = False
                End If

                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If
                Dim vnCompanyCode As String = Trim(DstCompany.SelectedValue)

                Dim vnOID As Integer
                vnQuery = "Select max(OID) from Sys_SsoSOrderVoidHeader_TR"
                vnOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoSOrderVoidHeader_TR("
                vnQuery += vbCrLf & "OID,SalesOrderHOID,SOVoidDate,SOVoidNote,SOVoidNo,"
                vnQuery += vbCrLf & "TransCode,"
                vnQuery += vbCrLf & "CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "values(" & vnOID & "," & Val(TxtSOOID.Text) & ",'" & TxtSOVoidDate.Text & "',"
                vnQuery += vbCrLf & "'" & fbuFormatString(Trim(TxtSOVoidNote.Text)) & "','" & Trim(TxtSOVoidNo.Text) & "',"
                vnQuery += vbCrLf & "'" & stuTransCode.SsoVoidSO & "',"
                vnQuery += vbCrLf & Session("UserOID") & ",getdate())"
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

                pbuInsertStatusVoidSO(vnOID, enuTCVOSO.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"

                TxtTransID.Text = vnOID

                HdfTransStatus.Value = enuTCVOSO.Baru
            Else
                vnQuery = "Select count(1) From Sys_SsoSOrderVoidHeader_TR Where SOVoidNo='" & Trim(TxtSOVoidNo.Text) & "' and TransStatus<>" & enuTCVOSO.Cancelled & " and OID=" & TxtTransID.Text
                If fbuGetDataStrSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgSONo.Text = "Nomor Void " & TxtSOVoidNo.Text & " Sudah Ada"
                    vnSave = False
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()
                vnBeginTrans = True

                vnQuery = "Update Sys_SsoSOrderVoidHeader_TR set"
                vnQuery += vbCrLf & "SOVoidDate='" & Trim(TxtSOVoidDate.Text) & "',"
                vnQuery += vbCrLf & "SOVoidNo='" & Trim(TxtSOVoidNo.Text) & "',"
                vnQuery += vbCrLf & "SOVoidNote='" & fbuFormatString(Trim(TxtSOVoidNote.Text)) & "',"
                vnQuery += vbCrLf & "ModificationUserOID=" & Session("UserOID") & ",ModificationDatetime=getdate()"
                vnQuery += vbCrLf & "Where OID=" & TxtTransID.Text
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                pbuInsertStatusVoidSO(TxtTransID.Text, enuTCVOSO.Baru, Session("UserOID"), vnSQLConn, vnSQLTrans)

                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                Session(csModuleName & stuSession.Simpan) = "Done"
            End If

            psEnableInput(False)
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

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "SalesOrderNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            If vnIdx >= GrvList.Rows.Count Then Exit Sub

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

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        If HdfProcess.Value = "CancelVoidSO" Then
            If Trim(TxtConfirmNote.Text) = "" Then
                LblConfirmWarning.Text = "Isi Note untuk Cancel"
                Exit Sub
            End If
            psCancelVoidSO()
        ElseIf HdfProcess.Value = "PrepareVoidSO" Then
            psPrepareVoidSO()
            psButtonStatus()
            psShowConfirm(False)
        End If
    End Sub

    Private Sub BtnPrepare_Click(sender As Object, e As EventArgs) Handles BtnPrepare.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Prepare) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        LblConfirmMessage.Text = "Anda Prepare Void SO ?<br />WARNING : Prepare Tidak Dapat Dibatalkan"
        HdfProcess.Value = "PrepareVoidSO"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = False
        LblConfirmError.Text = ""

        psShowConfirm(True)
    End Sub

    Private Sub psCancelVoidSO()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psCancelVoidSO", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        LblConfirmError.Text = ""
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
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0")

            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoSOrderVoidHeader_TR set TransStatus=" & enuTCVOSO.Cancelled & ",SOVoidCancelNote='" & fbuFormatString(Trim(TxtConfirmNote.Text)) & "',"
            vnQuery += vbCrLf & "CancelledUserOID=" & Session("UserOID") & ",CancelledDatetime=getdate() Where OID=" & TxtTransID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2.2")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoSalesOrderHeader_TR Set SOVoid=0,SOVoidDatetime=Null,SOVoidNo='',SOVoidNote='' Where OID=" & TxtSOOID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQLTrans(vnDtb, vnQuery, vnSQLConn, vnSQLTrans)

            vnQuery = "Update Sys_SsoSalesOrder Set SOVoid=0,SOVoidDatetime=Null,SOVoidNo='',SOVoidNote='' Where CompanyCode='" & DstCompany.SelectedValue & "' and SalesOrderHOID=" & TxtSOOID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("4")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQLTrans(vnDtb, vnQuery, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("5")
            pbuInsertStatusVoidSO(TxtTransID.Text, enuTCVOSO.Cancelled, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("6")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Batal Invoice Prioritas Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            psDisplayData(vnSQLConn)

            psButtonStatus()
            psShowConfirm(False)

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

    Private Sub psPrepareVoidSO()
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psPrepareVoidSO", TxtTransID.Text, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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
            Dim vnHOID As String = TxtTransID.Text
            Dim vnInvOID As String = TxtSOOID.Text

            Dim vnQuery As String
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("0")
            Dim vnDtb As New DataTable

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoSOrderVoidHeader_TR set TransStatus=" & enuTCVOSO.Prepared & ",PreparedUserOID=" & Session("UserOID") & ",PreparedDatetime=getdate() Where OID=" & vnHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Declare @vnGetDate datetime"
            vnQuery += vbCrLf & "Set @vnGetDate=getdate()"
            vnQuery += vbCrLf & "Update Sys_SsoSalesOrderHeader_TR Set SOVoid=1,SOVoidDatetime=@vnGetDate,SOVoidNo='" & TxtSOVoidNo.Text & "',SOVoidNote='" & fbuFormatString(TxtSOVoidNote.Text) & "' Where OID=" & vnInvOID
            vnQuery += vbCrLf & "Update Sys_SsoSalesOrder Set SOVoid=1,SOVoidDatetime=@vnGetDate,SOVoidNo='" & TxtSOVoidNo.Text & "',SOVoidNote='" & fbuFormatString(TxtSOVoidNote.Text) & "' Where CompanyCode='" & DstCompany.SelectedValue & "' and SalesOrderHOID=" & TxtSOOID.Text
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("3")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQLTrans(vnDtb, vnQuery, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("5")
            pbuInsertStatusVoidSO(vnHOID, enuTCVOSO.Prepared, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("6")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Prepare Invoice Prioritas Sukses")
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

    Protected Sub BtnCancelVoidSO_Click(sender As Object, e As EventArgs) Handles BtnCancelVoidSO.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Cancel_Trans) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        LblConfirmMessage.Text = "Anda Membatalkan Void SO ?<br />WARNING : Batal Tidak Dapat Dibatalkan"
        HdfProcess.Value = "CancelVoidSO"
        LblConfirmWarning.Text = ""
        TxtConfirmNote.Text = ""
        tbConfirmNote.Visible = True
        LblConfirmError.Text = ""
        psShowConfirm(True)
    End Sub

    Private Sub psGenerateCrp(ByRef vriCrpFileName As String)
        'vriCrpFileName = stuDcmCrpName.CrpDcmNotaCAS

        'vbuCrpQuery = "Select * From fnTbl_DcmNotaCAS(" & TxtTransID.Text & ",'" & Session("UserID") & "')"
    End Sub

    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
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

    Private Sub psShowSO(vriBo As Boolean)
        If vriBo Then
            DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "visible"

            TxtLsSONo.Focus()
        Else
            DivLsSO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
        BtnList.Enabled = Not vriBo
    End Sub
    Protected Sub BtnLsSOFind_Click(sender As Object, e As EventArgs) Handles BtnLsSOFind.Click
        If DstCompany.SelectedIndex = 0 And Trim(TxtLsSOCustomer.Text) = "" And Trim(TxtLsSONo.Text) = "" And IsDate(TxtLsSOStart.Text) = False And IsDate(TxtLsSOEnd.Text) = False Then
            LblMsgLsSONo.Text = "Pilih Company, Nomor SO, Customer atau Periode Tanggal SO"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvLsSO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvLsSO(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String = ""

        vnCriteria = "      Where soh.SOVoid=0 and not soh.OID in(Select b.SalesOrderHOID From Sys_SsoSOrderVoidHeader_TR b Where b.TransStatus<>" & enuTCVOSO.Cancelled & ")"

        If DstCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and soh.CompanyCode='" & DstCompany.SelectedValue & "'"
        End If
        If Trim(TxtLsSOCustomer.Text) <> "" Then
            vnCriteria += vbCrLf & "            and soh.NAMA_CUSTOMER like '%" & fbuFormatString(Trim(TxtLsSOCustomer.Text)) & "%'"
        End If
        If Trim(TxtLsSONo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderNo like '%" & fbuFormatString(Trim(TxtLsSONo.Text)) & "%'"
        End If
        If IsDate(TxtLsSOStart.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate >= '" & TxtLsSOStart.Text & "'"
        End If
        If IsDate(TxtLsSOEnd.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate <= '" & TxtLsSOEnd.Text & "'"
        End If

        vnQuery = "Select soh.OID,soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,soh.SUB vSUB,soh.NAMA_CUSTOMER,ALAMAT,soh.NAMA_KOTA"
        vnQuery += vbCrLf & "       From Sys_SsoSalesOrderHeader_TR soh"

        If vnUserCompanyCode <> "" And DstCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=soh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by soh.CompanyCode,soh.SalesOrderNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvLsSO.DataSource = vnDtb
        GrvLsSO.DataBind()

        TxtLsSONo.Focus()
    End Sub

    Private Sub GrvLsSO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsSO.PageIndexChanging
        GrvLsSO.PageIndex = e.NewPageIndex
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvLsSO(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvLsSO_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsSO.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If e.CommandName = "SalesOrderNo" Then
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            If vnIdx >= GrvLsSO.Rows.Count Then Exit Sub

            Dim vnGRowList As GridViewRow = GrvLsSO.Rows(vnIdx)

            Dim vnCompanyCode As String = vnGRowList.Cells(ensColSO.CompanyCode).Text
            Dim vnSalesOrderNo As String = DirectCast(vnGRowList.Cells(ensColSO.SalesOrderNo).Controls(0), LinkButton).Text

            TxtSONo.Text = vnSalesOrderNo
            TxtSOCustAddress.Text = vnGRowList.Cells(ensColSO.ALAMAT).Text
            TxtSOCustCity.Text = vnGRowList.Cells(ensColSO.NAMA_KOTA).Text
            TxtSOCustCode.Text = vnGRowList.Cells(ensColSO.vSUB).Text
            TxtSOCustName.Text = vnGRowList.Cells(ensColSO.NAMA_CUSTOMER).Text
            TxtSODate.Text = vnGRowList.Cells(ensColSO.vSalesOrderDate).Text
            TxtSOOID.Text = vnGRowList.Cells(ensColSO.OID).Text

            psShowSO(False)
        End If
    End Sub

    Protected Sub BtnLsSOClose_Click(sender As Object, e As EventArgs) Handles BtnLsSOClose.Click
        psShowSO(False)
    End Sub

    Private Sub BtnSONo_Click(sender As Object, e As EventArgs) Handles BtnSONo.Click
        psShowSO(True)
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
        If Not IsNumeric(TxtTransID.Text) Then Exit Sub
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

        Dim vnName1 As String = "Preview"
        Dim vnType As Type = Me.GetType()
        Dim vnClientScript As ClientScriptManager = Page.ClientScript
        If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
            Dim vnParam As String
            vnParam = "vqTrOID=" & TxtTransID.Text
            vnParam += "&vqTrCode=" & stuTransCode.SsoVoidSO
            vnParam += "&vqTrNo=" & TxtSONo.Text

            vbuPreviewOnClose = "0"

            ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
            psShowPreview(True)
        End If
    End Sub

    Protected Sub GrvSO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsSO.SelectedIndexChanged

    End Sub
End Class