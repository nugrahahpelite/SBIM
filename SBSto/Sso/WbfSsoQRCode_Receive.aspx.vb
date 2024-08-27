Imports System.Data.SqlClient
Imports Spire.Barcode
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.IO
Public Class WbfSsoQRCode_Receive
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoQRCode_Receive"

    Dim settings As BarcodeSettings

    Dim vsIOFileStream As System.IO.FileStream
    Dim vsFileLength As Long

    Dim vsQRDir As String

    Const csFileFormat = ".jpg"

    Enum ensColLsRcvPO
        RcvPONo = 0
        vRcvPODate = 1
        RcvPOSupplierName = 2
        RcvPOTypeName = 3
        OID = 4
        RcvPORefTypeOID = 5
        RcvRefTypeOID = 6
        RcvPORefOID = 7
    End Enum

    Enum ensColLsBrg
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
    End Enum

    Private Sub psClearData()
        TxtBrgCode.Text = ""
        TxtBrgName.Text = ""
        TxtBrgUnit.Text = ""
        TxtPrintCount.Text = ""
        TxtPrintNote.Text = ""
        TxtOID.Text = ""
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Visible = False
        LblMsgBrgName.Visible = False
        LblMsgPrintCount.Visible = False
        LblMsgPrintNote.Visible = False
        LblMsgRcvPO.Visible = False
        LblMsgWarehouse.Visible = False
        LblMsgError.Visible = False
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPrintCount.ReadOnly = Not vriBo
        TxtPrintNote.ReadOnly = Not vriBo
        BtnBrgCode.Enabled = vriBo
        BtnBrgCode.Visible = BtnBrgCode.Enabled
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnPrint.Visible = vriBo
        BtnFind.Enabled = Not vriBo
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Master"
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPrintQRBarang, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, True, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, True, vnSQLConn)
            End If

            '<--- 21 Feb 2023 disederhanakan Agus
            'If Session("UserWarehouseCode") = "" Then
            '    pbuFillDstWarehouse(DstWarehouse, False, vnSQLConn)
            'Else
            '    pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstWarehouse, False, vnSQLConn)
            'End If
            pbuFillDstWarehouse_ByUserOID(Session("UserOID"), DstWarehouse, False, vnSQLConn)
            '<<==21 Feb 2023 disederhanakan Agus

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psDefaultDisplay()
        DivLsRcvPO.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsRcvPO.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPreview.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub

    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
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

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.OID,PM.CompanyCode,WM.WarehouseName,RC.RcvPONo,Convert(varchar(11),RC.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "      PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,PM.PrintCount,PM.PrintNote,PrintDatetime,"
        vnQuery += vbCrLf & "      GM.UserName vPrintUserName"
        vnQuery += vbCrLf & " From Sys_SsoRcvPOPrintQRCode_TR PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR RC with(nolock) on RC.OID=PM.RcvPOHOID"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.PrintWarehouseOID"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA GM on GM.OID=PM.PrintUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.SOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=PM.PrintWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "Where 1=1"
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If

        If Trim(TxtListPenerimaan.Text) <> "" Then
            vnQuery += vbCrLf & "     and RC.RcvPONo like '%" & fbuFormatString(Trim(TxtListPenerimaan.Text)) & "%'"
        End If

        If Trim(TxtListBarang.Text) <> "" Then
            vnQuery += vbCrLf & "     and (PM.BRGCODE like '%" & fbuFormatString(Trim(TxtListBarang.Text)) & "%' OR PM.BRGNAME like '%" & fbuFormatString(Trim(TxtListBarang.Text)) & "%')"
        End If

        vnQuery += vbCrLf & "Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If
        If Val(HdfRcvPOHOID.Value) = 0 Then
            LblMsgRcvPO.Text = "Isi Penerimaan"
            LblMsgRcvPO.Visible = True
            Exit Sub
        End If

        psFillGrvLsBrg()
    End Sub

    Protected Sub BtnLsBrgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBrgClose.Click
        psShowLsBrg(False)
    End Sub

    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Sub psShowLsRcvPO(vriBo As Boolean)
        If vriBo Then
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsRcvPO.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psFillGrvLsRcvPO()
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
        vnQuery = "Select PM.RcvPONo,Convert(varchar(11),PM.RcvPODate)vRcvPODate,PM.RcvPOSupplierName,PT.RcvPOTypeName,PM.OID,isnull(PM.RcvPORefTypeOID,999)RcvPORefTypeOID,PM.RcvRefTypeOID,PM.RcvPORefOID"
        vnQuery += vbCrLf & " From Sys_SsoRcvPOHeader_TR PM"
        vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOType_MA PT on PT.OID=PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & "Where (PM.RcvRefTypeOID=" & enuRcvType.Pembelian & " or PM.RcvRefTypeOID=" & enuRcvType.Lain_lain & " or PM.RcvRefTypeOID=" & enuRcvType.Karantina & ") and"
        vnQuery += vbCrLf & "      PM.RcvPOCompanyCode='" & DstCompany.SelectedValue & "' and PM.RcvPONo like '%" & Trim(TxtLsRcvPONo.Text) & "%'"
        vnQuery += vbCrLf & " Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsRcvPO.DataSource = vnDtb
        GrvLsRcvPO.DataBind()

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
        vnQuery += vbCrLf & "Where CompanyCode='" & DstCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtLsBrg.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtLsBrg.Text) & "%')"

        If HdfRcvRefTypeOID.Value = enuRcvType.Lain_lain Then
            vnQuery += vbCrLf & "      and PM.BRGCODE in(Select b.BRGCODE From Sys_SsoRcvMscDetail_TR b with(nolock) Where RcvMscHOID=" & HdfRcvPORefOID.Value & ")"
        ElseIf HdfRcvRefTypeOID.Value = enuRcvType.Karantina Then
            vnQuery += vbCrLf & "      and PM.BRGCODE in(Select b.BRGCODE From Sys_SsoRcvKRDetail_TR b with(nolock) Where RcvKRHOID=" & HdfRcvPORefOID.Value & ")"
        End If

        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsBrg.DataSource = vnDtb
        GrvLsBrg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvLsBrg_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsBrg.PageIndexChanging
        GrvLsBrg.PageIndex = e.NewPageIndex
        psFillGrvLsBrg()
    End Sub

    Private Sub GrvLsBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsBrg.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvLsBrg.Rows(vnIdx)
            vnValue = DirectCast(vnGRow.Cells(ensColLsBrg.BRGCODE).Controls(0), LinkButton).Text
            TxtBrgCode.Text = vnValue
            TxtBrgName.Text = vnGRow.Cells(ensColLsBrg.BRGNAME).Text
            TxtBrgUnit.Text = vnGRow.Cells(ensColLsBrg.BRGUNIT).Text
            psShowLsBrg(False)
        End If
    End Sub

    Protected Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

        Dim vnSave As Boolean = True
        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
        End If
        If DstWarehouse.SelectedValue = "0" Then
            LblMsgWarehouse.Text = "Pilih Warehouse"
            LblMsgWarehouse.Visible = True
        End If
        If Len(Trim(TxtRcvPONo.Text)) = 0 Then
            LblMsgRcvPO.Text = "Isi Penerimaan"
            LblMsgRcvPO.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtBrgCode.Text)) = 0 Then
            LblMsgBrgName.Text = "Pilih Barang"
            LblMsgBrgName.Visible = True
            vnSave = False
        End If
        If Val(Trim(TxtPrintCount.Text)) = 0 Then
            LblMsgPrintCount.Text = "Isi Jumlah"
            LblMsgPrintCount.Visible = True
            vnSave = False
        End If
        If Not vnSave Then
            Exit Sub
        End If
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
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
        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            Dim vnHOID As Integer
            Dim vnQuery As String

            If HdfRcvPORefTypeOID.Value = enuRcvPOType.Import Then
                vnQuery = "Select count(1) From Sys_SsoPLDetail_TR with(nolock) Where PLHOID=(Select rcv.RcvPORefOID From Sys_SsoRcvPOHeader_TR rcv with(nolock) Where rcv.OID=" & HdfRcvPOHOID.Value & ") and BRGCODE='" & TxtBrgCode.Text & "'"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) = 0 Then
                    LblMsgBrgName.Text = TxtBrgCode.Text & " TIDAK ADA DI " & TxtRcvPONo.Text
                    LblMsgBrgName.Visible = True
                    Exit Sub
                End If
            End If

            Dim vnUserOID As String = Session("UserOID")
            Dim vnCompanyCode As String = DstCompany.SelectedValue

            vnQuery = "Select isnull(max(OID),0)+1 from Sys_SsoRcvPOPrintQRCode_TR"
            vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True
            vnQuery = "Insert into Sys_SsoRcvPOPrintQRCode_TR("
            vnQuery += vbCrLf & "OID,CompanyCode,PrintWarehouseOID,"
            vnQuery += vbCrLf & "RcvPOHOID,"
            vnQuery += vbCrLf & "BRGCODE,BRGNAME,BRGUNIT,"
            vnQuery += vbCrLf & "PrintCount,PrintNote,"
            vnQuery += vbCrLf & "PrintDatetime,PrintUserOID)"
            vnQuery += vbCrLf & "values("
            vnQuery += vbCrLf & vnHOID & ",'" & vnCompanyCode & "'," & DstWarehouse.SelectedValue & ","
            vnQuery += vbCrLf & HdfRcvPOHOID.Value & ","
            vnQuery += vbCrLf & "'" & fbuFormatString(TxtBrgCode.Text) & "','" & fbuFormatString(TxtBrgName.Text) & "','" & fbuFormatString(TxtBrgUnit.Text) & "',"
            vnQuery += vbCrLf & TxtPrintCount.Text & ",'" & fbuFormatString(Trim(TxtPrintNote.Text)) & "',"
            vnQuery += vbCrLf & "getdate()," & vnUserOID & ")"
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnQuery = "Select count(1) From Sys_SsoRcvPOBarangQRCode_TR Where RcvPOHOID=" & HdfRcvPOHOID.Value & " and BRGCODE='" & TxtBrgCode.Text & "'"
            If fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) = 0 Then
                If fsGenBrgQRCode(HdfRcvPOHOID.Value, TxtBrgCode.Text, TxtBrgCode.Text & Space(5) & Chr(10) & TxtBrgName.Text & Chr(10) & "No.Terima:" & TxtRcvPONo.Text & Chr(10) & "Tgl Terima:" & TxtRcvPODate.Text & Chr(10) & cbuQR_IDTerima & HdfRcvPOHOID.Value, vnSQLConn, vnSQLTrans) = True Then
                    vnBeginTrans = False
                    vnSQLTrans.Commit()
                    vnSQLTrans = Nothing

                    TxtOID.Text = vnHOID

                    psPreview(HdfRcvPOHOID.Value, TxtBrgCode.Text, TxtBrgCode.Text & vbCrLf & TxtRcvPONo.Text & vbCrLf & TxtRcvPODate.Text, vnSQLConn)
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

                psPreview(HdfRcvPOHOID.Value, TxtBrgCode.Text, TxtBrgCode.Text & vbCrLf & TxtRcvPONo.Text & vbCrLf & TxtRcvPODate.Text, vnSQLConn)
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
            If Val(TxtPrintCount.Text) Mod 2 = 0 Then
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
    Private Sub psGenerateCrp_20230103_Orig_Bef_Ganjil(ByRef vriCrpFileName As String, vriBarangCode As String, vriSQLConn As SqlConnection)
        'Barcode print 2x (printer Barcode)
        Dim vnUserOID As String = Session("UserOID")
        Dim vnQuery As String
        Dim vn As Integer
        Dim vnCount As Integer = Val(TxtPrintCount.Text)

        Dim vnSQLTrans As SqlTransaction = Nothing

        Try
            vnSQLTrans = vriSQLConn.BeginTransaction("inp")
            vnQuery = "Delete Sys_SsoPrintQRBarang_Temp Where UserOID=" & vnUserOID
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)

            For vn = 0 To vnCount - 1
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,BcProductGenCode011,BcProductGenCodeImg011,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ",'" & vriBarangCode & "',BRGCODEQRCodeImg,'" & vriBarangCode & "',BRGCODEQRCodeImg"
                vnQuery += vbCrLf & "       From " & fbuGetDBMaster() & "Sys_MstBarangQRCode_MA Where BRGCODE='" & vriBarangCode & "'"
                pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)
            Next

            vriCrpFileName = stuSsoCrp.CrpBnsrphBarcodeSelectionQR

            vbuCrpQuery = "Select * From Sys_SsoPrintQRBarang_Temp Where UserOID=" & vnUserOID

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing

        Catch ex As Exception
            vnSQLTrans.Rollback()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
        End Try
    End Sub

    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub

    Private Function fsGenBrgQRCode(vriRcvPOHOID As String, vriBarangCode As String, vriQRData As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
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

    Private Sub psGenerateSN(vriHOID As String, vriGdgCode As String, vriGdgName As String, vriBrgCode As String, vriBrgName As String, vriYMD As String, vriSNCount As Integer, vriSNStart As Integer, vriSNEnd As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        Dim vnSerialNo As String
        Dim vnQRCode As String
        Dim vn As Integer
        Dim vnCmd As SqlCommand

        Dim vnFileName As String
        Dim vnFileByte() As Byte

        For vn = vriSNStart To vriSNEnd
            vnSerialNo = vriBrgCode & vriYMD & Format(vn, "00000#")
            vnQRCode = vnSerialNo & Chr(10) & vriBrgCode & Chr(10) & vriBrgName & Chr(10) & "Gudang : " & vriGdgCode & " - " & vriGdgName

            vnFileName = Format(Date.Now, "yyyyMMdd_HHmmss") & "_" & vnSerialNo & "~sm" & csFileFormat

            psGenerateBarCode(vnFileName, vnQRCode)

            vsIOFileStream = System.IO.File.OpenRead(vsQRDir & vnFileName)

            vsFileLength = vsIOFileStream.Length
            ReDim vnFileByte(vsFileLength)

            vsIOFileStream.Read(vnFileByte, 0, vsFileLength)

            vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBrgQRGenData_TR"
            vnQuery += vbCrLf & "(QRGenHOID,BRGCODE,BRGSerialNo,BRGQRCode,BRGQRCodeImg)"
            vnQuery += vbCrLf & "Values("
            vnQuery += vbCrLf & vriHOID & ",'" & vriBrgCode & "','" & vnSerialNo & "','" & vnQRCode & "',@vnBRGQRCodeImg"
            vnQuery += vbCrLf & ")"
            'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnCmd = New SqlClient.SqlCommand(vnQuery, vriSQLConn)
            vnCmd.Parameters.AddWithValue("@vnBRGQRCodeImg", vnFileByte)
            vnCmd.Transaction = vriSQLTrans
            vnCmd.ExecuteNonQuery()

        Next
    End Sub

    Private Sub psGenerateBarCode(vriFileName As String, vriBrgQRCode As String)
        Dim vnBrgQRCode As String
        vnBrgQRCode = vriBrgQRCode

        'set the configuration of barcode
        settings = New BarcodeSettings()
        Dim data As String = vnBrgQRCode
        'Dim type As String = "Code128"
        Dim type As String = "QRCode"

        settings.Data2D = data
        settings.Data = vnBrgQRCode

        settings.Type = CType(System.Enum.Parse(GetType(BarCodeType), type), BarCodeType)
        settings.HasBorder = True
        settings.BorderDashStyle = CType(System.Enum.Parse(GetType(DashStyle), "Solid"), DashStyle)

        Dim fontSize As Short = 12
        Dim font As String = "Arial"

        settings.TextFont = New Font(font, fontSize, FontStyle.Bold)

        Dim barHeight As Short = 15

        settings.BarHeight = barHeight

        'settings.X = 1.9
        'settings.Y = 1.9

        settings.ShowText = False
        settings.ShowTextOnBottom = True
        settings.BorderColor = Color.White

        settings.ShowCheckSumChar = True

        'generate the barcode use the settings
        Dim generator As New BarCodeGenerator(settings)
        Dim barcode As Image = generator.GenerateImage()

        vsQRDir = Server.MapPath("~") & "\QRDir\"

        If Dir(vsQRDir & vriFileName) = "" Then
            barcode.Save(vsQRDir & vriFileName)
        End If
    End Sub

    Protected Sub GrvLsBrg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsBrg.SelectedIndexChanged

    End Sub

    Protected Sub BtnBrgCode_Click(sender As Object, e As EventArgs) Handles BtnBrgCode.Click
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psShowLsBrg(True)
    End Sub

    Private Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Protected Sub BtnLsRcvPOClose_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOClose.Click
        psShowLsRcvPO(False)
    End Sub

    Protected Sub BtnRcvPO_Click(sender As Object, e As EventArgs) Handles BtnRcvPO.Click
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        TxtBrgCode.Text = ""
        TxtBrgName.Text = ""
        TxtBrgUnit.Text = ""

        psShowLsRcvPO(True)
    End Sub

    Private Sub GrvLsRcvPO_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsRcvPO.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsRcvPO.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(ensColLsRcvPO.RcvPONo).Controls(0), LinkButton).Text
            TxtRcvPONo.Text = vnValue
            TxtRcvPODate.Text = vnRow.Cells(ensColLsRcvPO.vRcvPODate).Text
            HdfRcvPOHOID.Value = vnRow.Cells(ensColLsRcvPO.OID).Text
            HdfRcvPORefTypeOID.Value = vnRow.Cells(ensColLsRcvPO.RcvPORefTypeOID).Text
            HdfRcvRefTypeOID.Value = vnRow.Cells(ensColLsRcvPO.RcvRefTypeOID).Text
            HdfRcvPORefOID.Value = vnRow.Cells(ensColLsRcvPO.RcvPORefOID).Text
            psShowLsRcvPO(False)
        End If
    End Sub

    Protected Sub BtnLsRcvPOFind_Click(sender As Object, e As EventArgs) Handles BtnLsRcvPOFind.Click
        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If

        psFillGrvLsRcvPO()
    End Sub

    Protected Sub GrvLsRcvPO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvLsRcvPO.SelectedIndexChanged

    End Sub

    Private Sub GrvLsRcvPO_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsRcvPO.PageIndexChanging
        GrvLsRcvPO.PageIndex = e.NewPageIndex
        psFillGrvLsRcvPO()
    End Sub

    Protected Sub DstCompany_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DstCompany.SelectedIndexChanged
        TxtBrgCode.Text = ""
        TxtBrgName.Text = ""
        TxtBrgUnit.Text = ""
    End Sub
End Class