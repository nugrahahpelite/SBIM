Imports System.Data.SqlClient
Public Class WbfSsoPaketMs
    Inherits System.Web.UI.Page
    Enum ensColList
        CompanyCode = 0
        PAKETCODE = 1
        PAKETNAME = 2
        OID = 3
    End Enum
    Enum ensColListBrg
        BRGCODE = 0
        BRGNAME = 1
        BRGUNIT = 2
    End Enum
    Enum ensColDetail
        vAddItem = 0
        BRGCODE_ORIG = 1
        BRGCODE = 2
        BRGNAME = 3
        PaketQty = 4
        TxtPaketQty = 5
        vDelItem = 6
        vMessage = 7
    End Enum

    Private Sub psClearData()
        TxtOID.Text = ""
        TxtPaketCode.Text = ""
        TxtPaketName.Text = ""
    End Sub

    Private Sub psClearMessage()
        LblMsgPaketCode.Visible = False
        LblMsgPaketName.Visible = False
        LblMsgErrorNE.Visible = False
    End Sub
    Private Sub psDefaultDisplay()
        DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanListBrg.Style(HtmlTextWriterStyle.Position) = "absolute"
        PanListBrg.Style(HtmlTextWriterStyle.MarginTop) = "-450px"
    End Sub
    Private Sub psDisplayData(vriOID As String)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select * From " & fbuGetDBMaster() & "Sys_MstPaketH_MA Where OID=" & vriOID
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            DstCompany.SelectedValue = fbuValStr(vnDtb.Rows(0).Item("CompanyCode"))
            TxtPaketCode.Text = fbuValStr(vnDtb.Rows(0).Item("PAKETCODE"))
            TxtPaketName.Text = fbuValStr(vnDtb.Rows(0).Item("PAKETNAME"))
        End If
        vnDtb.Dispose()

        psFillGrvDetail(vriOID, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        BtnEdit.Enabled = (Session("UserAdmin") = 1)
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPaketCode.ReadOnly = Not vriBo
        TxtPaketName.ReadOnly = Not vriBo
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsBarang, vnSQLConn)

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
    Private Sub psShowListBrg(vriBo As Boolean)
        If vriBo Then
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
            TxtListBrg.Focus()
        Else
            DivListBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtPaketCode.Text)) = 0 Then
            LblMsgPaketCode.Text = "Isi Kode Paket"
            LblMsgPaketCode.Visible = True
            vnSave = False
        End If
        If Len(Trim(TxtPaketName.Text)) = 0 Then
            LblMsgPaketName.Text = "Isi Nama Paket"
            LblMsgPaketName.Visible = True
            vnSave = False
        End If
        Try
            Dim vnSQLConn As New SqlConnection
            Dim vnSQLTrans As SqlTransaction
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnQuery As String

            Dim vnHOID As String

            Dim vnDBMaster As String = fbuGetDBMaster()

            Dim vnUserOID As String = Session("UserOID")

            Dim vnPAKETCODE As String = fbuFormatString(Trim(TxtPaketCode.Text))
            Dim vnPAKETNAME As String = fbuFormatString(Trim(TxtPaketName.Text))
            If HdfActionStatus.Value = cbuActionNew Then
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_MstPaketH_MA Where CompanyCode='" & DstCompany.SelectedValue & "' and (PAKETCODE='" & vnPAKETCODE & "' or PAKETNAME='" & vnPAKETNAME & "')"
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgPaketName.Text = "Paket " & Trim(TxtPaketCode.Text) & " atau " & Trim(TxtPaketName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Paket"
                    LblMsgPaketName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnQuery = "Select max(OID) From " & vnDBMaster & "Sys_MstPaketH_MA with(nolock)"
                vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Insert into " & vnDBMaster & "Sys_MstPaketH_MA(OID,"
                vnQuery += vbCrLf & "PAKETCODE,"
                vnQuery += vbCrLf & "PAKETNAME,"
                vnQuery += vbCrLf & "CompanyCode,"
                vnQuery += vbCrLf & "CreationDatetime,CreationUserOID)"
                vnQuery += "values(" & vnHOID & ","
                vnQuery += vbCrLf & "'" & vnPAKETCODE & "',"
                vnQuery += vbCrLf & "'" & vnPAKETNAME & "',"
                vnQuery += vbCrLf & "'" & DstCompany.SelectedValue & "',"
                vnQuery += vbCrLf & "getdate()," & vnUserOID & ");"
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            Else
                vnHOID = TxtOID.Text
                vnQuery = "Select count(*) from " & vnDBMaster & "Sys_MstPaketH_MA Where CompanyCode='" & DstCompany.SelectedValue & "' and PAKETNAME='" & vnPAKETNAME & "' and OID<>" & vnHOID
                If fbuGetDataNumSQL(vnQuery, vnSQLConn) > 0 Then
                    LblMsgPaketName.Text = "Paket Name " & Trim(TxtPaketName.Text) & " Sudah digunakan" & vbCrLf & "Silakan Cek Daftar Paket"
                    LblMsgPaketName.Visible = True
                    vnSave = False
                End If
                If Not vnSave Then
                    vnSQLConn.Close()
                    vnSQLConn.Dispose()
                    vnSQLConn = Nothing
                    Exit Sub
                End If

                vnSQLTrans = vnSQLConn.BeginTransaction()

                vnQuery = "Update " & vnDBMaster & "Sys_MstPaketH_MA set "
                vnQuery += vbCrLf & "PAKETNAME='" & vnPAKETNAME & "',"
                vnQuery += vbCrLf & "ModificationDatetime=getdate(),ModificationUserOID=" & vnUserOID
                vnQuery += vbCrLf & " Where OID=" & vnHOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)
            End If

            If fsSaveDetail(vnHOID, vnSQLConn, vnSQLTrans) Then
                Dim vnHistoryOID As String
                vnQuery = "Select max(OID) From Sys_MstPaketH_HS with(nolock)"
                vnHistoryOID = fbuGetDataNumSQLTrans(vnQuery, vnSQLConn, vnSQLTrans) + 1

                vnQuery = "Insert into Sys_MstPaketH_HS Select " & vnHistoryOID & ",OID,CompanyCode,PAKETCODE,PAKETNAME,getdate()," & vnUserOID & " From " & vnDBMaster & "Sys_MstPaketH_MA with(nolock) Where OID=" & vnHOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnQuery = "Insert into Sys_MstPaketD_HS Select " & vnHistoryOID & ",PaketHOID,BRGCODE,PaketQty From " & vnDBMaster & "Sys_MstPaketD_MA with(nolock) Where PaketHOID=" & vnHOID
                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                HdfActionStatus.Value = cbuActionNorm
                psFillGrvDetail(vnHOID, vnSQLConn)

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing

                psEnableInput(False)
                psEnableSave(False)
                BtnEdit.Enabled = True

            Else
                vnSQLTrans.Rollback()
                vnSQLTrans = Nothing

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
            End If


        Catch ex As Exception
            LblMsgErrorNE.Text = ex.Message
            LblMsgErrorNE.Visible = True
        End Try
    End Sub

    Private Function fsSaveDetail(vriHOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        Dim vnReturn As Boolean = True
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnQuery As String

        Dim vn As Integer
        Dim vnGRow As GridViewRow

        Dim vnTxtPaketQty As TextBox
        Dim vnBrgCode As String
        Dim vnBrgCode_Orig As String
        Dim vnBrgCodeList As String = ""
        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtPaketQty = vnGRow.FindControl("TxtPaketQty")

            If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE_ORIG).Text) = "" Then
                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) <> "" Then
                    vnBrgCode = UCase(vnGRow.Cells(ensColDetail.BRGCODE).Text)

                    If InStr(vnBrgCodeList, vnBrgCode) = 0 Then
                        vnQuery = "Insert into " & vnDBMaster & "Sys_MstPaketD_MA"
                        vnQuery += vbCrLf & "(PaketHOID,"
                        vnQuery += vbCrLf & "BRGCODE,"
                        vnQuery += vbCrLf & "PaketQty"
                        vnQuery += vbCrLf & ")"
                        vnQuery += vbCrLf & "values('" & vriHOID & "',"
                        vnQuery += vbCrLf & "'" & vnBrgCode & "',"
                        vnQuery += vbCrLf & Val(Replace(vnTxtPaketQty.Text, "", "")) & ")"

                        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                    Else
                        vnGRow.Cells(ensColDetail.vMessage).Text = vnBrgCode & " Sudah di Paket"
                        vnReturn = False
                    End If

                    vnBrgCodeList += "," & vnBrgCode
                End If
            Else
                vnBrgCode_Orig = UCase(vnGRow.Cells(ensColDetail.BRGCODE_ORIG).Text)

                If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE).Text) = "" Then
                    vnQuery = "Delete " & vnDBMaster & "Sys_MstPaketD_MA Where PaketHOID=" & vriHOID & " and BRGCODE='" & vnBrgCode_Orig & "'"

                Else
                    vnQuery = "Update " & vnDBMaster & "Sys_MstPaketD_MA SET"
                    vnQuery += vbCrLf & "PaketQty=" & Val(Replace(vnTxtPaketQty.Text, "", "")) & ""
                    vnQuery += vbCrLf & "Where PaketHOID=" & vriHOID & " and BRGCODE='" & vnBrgCode_Orig & "'"
                End If

                pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)

                vnBrgCodeList += "," & vnBrgCode_Orig
            End If
        Next

        Return vnReturn
    End Function

    Protected Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psClearData()
        psEnableInput(True)
        psEnableSave(True)
        HdfActionStatus.Value = cbuActionNew

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvDetail("", vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
        psEnableSave(False)
        HdfActionStatus.Value = cbuActionNorm
        If Val(TxtOID.Text) = 0 Then
            psClearData()
        Else
            psDisplayData(TxtOID.Text)
        End If
    End Sub

    Protected Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If Trim(TxtPaketCode.Text) = "" Then Exit Sub
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        psEnableInput(True)
        psEnableSave(True)
        TxtPaketCode.ReadOnly = True
        HdfActionStatus.Value = cbuActionEdit

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        psFillGrvDetail(TxtOID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
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
            Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)
            Dim vnOID As String = vnGRow.Cells(ensColList.OID).Text
            TxtOID.Text = vnOID

            psDisplayData(vnOID)

            psEnableInput(False)
            psEnableSave(False)
            HdfActionStatus.Value = cbuActionNorm
        End If
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
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnCriteria As String = fbuFormatString(Trim(TxtKriteria.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.CompanyCode,PM.PAKETCODE,PM.PAKETNAME,PM.OID,"
        vnQuery += vbCrLf & "PM.CreationDatetime,SM.UserName CreationUserName,"
        vnQuery += vbCrLf & "PM.ModificationDatetime,AM.UserName ModificationUserName"
        vnQuery += vbCrLf & " From " & vnDBMaster & "Sys_MstPaketH_MA PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.CreationUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA AM on AM.OID=PM.ModificationUserOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.PLCompanyCode and uc.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "      Where (PM.PAKETCODE like '%" & vnCriteria & "%' OR PM.PAKETNAME like '%" & vnCriteria & "%')"

        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "            and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If

        vnQuery += vbCrLf & " Order by PM.CompanyCode,PM.PAKETCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
    Private Sub psFillGrvDetail(vriPaketHOID As String, vriSQLConn As SqlConnection)
        psClearMessage()
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnvAddItem As String = "..."
        Dim vnBRGCODE As String = ""
        Dim vnBRGNAME As String = ""
        Dim vnPaketQty As String = "0"
        Dim vnvDelItem As String = "Hapus"
        Dim vnvMessage As String = ""

        vnQuery = "Select ''vAddItem,pkd.BRGCODE,mbr.BRGNAME,pkd.PaketQty,'Hapus'vDelItem,''vMessage"
        vnQuery += vbCrLf & "  From " & vnDBMaster & "Sys_MstPaketD_MA pkd with(nolock)"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstPaketH_MA pkh with(nolock) on pkh.OID=pkd.PaketHOID"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mbr with(nolock) on mbr.BRGCODE=pkd.BRGCODE and mbr.CompanyCode=pkh.CompanyCode"
        vnQuery += vbCrLf & " Where pkd.PaketHOID='" & vriPaketHOID & "'"
        vnQuery += vbCrLf & "Order by pkd.BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        Dim vn As Integer
        If HdfActionStatus.Value = cbuActionNorm Then
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.PaketQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.PaketQty).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.TxtPaketQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.TxtPaketQty).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.vMessage).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.vMessage).ItemStyle.CssClass = "myDisplayNone"
        Else
            GrvDetail.Columns(ensColDetail.vAddItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vAddItem).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vDelItem).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vDelItem).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.PaketQty).HeaderStyle.CssClass = "myDisplayNone"
            GrvDetail.Columns(ensColDetail.PaketQty).ItemStyle.CssClass = "myDisplayNone"

            GrvDetail.Columns(ensColDetail.TxtPaketQty).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.TxtPaketQty).ItemStyle.CssClass = ""

            GrvDetail.Columns(ensColDetail.vMessage).HeaderStyle.CssClass = ""
            GrvDetail.Columns(ensColDetail.vMessage).ItemStyle.CssClass = ""

            For vn = 0 To 10
                vnDtb.Rows.Add(New Object() {vnvAddItem, vnBRGCODE, vnBRGNAME, vnPaketQty, vnvDelItem, vnvMessage})
            Next
        End If

        GrvDetail.DataSource = vnDtb
        GrvDetail.DataBind()

        Dim vnGRow As GridViewRow
        Dim vnTxtPaketQty As TextBox

        For vn = 0 To GrvDetail.Rows.Count - 1
            vnGRow = GrvDetail.Rows(vn)
            vnTxtPaketQty = vnGRow.FindControl("TxtPaketQty")

            vnTxtPaketQty.Text = fbuValStrHtml(vnGRow.Cells(ensColDetail.PaketQty).Text)
        Next
    End Sub

    Protected Sub BtnListBrgClose_Click(sender As Object, e As EventArgs) Handles BtnListBrgClose.Click
        psShowListBrg(False)
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
        If e.CommandName = "BRGCODE" Then
            Dim vnKodeBarang As String = DirectCast(vnGRowList.Cells(ensColListBrg.BRGCODE).Controls(0), LinkButton).Text

            Dim vnGRowDetail As GridViewRow = GrvDetail.Rows(HdfDetailRowIdx.Value)

            vnGRowDetail.Cells(ensColDetail.BRGCODE).Text = vnKodeBarang
            vnGRowDetail.Cells(ensColDetail.BRGNAME).Text = vnGRowList.Cells(ensColListBrg.BRGNAME).Text

            psShowListBrg(False)
        End If
    End Sub

    Protected Sub BtnListBrgFind_Click(sender As Object, e As EventArgs) Handles BtnListBrgFind.Click
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

    Protected Sub GrvDetail_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvDetail.SelectedIndexChanged

    End Sub

    Private Sub GrvDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvDetail.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvDetail.Rows(vnIdx)

        If e.CommandName = "vAddItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True

                Exit Sub
            End If
            If fbuValStrHtml(vnGRow.Cells(ensColDetail.BRGCODE_ORIG).Text) = "" Then
                HdfDetailRowIdx.Value = vnIdx
                psShowListBrg(True)
            End If

        ElseIf e.CommandName = "vDelItem" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                LblMsgError.Visible = True
                Exit Sub
            End If
            vnGRow.Cells(ensColDetail.BRGCODE).Text = ""
            vnGRow.Cells(ensColDetail.BRGNAME).Text = ""

            Dim vnTxtPaketQty As TextBox
            vnTxtPaketQty = vnGRow.FindControl("TxtPaketQty")
            vnTxtPaketQty.Text = "0"

            vnGRow.Visible = False
        End If
    End Sub

    Protected Sub GrvListBrg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvListBrg.SelectedIndexChanged

    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub
End Class