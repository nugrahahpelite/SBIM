Imports System.Data.SqlClient
Public Class WbfSsoStockKarantina
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoStockKarantina"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Enum ensColList
        vStorageOID = 0
        vStorageStockOID = 1
        vStorageInfoHtml = 2
        TransCode = 3
        TransName = 4
        TransOID = 5
        CompanyCode = 6
        RcvPOHOID = 7
        RcvPONo = 8
        vRcvPODate = 9
        BRGCODE = 10
        BRGNAME = 11
        NoteKarantina = 12
        QtyKarantina = 13
        QtyKrRelease = 14
        QtyKrReceive = 15
        vQtyKrOutstanding = 16
        TransStatus = 17
        TransStatusDescr = 18
        vApprove = 19
        vCreationDatetime = 20
        vPutawayDone = 21
        vApproved = 22
    End Enum
    Private Sub psDefaultDisplay()
        DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivConfirm.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
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
        Session("CurrentFolder") = "Approval"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            psDefaultDisplay()

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoStockKarantina, vnSQLConn)

            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstListCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, False, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psClearMessage()
        LblMsgError.Text = ""
        LblMsgListCompany.Text = ""
        LblMsgListBrg.Text = ""
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        psClearMessage()

        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            Exit Sub
        End If
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
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

        If ChkSt_Baru.Checked = False And ChkSt_OnPutaway.Checked = False And ChkSt_PutawayDone.Checked = False And ChkSt_Approved.Checked = False Then
            ChkSt_Baru.Checked = True
            ChkSt_PutawayDone.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If ChkSt_Baru.Checked = True Then
            vnCrStatus += enuTCSTKR.Baru & ","
        End If
        If ChkSt_OnPutaway.Checked = True Then
            vnCrStatus += enuTCSTKR.On_Putaway & ","
        End If
        If ChkSt_PutawayDone.Checked = True Then
            vnCrStatus += enuTCSTKR.Putaway_Done & ","
        End If
        If ChkSt_Approved.Checked = True Then
            vnCrStatus += enuTCSTKR.Approved & ","
        End If
        If vnCrStatus <> "" Then
            vnCrStatus = vbCrLf & "      and skr.TransStatus in(" & Mid(vnCrStatus, 1, Len(vnCrStatus) - 1) & ")"
        End If

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select mst.vStorageOID,skr.OID vStorageStockOID,mst.vStorageInfoHtml,stn.TransCode,stn.TransName,skr.TransOID,skr.CompanyCode,skr.RcvPOHOID,rch.RcvPONo,convert(varchar(11),rch.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "      skr.BRGCODE,msb.BRGNAME,"
        vnQuery += vbCrLf & "	   skr.NoteKarantina,skr.QtyKarantina,skr.QtyKrRelease,skr.QtyKrReceive,"
        vnQuery += vbCrLf & "      (abs(skr.QtyKarantina) - abs(skr.QtyKrRelease) - abs(skr.QtyKrReceive))vQtyKrOutstanding,"
        vnQuery += vbCrLf & "      skr.TransStatus,sts.TransStatusDescr,"
        vnQuery += vbCrLf & "      case when skr.TransStatus = 0 and skr.QtyKarantina < 0 then 'Approve'"
        vnQuery += vbCrLf & "           when skr.TransStatus = " & enuTCSTKR.Putaway_Done & " and skr.QtyKarantina > 0 then 'Approve'"
        vnQuery += vbCrLf & "           else '' end vApprove,"
        vnQuery += vbCrLf & "      convert(varchar(11),skr.CreationDatetime,106)+' '+convert(varchar(5),skr.CreationDatetime,108) vCreationDatetime,"
        vnQuery += vbCrLf & "      convert(varchar(11),skr.PutawayDoneDatetime,106)+' '+convert(varchar(5),skr.PutawayDoneDatetime,108)+'<br />'+usp.UserName vPutawayDone,"
        vnQuery += vbCrLf & "      convert(varchar(11),skr.ApprovedDatetime,106)+' '+convert(varchar(5),skr.ApprovedDatetime,108)+'<br />'+usa.UserName vApproved"
        vnQuery += vbCrLf & "      From Sys_SsoStockKarantina_TR skr with(nolock)"
        vnQuery += vbCrLf & "	        inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=skr.RcvPOHOID"
        vnQuery += vbCrLf & "			inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=skr.TransCode_Source"
        vnQuery += vbCrLf & "			inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=skr.TransCode and sts.TransStatus=skr.TransStatus"
        vnQuery += vbCrLf & "			left outer join Sys_SsoUser_MA usp with(nolock) on usp.OID=skr.PutawayDoneUserOID"
        vnQuery += vbCrLf & "			left outer join Sys_SsoUser_MA usa with(nolock) on usa.OID=skr.ApprovedUserOID"
        vnQuery += vbCrLf & "           inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=skr.BRGCODE and msb.CompanyCode=skr.CompanyCode"
        vnQuery += vbCrLf & "			inner join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) mst on mst.vStorageOID=skr.StorageOID"
        vnQuery += vbCrLf & "Where 1=1"

        vnQuery += vbCrLf & vnCrStatus
        vnQuery += vbCrLf & "            and msb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "            and msb.BRGCODE like '%" & vnCrBrgCode & "%' and msb.BRGNAME like '%" & vnCrBrgName & "%'"

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and mst.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If
        If ChkOSOnly.Checked Then
            vnQuery += vbCrLf & "            and (abs(skr.QtyKarantina) - abs(skr.QtyKrRelease) - abs(skr.QtyKrReceive)) > 0"
        End If
        vnQuery += vbCrLf & " Order by skr.RcvPOHOID,msb.BRGCODE"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub
    Protected Sub BtnListBrgCode_Click(sender As Object, e As EventArgs) Handles BtnListBrgCode.Click
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psShowLsBrg(True)
    End Sub
    Private Sub psShowLsBrg(vriBo As Boolean)
        If vriBo Then
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnLsBrgClose_Click(sender As Object, e As EventArgs) Handles BtnLsBrgClose.Click
        psShowLsBrg(False)
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
        vnQuery += vbCrLf & "Where CompanyCode='" & DstListCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtLsBrg.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtLsBrg.Text) & "%')"
        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvLsBrg.DataSource = vnDtb
        GrvLsBrg.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub GrvLsBrg_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvLsBrg.PageIndexChanging
        GrvLsBrg.PageIndex = e.NewPageIndex
        psFillGrvLsBrg()
    End Sub

    Private Sub GrvLsBrg_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvLsBrg.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvLsBrg.Rows(vnIdx)
            vnValue = DirectCast(vnRow.Cells(0).Controls(0), LinkButton).Text
            TxtListBrgCode.Text = vnValue
            TxtListBrgName.Text = vnRow.Cells(1).Text
            psShowLsBrg(False)
        End If
    End Sub

    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
        If DstListCompany.SelectedValue = "" Then
            LblMsgListCompany.Text = "Pilih Company"
            Exit Sub
        End If

        psFillGrvLsBrg()
    End Sub
    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Private Sub psApprove(vriStkrHOID As String)
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "psApprove", vriStkrHOID, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
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

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoStockKarantina_TR set TransStatus=" & enuTCSTKR.Approved & ",ApprovedUserOID=" & Session("UserOID") & ",ApprovedDatetime=getdate() Where OID=" & vriStkrHOID
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("1")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("2")
            vsTextStream.WriteLine("pbuInsertStatusSTKR...Start")
            pbuInsertStatusSTKR(vriStkrHOID, enuTCSTKR.Approved, Session("UserOID"), vnSQLConn, vnSQLTrans)
            vsTextStream.WriteLine("pbuInsertStatusSTKR...End")

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            vsTextStream.WriteLine("Prepare Sukses")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            Dim vnGRow As GridViewRow
            vnGRow = GrvList.Rows(HdfDetailRowIdx.Value)
            vnGRow.Cells(ensColList.TransStatus).Text = enuTCSTKR.Approved
            vnGRow.Cells(ensColList.TransStatusDescr).Text = "Approved"
            DirectCast(vnGRow.Cells(ensColList.vApprove).Controls(0), LinkButton).Text = ""
            vnGRow.Cells(ensColList.vApproved).Text = "Just Approved"

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

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        If e.CommandName = "vApprove" Then
            If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Approve) = False Then
                LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
                Exit Sub
            End If

            Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnGRow As GridViewRow = GrvList.Rows(vnRowIdx)
            HdfDetailRowIdx.Value = vnRowIdx

            psShowConfirm(True)
        End If
    End Sub

    Private Sub BtnConfirmNo_Click(sender As Object, e As EventArgs) Handles BtnConfirmNo.Click
        psShowConfirm(False)
    End Sub

    Private Sub BtnConfirmYes_Click(sender As Object, e As EventArgs) Handles BtnConfirmYes.Click
        psApprove(GrvList.Rows(HdfDetailRowIdx.Value).Cells(ensColList.vStorageStockOID).Text)
        psShowConfirm(False)
    End Sub
End Class