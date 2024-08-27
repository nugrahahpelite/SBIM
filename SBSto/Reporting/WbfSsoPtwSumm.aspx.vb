Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel

Public Class WbfSsoPtwSumm
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoPtwSumm"


    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsProcessDate As String
    Dim vsLogFolder As String
    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameOnly As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String

    Enum ensColList
        TransCode = 0
    End Enum
    Public Structure stuCrpPreviewType
        Const ByDataTable = "DataTable"
        Const ByQuery = "Query"
        Const ByQueryPopwin = "QueryPopwin"
        Const ByDataTablePopwin = "DataTablePopwin"
    End Structure
    Private Sub psDefaultDisplay()
        DivLsBrg.Style(HtmlTextWriterStyle.MarginTop) = "-175px"
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"



        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPrOption.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
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

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select ptw.TransCode,stn.TransName,ptw.vPtwCompanyCode,ptw.vPtwNo,"
        vnQuery += vbCrLf & "      convert(varchar(11),ptw.vPtwDate,106)vPtwDate,"
        vnQuery += vbCrLf & "      mwh.WarehouseName,mwh_d.WarehouseName vWarehouseName_Dest,"
        vnQuery += vbCrLf & "      ptw.PCKNo,ptw.PCLRefHNo,sts.TransStatusDescr,"
        vnQuery += vbCrLf & "      ptw.RcvPOHOID,ptw.RcvPONo,"
        vnQuery += vbCrLf & "      ptw.BRGCODE,msb.BRGNAME,ptw.vSumPtwScan1Qty,ptw.vPtwReceiveQty,ptw.vSumPtwScan2Qty,ptw.CreationDatetime"
        vnQuery += vbCrLf & "      From fnTbl_SsoPutaway_Summary('" & Session("UserID") & "') ptw"
        vnQuery += vbCrLf & "			inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=ptw.TransCode"
        vnQuery += vbCrLf & "			inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=ptw.TransCode and sts.TransStatus=ptw.TransStatus"
        vnQuery += vbCrLf & "           inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=ptw.BRGCODE and msb.CompanyCode=ptw.vPtwCompanyCode"
        vnQuery += vbCrLf & "			inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=ptw.WarehouseOID"
        vnQuery += vbCrLf & "			left outer join " & vnDBMaster & "Sys_Warehouse_MA mwh_d with(nolock) on mwh_d.OID=ptw.WarehouseOID_Dest"
        vnQuery += vbCrLf & "Where 1=1"
        vnQuery += vbCrLf & "            and msb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vnQuery += vbCrLf & "            and msb.BRGCODE like '%" & vnCrBrgCode & "%' and msb.BRGNAME like '%" & vnCrBrgName & "%'"

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "            and ptw.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If

        If IsDate(TxtListStart.Text) Then
            vnQuery += vbCrLf & "            and ptw.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnQuery += vbCrLf & "            and ptw.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        vnQuery += vbCrLf & " Order by ptw.CreationDatetime DESC"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        BtnPreview.Enabled = True
        BtnPreview.Visible = True

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



    Private Sub psShowPrOption(vriBo As Boolean)
        If vriBo Then
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPrOption.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Protected Sub BtnProOK_Click(sender As Object, e As EventArgs) Handles BtnProOK.Click
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
        psCrpXls()
    End Sub

    Private Sub psCrpXls()
        If LCase(RdbProXls.SelectedValue) = "pdf" Then
            Dim vnCrpFileName As String = ""

            psGenerateCrpPtwSumm(vnCrpFileName)

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
        End If
    End Sub

    Private Sub psGenerateCrpSummPutw(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSummPutw
        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))
        vbuCrpQuery = "Select ptw.TransCode,stn.TransName,ptw.vPtwCompanyCode,ptw.vPtwNo,"
        vbuCrpQuery += vbCrLf & "      convert(varchar(11),ptw.vPtwDate,106)vPtwDate,"
        vbuCrpQuery += vbCrLf & "      mwh.WarehouseName,mwh_d.WarehouseName vWarehouseName_Dest,"
        vbuCrpQuery += vbCrLf & "      ptw.PCKNo,ptw.PCLRefHNo,sts.TransStatusDescr,"
        vbuCrpQuery += vbCrLf & "      ptw.RcvPOHOID,ptw.RcvPONo,"
        vbuCrpQuery += vbCrLf & "      ptw.BRGCODE,msb.BRGNAME,ptw.vSumPtwScan1Qty,ptw.vPtwReceiveQty,ptw.vSumPtwScan2Qty,ptw.CreationDatetime"
        vbuCrpQuery += vbCrLf & "      From fnTbl_SsoPutaway_Summary('" & Session("UserID") & "') ptw"
        vbuCrpQuery += vbCrLf & "			inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=ptw.TransCode"
        vbuCrpQuery += vbCrLf & "			inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=ptw.TransCode and sts.TransStatus=ptw.TransStatus"
        vbuCrpQuery += vbCrLf & "           inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=ptw.BRGCODE and msb.CompanyCode=ptw.vPtwCompanyCode"
        vbuCrpQuery += vbCrLf & "			inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=ptw.WarehouseOID"
        vbuCrpQuery += vbCrLf & "			left outer join " & vnDBMaster & "Sys_Warehouse_MA mwh_d with(nolock) on mwh_d.OID=ptw.WarehouseOID_Dest"
        vbuCrpQuery += vbCrLf & "Where 1=1"
        vbuCrpQuery += vbCrLf & "            and msb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vbuCrpQuery += vbCrLf & "            and msb.BRGCODE like '%" & vnCrBrgCode & "%' and msb.BRGNAME like '%" & vnCrBrgName & "%'"

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vbuCrpQuery += vbCrLf & "            and ptw.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If
    End Sub


    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub


    'Protected Sub BtnStatus_Click(sender As Object, e As EventArgs) Handles BtnStatus.Click
    '    If Not IsNumeric(TxtTransID.Text) Then Exit Sub
    '    If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)

    '    Dim vnName1 As String = "Preview"
    '    Dim vnType As Type = Me.GetType()
    '    Dim vnClientScript As ClientScriptManager = Page.ClientScript
    '    If (Not vnClientScript.IsStartupScriptRegistered(vnType, vnName1)) Then
    '        Dim vnParam As String
    '        vnParam = "vqTrOID=" & TxtTransID.Text
    '        vnParam += "&vqTrCode=" & stuTransCode.SsoSSOH
    '        vnParam += "&vqTrNo=" & TxtSONo.Text

    '        vbuPreviewOnClose = "0"

    '        ifrPreview.Src = "WbfSsoTransStatus.aspx?" & vnParam
    '        psShowPreview(True)

    '        'Dim vnWinOpen As String
    '        'vnWinOpen = fbuOpenTransStatus(Session("RootFolder"), vnParam)
    '        'vnClientScript.RegisterStartupScript(vnType, vnName1, vnWinOpen, True)
    '        'vnClientScript = Nothing
    '    End If
    'End Sub
    Protected Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Print) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If
        psClearMessage()
        psCrpXls()
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuPreviewOnClose = "1"
        psShowPreview(False)
    End Sub

    Private Sub psGenerateCrpPtwSumm(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSummPutw

        Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
        Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))

        vbuCrpQuery = "Select '" & TxtListStart.Text & "'vPeriodeStart,'" & TxtListEnd.Text & "'vPeriodeEnd, ptw.TransCode,stn.TransName,ptw.vPtwCompanyCode,ptw.vPtwNo,"
        vbuCrpQuery += "   convert(varchar(11),ptw.vPtwDate,106)vPtwDate,"
        vbuCrpQuery += "   mwh.WarehouseName,mwh_d.WarehouseName vWarehouseName_Dest,"
        vbuCrpQuery += "   ptw.PCKNo,ptw.PCLRefHNo,sts.TransStatusDescr,"
        vbuCrpQuery += "   ptw.RcvPOHOID,ptw.RcvPONo,"
        vbuCrpQuery += "   ptw.BRGCODE,msb.BRGNAME,ptw.vSumPtwScan1Qty,ptw.vPtwReceiveQty,ptw.vSumPtwScan2Qty,ptw.CreationDatetime"
        vbuCrpQuery += "   From fnTbl_SsoPutaway_Summary('" & Session("UserID") & "') ptw"
        vbuCrpQuery += "        inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=ptw.TransCode"
        vbuCrpQuery += "        inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=ptw.TransCode and sts.TransStatus=ptw.TransStatus"
        vbuCrpQuery += "        inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=ptw.BRGCODE and msb.CompanyCode=ptw.vPtwCompanyCode"
        vbuCrpQuery += "        inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=ptw.WarehouseOID"
        vbuCrpQuery += "        left outer join " & vnDBMaster & "Sys_Warehouse_MA mwh_d with(nolock) on mwh_d.OID=ptw.WarehouseOID_Dest "
        vbuCrpQuery += "Where 1=1"
        vbuCrpQuery += "        and msb.CompanyCode='" & DstListCompany.SelectedValue & "'"
        vbuCrpQuery += "        and msb.BRGCODE like '%" & TxtListBrgCode.Text & "%' and msb.BRGNAME like '%" & TxtListBrgName.Text & "%'"

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vbuCrpQuery += "        and ptw.WarehouseOID=" & DstListWarehouse.SelectedValue
        End If

        If IsDate(TxtListStart.Text) Then
            vbuCrpQuery += "        and ptw.vPtwDate >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vbuCrpQuery += "        and ptw.vPtwDate <= '" & TxtListEnd.Text & "'"
        End If

        vbuCrpQuery += "Order by ptw.CreationDatetime DESC"
    End Sub

    Public Sub pbuCreateXlsx_PtwSumm(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTally", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")
            Dim vnCrBrgCode As String = fbuFormatString(Trim(TxtListBrgCode.Text))
            Dim vnCrBrgName As String = fbuFormatString(Trim(TxtListBrgName.Text))


            vnQuery = "Select ptw.TransCode,stn.TransName,ptw.vPtwCompanyCode,ptw.vPtwNo,"
            vnQuery += vbCrLf & "      convert(varchar(11),ptw.vPtwDate,106)vPtwDate,"
            vnQuery += vbCrLf & "      mwh.WarehouseName,mwh_d.WarehouseName vWarehouseName_Dest,"
            vnQuery += vbCrLf & "      ptw.PCKNo,ptw.PCLRefHNo,sts.TransStatusDescr,"
            vnQuery += vbCrLf & "      ptw.RcvPOHOID,ptw.RcvPONo,"
            vnQuery += vbCrLf & "      ptw.BRGCODE,msb.BRGNAME,ptw.vSumPtwScan1Qty,ptw.vPtwReceiveQty,ptw.vSumPtwScan2Qty,ptw.CreationDatetime"
            vnQuery += vbCrLf & "      From fnTbl_SsoPutaway_Summary('" & Session("UserID") & "') ptw"
            vnQuery += vbCrLf & "			inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=ptw.TransCode"
            vnQuery += vbCrLf & "			inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=ptw.TransCode and sts.TransStatus=ptw.TransStatus"
            vnQuery += vbCrLf & "           inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=ptw.BRGCODE and msb.CompanyCode=ptw.vPtwCompanyCode"
            vnQuery += vbCrLf & "			inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=ptw.WarehouseOID"
            vnQuery += vbCrLf & "			left outer join " & vnDBMaster & "Sys_Warehouse_MA mwh_d with(nolock) on mwh_d.OID=ptw.WarehouseOID_Dest"
            vnQuery += vbCrLf & "Where 1=1"
            vnQuery += vbCrLf & "            and msb.CompanyCode='" & DstListCompany.SelectedValue & "'"
            vnQuery += vbCrLf & "            and msb.BRGCODE like '%" & vnCrBrgCode & "%' and msb.BRGNAME like '%" & vnCrBrgName & "%'"

            If Val(DstListWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and ptw.WarehouseOID=" & DstListWarehouse.SelectedValue
            End If

            If IsDate(TxtListStart.Text) Then
                vnQuery += vbCrLf & "            and ptw.vPtwDate >= '" & TxtListStart.Text & "'"
            End If
            If IsDate(TxtListEnd.Text) Then
                vnQuery += vbCrLf & "            and ptw.vPtwDate <= '" & TxtListEnd.Text & "'"
            End If

            vnQuery += vbCrLf & " Order by ptw.CreationDatetime DESC"


            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = vriFileName & "-OID-" & vriTransID & "-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
            vriFileName = vnFNm & ".xlsx"

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Persiapan Membuat File Xlsx...")

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Creating Excel Workbook...")
            Dim vnWb As New XLWorkbook

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Creating Excel Worksheet...")
            vnWb.AddWorksheet("Sheet1")

            Dim vnIXLWorksheet As IXLWorksheet = vnWb.Worksheet(1)
            Dim vnXRow As Integer = 1
            Dim vnXCol As Integer = 0

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Header Report...")

            Dim vnDRow As DataRow
            Dim vnTransCode As String = ""
            Dim vnCompName As String = ""
            Dim vnWarehouseName As String = ""
            Dim vnWarehouseName_Dest As String = ""
            Dim vnPtwNo As String = ""
            Dim vnPtwDate As String = ""
            Dim vnStatus As String = ""
            Dim vnTransName As String = ""
            Dim vnCreationDateTime As String = ""

            If vnDtb.Rows.Count = 0 Then
                vnDRow = vnDtb.Rows(0)
            Else
                vnDRow = vnDtb.Rows(0)
                vnPtwNo = vnDRow.Item("vPtwNo")
                vnCompName = vnDRow.Item("vPtwCompanyCode")
                vnWarehouseName = vnDRow.Item("WarehouseName")
                vnWarehouseName_Dest = vnDRow.Item("vWarehouseName_Dest")
                vnTransCode = vnDRow.Item("TransCode")
                vnPtwDate = vnDRow.Item("vPtwDate")
                vnStatus = vnDRow.Item("TransStatusDescr")
                vnTransName = vnDRow.Item("TransName")
                vnCreationDateTime = fbuValStr(vnDRow.Item("CreationDatetime"))
            End If

            '<---------------ROW 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMBER BERKAT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 2
            vnXRow = vnXRow + 1
            vnXCol = 1

            If vriVarianOnly = 0 Then
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMMARY PUTAWAY"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMMARY PUTAWAY - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'NO SO
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PTW No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnPtwNo


            'Company
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'SOID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "TRANS CODE"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnTransCode


            '<---------------ROW 5
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "DATE"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnPtwDate


            'Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnWarehouseName

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Creation"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCreationDateTime


            'Sub Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse Dest"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnWarehouseName_Dest


            'CloseNote
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PCK No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("")

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 10
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "vPtwNo"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No.Invoice"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "RcvPOHOID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Diterima"

            For vnXCol = 1 To vnColCount
                vnIXLWorksheet.Row(vnRowIdxHead).Cell(vnXCol).Style.Fill.BackgroundColor = XLColor.LightGreen
            Next

            vnXRow = vnRowIdxHead
            vsTextStream.WriteLine("Proses : Mengisi Data...")
            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("TIDAK ADA DATA")
                vnXRow = vnXRow + 1
                vnXCol = 1
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol + 1).Value = "TIDAK ADA DATA"
            Else
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCKNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumPtwScan1Qty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumPtwScan2Qty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwReceiveQty")
                Next

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    vnXCol = 6
                    vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                    vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                Next
            End If

            vsTextStream.WriteLine("Files Names " & vriFileName)

            vsTextStream.WriteLine("-------------------------------EOF-------------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            Dim Response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response

            Using MyMemoryStream As New MemoryStream()
                vnWb.SaveAs(MyMemoryStream)

                Response.Buffer = True
                Response.Clear()
                Response.ClearHeaders()
                Response.ClearContent()

                'Response.ContentType = "application/vnd ms excel xlsx"
                Response.ContentType = "application/vnd.xls"

                Response.AddHeader("content-disposition", "attachment; filename=" & vriFileName & ";")
                Response.Charset = ""

                MyMemoryStream.WriteTo(Response.OutputStream)
                MyMemoryStream.Close()
                Response.OutputStream.Close()
                Response.Flush()

                '09 Jan 2023
                'Response.End()

                Response.SuppressContent = True
                HttpContext.Current.ApplicationInstance.CompleteRequest()
            End Using

            '<---09 Jan 2023
            'Replace following
            'HttpContext.Current.Response.End();

            'with this :
            'HttpContext.Current.Response.Flush(); // Sends all currently buffered output To the client.
            'HttpContext.Current.Response.SuppressContent = True;  // Gets Or sets a value indicating whether To send HTTP content To the client.
            'HttpContext.Current.ApplicationInstance.CompleteRequest(); // Causes A
            '<==09 Jan 2023

        Catch ex As Exception
            pbMsgError = ex.Message
            If Not IsNothing(vsTextStream) Then
                vsTextStream.WriteLine("TERJADI ERROR : LAPORKAN KE IT")
                vsTextStream.WriteLine("ERROR DESCRIPTION : ")
                vsTextStream.WriteLine(ex.Message)

                vsTextStream.WriteLine("-------------------------------EOF-------------------------------")

                vsTextStream.Close()
                vsTextStream = Nothing
                vsFso = Nothing
            End If
            FileCopy(vsLogFileName, vsLogFileNameError)
        End Try
    End Sub

    Protected Sub BtnXLS_Click(sender As Object, e As EventArgs) Handles BtnXLS.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnFileName As String = ""
        pbuCreateXlsx_SummaryPutaway(vnFileName, Session("UserOID"), DstListWarehouse, DstListCompany, TxtListBrgCode, TxtListStart, TxtListEnd, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class