Imports System.Data.SqlClient
Public Class WbfSsoSOStatus
    Inherits System.Web.UI.Page
    Private Sub psDefaultDisplay()
        DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPreview.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Reporting"

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoStorageStock, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub
    Private Sub psShowPreview(vriBo As Boolean)
        If vriBo Then
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "visible"
        Else
            DivPreview.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psClearMessage()
        LblMsgError.Text = ""
    End Sub
    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        psClearMessage()

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

        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnQuery As String

        Dim vnDtb As New DataTable
        vnQuery = "Select so.SOHOID,so.SOCompanyCode,wh.WarehouseName,so.SONo,so.SONote,"
        vnQuery += vbCrLf & "       so.vTotalItem_System,so.vTotalQty_System,so.vTotalItem_Scanned,so.vPercentageItem_Scanned,"
        vnQuery += vbCrLf & "       so.vTotalQty_Scanned,so.vPercentageQty_Scanned,so.vTotalItem_Selisih,so.vPercentageItem_Selisih"
        vnQuery += vbCrLf & "       From fnTbl_SsoSOStatus('" & Session("UserID") & "')so"
        vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=so.SOWarehouseOID"

        If vnUserCompanyCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=so.SOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=so.SOWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "       Where 1=1"
        If ChkSt_Closed.Checked = True And ChkSt_NotClosed.Checked = False Then
            vnQuery += vbCrLf & "             and so.TransStatus=" & enuTCSSOH.Closed
        ElseIf ChkSt_Closed.Checked = False And ChkSt_NotClosed.Checked = True Then
            vnQuery += vbCrLf & "             and so.TransStatus<>" & enuTCSSOH.Closed
        End If

        vnQuery += vbCrLf & " order by so.SOHOID"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        LblMsgReturn.Text = Format(Date.Now, "dd MMM yyyy HH:mm:ss")
    End Sub

    Protected Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        psShowPreview(False)
    End Sub

    Private Sub psGenerateCrpSOStatus(ByRef vriCrpFileName As String)
        Dim vnDBMaster As String = fbuGetDBMaster()
        vriCrpFileName = stuSsoCrp.CrpSsoSOStatus

        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnUserWarehouseCode As String = Session("UserWarehouseCode")

        vbuCrpQuery = "Select so.*,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName"
        vbuCrpQuery += vbCrLf & "       From fnTbl_SsoSOStatus('" & Session("UserID") & "')so"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=so.SOSubWarehouseOID"
        vbuCrpQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=so.SOWarehouseOID"

        If vnUserCompanyCode = "" Then
        Else
            vbuCrpQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=so.SOCompanyCode and uc.UserOID=" & Session("UserOID")
        End If
        If vnUserWarehouseCode = "" Then
        Else
            vbuCrpQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=so.SOWarehouseOID and uw.UserOID=" & Session("UserOID")
        End If

        vbuCrpQuery += vbCrLf & "       Where 1=1"

        If ChkSt_Closed.Checked = True And ChkSt_NotClosed.Checked = False Then
            vbuCrpQuery += vbCrLf & "             and so.TransStatus=" & enuTCSSOH.Closed
        ElseIf ChkSt_Closed.Checked = False And ChkSt_NotClosed.Checked = True Then
            vbuCrpQuery += vbCrLf & "             and so.TransStatus!=" & enuTCSSOH.Closed
        End If

        vbuCrpQuery += vbCrLf & " order by so.SOHOID"
    End Sub

    Protected Sub BtnPdf_Click(sender As Object, e As EventArgs) Handles BtnPdf.Click
        Dim vnCrpFileName As String = ""
        psGenerateCrpSOStatus(vnCrpFileName)

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

    Protected Sub BtnXLS_Click(sender As Object, e As EventArgs) Handles BtnXLS.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnFileName As String = ""
        pbuCreateXlsx_SOStatus(vnFileName, Session("UserOID"), Session("UserCompanyCode"), Session("UserWarehouseCode"), ChkSt_Closed, ChkSt_NotClosed, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub
End Class