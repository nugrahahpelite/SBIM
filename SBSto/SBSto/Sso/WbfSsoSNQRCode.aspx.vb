Imports System.Data.SqlClient
Imports Spire.Barcode
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.IO
Public Class WbfSsoSNQRCode
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoSNQRCode"

    Dim settings As BarcodeSettings

    Dim vsIOFileStream As System.IO.FileStream
    Dim vsFileLength As Long

    Dim vsQRDir As String

    Const csFileFormat = ".jpg"
    Private Sub psClearData()
        TxtBrgCode.Text = ""
        TxtBrgName.Text = ""
        TxtBrgUnit.Text = ""
        TxtPrintSN.Text = ""
        TxtPrintNote.Text = ""
        TxtOID.Text = ""
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Visible = False
        LblMsgBrgName.Visible = False
        LblMsgPrintSN.Visible = False
        LblMsgPrintNote.Visible = False
        LblMsgError.Visible = False
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtPrintSN.ReadOnly = Not vriBo
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

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, True, vnSQLConn)
                pbuFillDstCompany(DstListCompany, True, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, True, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, True, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psDefaultDisplay()
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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.OID,PM.CompanyCode,PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,PM.PrintSN,PM.PrintNote,PM.PrintDatetime,"
        vnQuery += vbCrLf & "      GM.UserName vPrintUserName,"
        vnQuery += vbCrLf & "      case when abs(PrintReq)=1 then 'Y' else 'N' end vPrintReq,PM.PrintReqOID,"
        vnQuery += vbCrLf & "      RM.ReqSNDatetime,"
        vnQuery += vbCrLf & "      RU.UserName vReqSNUserName"
        vnQuery += vbCrLf & " From Sys_SsoPrintSNQRBarang_TR PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA GM on GM.OID=PM.PrintUserOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoPrintSNRequest_TR RM on RM.OID=PM.PrintReqOID"
        vnQuery += vbCrLf & "      left outer join Sys_SsoUser_MA RU on RU.OID=RM.ReqSNUserOID"

        vnQuery += vbCrLf & "Where 1=1"
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If

        If Trim(TxtListBarang.Text) <> "" Then
            vnQuery += vbCrLf & "     and (PM.BRGCODE like '%" & Trim(TxtListBarang.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtListBarang.Text) & "%')"
        End If
        If Trim(TxtListSN.Text) <> "" Then
            vnQuery += vbCrLf & "     and PM.PrintSN like '%" & Trim(TxtListSN.Text) & "%'"
        End If

        vnQuery += vbCrLf & "Order by PM.BRGNAME,PM.PrintSN"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
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
        vnQuery += vbCrLf & "Where abs(IsSN)=1 and CompanyCode='" & DstCompany.SelectedValue & "' and (PM.BRGCODE like '%" & Trim(TxtLsBrg.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtLsBrg.Text) & "%')"
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
            TxtBrgCode.Text = vnValue
            TxtBrgName.Text = vnRow.Cells(1).Text
            TxtBrgUnit.Text = vnRow.Cells(2).Text
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
        If Len(Trim(TxtBrgCode.Text)) = 0 Then
            LblMsgBrgName.Text = "Pilih Barang"
            LblMsgBrgName.Visible = True
            vnSave = False
        End If
        If Trim(TxtPrintSN.Text) = "" Then
            LblMsgPrintSN.Text = "Isi SN"
            LblMsgPrintSN.Visible = True
            vnSave = False
        End If
        If Trim(TxtPrintNote.Text) = "" Then
            LblMsgPrintNote.Text = "Isi Note"
            LblMsgPrintNote.Visible = True
            vnSave = False
        End If
        If Not vnSave Then
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
            Dim vnUserOID As String = Session("UserOID")
            Dim vnCompanyCode As String = DstCompany.SelectedValue

            vnQuery = "Select isnull(max(OID),0)+1 from Sys_SsoPrintSNQRBarang_TR"
            vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction()
            vnBeginTrans = True
            vnQuery = "Insert into Sys_SsoPrintSNQRBarang_TR("
            vnQuery += vbCrLf & "OID,CompanyCode,"
            vnQuery += vbCrLf & "BRGCODE,BRGNAME,BRGUNIT,"
            vnQuery += vbCrLf & "PrintSN,PrintNote,"
            vnQuery += vbCrLf & "PrintDatetime,PrintUserOID)"
            vnQuery += vbCrLf & "values("
            vnQuery += vbCrLf & vnHOID & ",'" & vnCompanyCode & "',"
            vnQuery += vbCrLf & "'" & TxtBrgCode.Text & "','" & TxtBrgName.Text & "','" & TxtBrgUnit.Text & "',"
            vnQuery += vbCrLf & "'" & TxtPrintSN.Text & "','" & fbuFormatString(Trim(TxtPrintNote.Text)) & "',"
            vnQuery += vbCrLf & "getdate()," & vnUserOID & ")"
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            If fsGenBrgSNQRCode(vnCompanyCode, TxtBrgCode.Text, TxtPrintSN.Text, vnSQLConn, vnSQLTrans) = True Then
                vnBeginTrans = False
                vnSQLTrans.Commit()
                vnSQLTrans = Nothing

                TxtOID.Text = vnHOID

                psPreview(TxtBrgCode.Text, TxtPrintSN.Text, vnSQLConn)
            Else
                LblMsgError.Text = "Print Gagal..." & vbCrLf & pbMsgError
                LblMsgError.Visible = True

                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans = Nothing
                End If
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

    Private Sub psPreview(vriBarangCode As String, vriSN As String, vriSQLConn As SqlConnection)
        psClearMessage()
        If Session("UserID") = "" Then Response.Redirect("Default.aspx", False)
        Dim vnCrpFileName As String = ""
        psGenerateCrp(vnCrpFileName, vriBarangCode, vriSN, vriSQLConn)

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

        vbuCrpOnClose = "0"

        ifrPreview.Src = vnRootURL & "Preview/WbfCrpViewer.aspx?" & vnParam
        psShowPreview(True)
    End Sub

    Private Sub psGenerateCrp(ByRef vriCrpFileName As String, vriBarangCode As String, vriSN As String, vriSQLConn As SqlConnection)
        'Barcode print 2x (printer Barcode)
        Dim vnUserOID As String = Session("UserOID")
        Dim vnQuery As String
        Dim vn As Integer
        Dim vnCount As Integer = 1

        Dim vnSQLTrans As SqlTransaction = Nothing

        Try
            vnSQLTrans = vriSQLConn.BeginTransaction("inp")
            vnQuery = "Delete Sys_SsoPrintQRBarang_Temp Where UserOID=" & vnUserOID
            pbuExecuteSQLTrans(vnQuery, cbuActionDel, vriSQLConn, vnSQLTrans)

            For vn = 0 To vnCount - 1
                vnQuery = "Insert into Sys_SsoPrintQRBarang_Temp("
                vnQuery += vbCrLf & "OID,UserOID,BcProductGenCode011,BcProductGenCodeImg011,BcProductGenCode012,BcProductGenCodeImg012)"
                vnQuery += vbCrLf & "Select " & vn & "," & vnUserOID & ",'" & vriBarangCode & "',SNQRCodeImg,'" & vriBarangCode & "',SNQRCodeImg"
                vnQuery += vbCrLf & "       From " & fbuGetDBMaster() & "Sys_MstBarangQRCodeSN_MA Where BRGCODE='" & vriBarangCode & "' and SNQRCode='" & vriSN & "'"
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
    Private Function fsGenBrgSNQRCode(vriCompanyCode As String, vriBarangCode As String, vriSN As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        Dim vnReturn As Boolean
        pbMsgError = ""
        Try
            Dim vnQuery As String
            vnQuery = "Select count(1) From " & fbuGetDBMaster() & "Sys_MstBarangQRCodeSN_MA Where CompanyCode='" & vriCompanyCode & "' and BRGCODE='" & vriBarangCode & "' and SNQRCode='" & vriSN & "'"
            If fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = 0 Then
                Dim vsIOFileStream As System.IO.FileStream
                Dim vsFileLength As Long
                Const csFileFormat = ".jpg"

                Dim vnCmd As SqlCommand
                Dim vnFileName As String
                Dim vnFileByte() As Byte

                vnFileName = vriBarangCode & "_SN_UserOID_" & Session("UserOID") & "_" & Format(Date.Now, "yyyyMMdd_HHmmss") & "~sm" & csFileFormat

                Dim vnQRDir As String = ""

                pbuGenerateQRCode(vnFileName, vriSN, vnQRDir)

                vsIOFileStream = System.IO.File.OpenRead(vnQRDir & vnFileName)

                vsFileLength = vsIOFileStream.Length
                ReDim vnFileByte(vsFileLength)

                vsIOFileStream.Read(vnFileByte, 0, vsFileLength)

                vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarangQRCodeSN_MA"
                vnQuery += vbCrLf & "(CompanyCode,BRGCODE,SNQRCode,SNQRCodeImg,CreationUserOID,CreationDatetime)"
                vnQuery += vbCrLf & "Values("
                vnQuery += vbCrLf & "'" & vriCompanyCode & "','" & vriBarangCode & "','" & vriSN & "',@vnSNQRCodeImg," & Session("UserOID") & ",getdate()"
                vnQuery += vbCrLf & ")"

                vnCmd = New SqlClient.SqlCommand(vnQuery, vriSQLConn)
                vnCmd.Parameters.AddWithValue("@vnSNQRCodeImg", vnFileByte)
                vnCmd.Transaction = vriSQLTrans
                vnCmd.ExecuteNonQuery()

                vnReturn = True
            Else
                vnReturn = True
            End If
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

        psShowLsBrg(True)
    End Sub

    Private Sub BtnPreviewClose_Click(sender As Object, e As EventArgs) Handles BtnPreviewClose.Click
        vbuCrpOnClose = "1"
        psShowPreview(False)
    End Sub
End Class