Imports System.Data.SqlClient
Imports Spire.Barcode
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.IO
Public Class WbfSsoProductQR
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoProductQR"

    Dim settings As BarcodeSettings

    Dim vsIOFileStream As System.IO.FileStream
    Dim vsFileLength As Long

    Dim vsQRDir As String

    Const csFileFormat = ".jpg"
    Private Sub psClearData()
        DstGudang.Text = ""
        TxtBrgCode.Text = ""
        TxtBrgName.Text = ""
        TxtBrgSNEnd.Text = ""
        TxtBrgSNStart.Text = ""
        TxtSNCount.Text = ""
        TxtOID.Text = ""
    End Sub

    Private Sub psClearMessage()
        LblMsgBrgName.Visible = False
        LblMsgGudang.Visible = False
        LblMsgCount.Visible = False
        LblMsgError.Visible = False
    End Sub

    Private Sub psDisplayData()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        If TxtOID.Text = "" Then Exit Sub

        vnQuery = "Select PM.*,BM.BRGNAME"
        vnQuery += vbCrLf & "From " & fbuGetDBMaster() & "Sys_MstBrgQRGen_TR PM"
        vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA BM on BM.CompanyCode=PM.CompanyCode and BM.BRGCODE=PM.BRGCODE"
        vnQuery += vbCrLf & "Where PM.OID=" & TxtOID.Text
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        If vnDtb.Rows.Count = 0 Then
            psClearData()
        Else
            TxtOID.Text = fbuValStr(vnDtb.Rows(0).Item("OID"))
            TxtBrgCode.Text = vnDtb.Rows(0).Item("BRGCODE")
            TxtBrgName.Text = vnDtb.Rows(0).Item("BRGNAME")
            TxtSNCount.Text = vnDtb.Rows(0).Item("QRGenCount")
            TxtBrgSNStart.Text = vnDtb.Rows(0).Item("QRGenSNStart")
            TxtBrgSNEnd.Text = vnDtb.Rows(0).Item("QRGenSNEnd")
            DstGudang.SelectedValue = vnDtb.Rows(0).Item("GdgCode")
        End If
        vnDtb.Dispose()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psEnableInput(vriBo As Boolean)
        TxtSNCount.ReadOnly = Not vriBo
        DstGudang.Enabled = vriBo
        BtnBrgCode.Enabled = vriBo
        BtnBrgCode.Visible = BtnBrgCode.Enabled
    End Sub

    Private Sub psEnableSave(vriBo As Boolean)
        BtnSimpan.Visible = vriBo
        BtnBatal.Visible = vriBo
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

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPrintQRBarang, vnSQLConn)

            pbuFillDstGudang(DstGudang, False, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            BtnBaru.Enabled = (Session("UserAdmin") = 1)
        End If
    End Sub

    Private Sub psDefaultDisplay()
        DivLsBrg.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanLsBrg.Style(HtmlTextWriterStyle.Position) = "absolute"
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
        vnQuery = "Select PM.OID,PM.CompanyCode,PM.GdgCode,PM.BRGCODE,BM.BRGNAME,PM.QRGenCount,QRGenSNStart,QRGenSNEnd,QRGenDatetime,"
        vnQuery += vbCrLf & "      GM.UserName vQRGenUserName"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstBrgQRGen_TR PM"
        vnQuery += vbCrLf & "      inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA BM on BM.CompanyCode=PM.CompanyCode and BM.BRGCODE=PM.BRGCODE"
        vnQuery += vbCrLf & "      	  inner join Sys_SsoUser_MA GM on GM.OID=PM.QRGenUserOID"
        vnQuery += vbCrLf & "Where PM.CompanyCode='" & Session("UserCompanyCode") & "' and (PM.BRGCODE like '%" & Trim(TxtKriteria.Text) & "%' OR BM.BRGNAME like '%" & Trim(TxtKriteria.Text) & "%')"
        vnQuery += vbCrLf & "Order by BM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If e.CommandName = "Select" Then
            Dim vnValue As String = ""
            Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
            Dim vnRow As GridViewRow = GrvList.Rows(vnIdx)

            TxtOID.Text = vnRow.Cells(0).Text
            psDisplayData()
        End If
    End Sub

    Protected Sub BtnLsBrg_Click(sender As Object, e As EventArgs) Handles BtnLsBrg.Click
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
        vnQuery += vbCrLf & "Where CompanyCode='" & Session("UserCompanyCode") & "' and (PM.BRGCODE like '%" & Trim(TxtLsBrg.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtLsBrg.Text) & "%')"
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
            psShowLsBrg(False)
        End If
    End Sub

    Private Sub BtnBaru_Click(sender As Object, e As EventArgs) Handles BtnBaru.Click
        psClearData()
        psEnableInput(True)
        psEnableSave(True)
        HdfActionStatus.Value = cbuActionNew
    End Sub

    Protected Sub BtnBatal_Click(sender As Object, e As EventArgs) Handles BtnBatal.Click
        psClearMessage()
        psEnableInput(False)
        psEnableSave(False)
        HdfActionStatus.Value = cbuActionNorm
        If TxtOID.Text = "" Then
            psClearData()
        Else
            psDisplayData()
        End If
    End Sub

    Protected Sub BtnSimpan_Click(sender As Object, e As EventArgs) Handles BtnSimpan.Click
        Spire.Barcode.BarcodeSettings.ApplyKey("M6XLO-2DRY1-WQ8DF-4DAP6-VOT0X")

        Dim vnSave As Boolean = True
        psClearMessage()

        If Len(Trim(TxtBrgCode.Text)) = 0 Then
            LblMsgBrgName.Text = "Pilih Barang"
            LblMsgBrgName.Visible = True
            vnSave = False
        End If
        If Val(Trim(TxtSNCount.Text)) = 0 Then
            LblMsgCount.Text = "Isi Jumlah"
            LblMsgCount.Visible = True
            vnSave = False
        End If
        If DstGudang.SelectedIndex <= 0 Then
            LblMsgGudang.Text = "Pilih Gudang"
            LblMsgGudang.Visible = True
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
            Dim vnCompanyCode As String = Session("UserCompanyCode")
            Dim vnGdgCode As String = DstGudang.SelectedValue
            Dim vnGdgName As String
            Dim vnBrgCode As String = TxtBrgCode.Text
            Dim vnBrgName As String

            Dim vnYMD As String
            Dim vnSNStart As Integer
            Dim vnSNEnd As Integer
            Dim vnSNCount As Integer = TxtSNCount.Text

            vnQuery = "Select convert(varchar(8),getdate(),112)"
            vnYMD = fbuGetDataStrSQL(vnQuery, vnSQLConn)

            vnQuery = "Select isnull(max(OID),0)+1 from " & fbuGetDBMaster() & "Sys_MstBrgQRGen_TR"
            vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn)

            vnQuery = "Select BRGNAME from " & fbuGetDBMaster() & "Sys_MstBarang_MA Where BRGCODE='" & vnBrgCode & "'"
            vnBrgName = fbuGetDataStrSQL(vnQuery, vnSQLConn)

            vnQuery = "Select GdgName from " & fbuGetDBMaster() & "Sys_MstGudang_MA Where GdgCode='" & vnGdgCode & "'"
            vnGdgName = fbuGetDataStrSQL(vnQuery, vnSQLConn)

            vnQuery = "Select isnull(max(substring(BRGSerialNo,len(BRGSerialNo)-5,10)),0)+1 From " & fbuGetDBMaster() & "Sys_MstBrgQRGenData_TR"
            vnQuery += vbCrLf & "Where BRGCODE='" & vnBrgCode & "' and substring(BRGSerialNo,1,len(BRGCODE)+8) = BRGCODE + '" & vnYMD & "'"
            vnSNStart = fbuGetDataNumSQL(vnQuery, vnSQLConn)
            vnSNEnd = vnSNStart + vnSNCount - 1

            vnSQLTrans = vnSQLConn.BeginTransaction()

            vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBrgQRGen_TR("
            vnQuery += vbCrLf & "OID,CompanyCode,GdgCode,BRGCODE,"
            vnQuery += vbCrLf & "QRGenCount,QRGenSNStart,QRGenSNEnd,"
            vnQuery += vbCrLf & "QRGenDatetime,QRGenUserOID)"
            vnQuery += "values(" & vnHOID & ",'" & vnCompanyCode & "','" & vnGdgCode & "','" & vnBrgCode & "',"
            vnQuery += vbCrLf & vnSNCount & ","
            vnQuery += vbCrLf & "'" & vnBrgCode & vnYMD & Format(vnSNStart, "00000#") & "',"
            vnQuery += vbCrLf & "'" & vnBrgCode & vnYMD & Format(vnSNEnd, "00000#") & "',"
            vnQuery += vbCrLf & "getdate()," & vnUserOID & ")"
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            psGenerateSN(vnHOID, vnGdgCode, vnGdgName, vnBrgCode, vnBrgName, vnYMD, vnSNCount, vnSNStart, vnSNEnd, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans = Nothing

            psEnableInput(False)
            psEnableSave(False)
            HdfActionStatus.Value = cbuActionNorm

            TxtOID.Text = vnHOID

            psDisplayData()

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

    Private Sub psGenerateSN_20221207_Orig_Tanpa_Image(vriHOID As String, vriGdgCode As String, vriBrgCode As String, vriYMD As String, vriSNCount As Integer, vriSNStart As Integer, vriSNEnd As Integer, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        Dim vnSerialNo As String
        Dim vnQRCode As String
        Dim vn As Integer

        For vn = vriSNStart To vriSNEnd
            vnSerialNo = vriBrgCode & vriYMD & Format(vn, "00000#")
            vnQRCode = vnSerialNo & vbCrLf & "Gudang : " & vriGdgCode
            vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBrgQRGenData_TR"
            vnQuery += vbCrLf & "(QRGenHOID,BRGCODE,BRGSerialNo,BRGQRCode)"
            vnQuery += vbCrLf & "Values("
            vnQuery += vbCrLf & vriHOID & ",'" & vriBrgCode & "','" & vnSerialNo & "','" & vnQRCode & "'"
            vnQuery += vbCrLf & ")"
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
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
        psShowLsBrg(True)
    End Sub
End Class