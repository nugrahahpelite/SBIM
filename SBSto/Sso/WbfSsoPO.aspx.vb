Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoPO
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoPO"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsTextStream_Data As Scripting.TextStream
    Dim vsFso_Data As Scripting.FileSystemObject

    Dim vsProcessDate As String
    Dim vsLogFolder As String
    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameOnly As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String

    Enum ensColPOH
        CompanyCode = 0
        PO_NO = 1
        vPO_DATE = 2
        vETA_DATE = 3
        vSupplier = 4
        vPLExist = 5
        vGRExist = 6
        TransStatusDescr = 7
        TransStatus = 8
        POHOID = 9
    End Enum

    Enum ensColPLH
        OID = 0
    End Enum

    Enum ensColRcvPOH
        OID = 0
        RcvPONo = 1
        vRcvPODate = 2
        RcvPORefNo = 3
        RcvPOTypeName = 4
        WarehouseName = 5
        TransStatus = 6
        TransStatusDescr = 7
        RcvPORefTypeOID = 8
        RcvPORefOID = 9
    End Enum

    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Sso"

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoPOPembelian, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, True, vnSQLConn)
                pbuFillDstCompany(DstPOCompany, True, vnSQLConn)
                pbuFillDstCompany(DstPOSapCompany, True, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, True, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstPOCompany, True, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstPOSapCompany, True, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub
    Private Sub psDefaultDisplay()
        DivPOHEta.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPOHEta.Style(HtmlTextWriterStyle.Position) = "absolute"

        DivPOHClo.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        PanPOHClo.Style(HtmlTextWriterStyle.Position) = "absolute"
    End Sub
    Private Sub psShowPOHEta(vriBo As Boolean)
        If vriBo Then
            DivPOHEta.Style(HtmlTextWriterStyle.Visibility) = "visible"
            BtnPOHEtaYes.Visible = True
        Else
            DivPOHEta.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub psShowPOHClo(vriBo As Boolean)
        If vriBo Then
            DivPOHClo.Style(HtmlTextWriterStyle.Visibility) = "visible"
            BtnPOHCloYes.Visible = True
        Else
            DivPOHClo.Style(HtmlTextWriterStyle.Visibility) = "hidden"
        End If
    End Sub
    Private Sub GrvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvList.PageIndexChanging
        GrvList.PageIndex = e.NewPageIndex
        psFillGrvList()
    End Sub

    Protected Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgFindError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgFindError.Visible = True
            Exit Sub
        End If
        psFillGrvList()
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Text = ""
        LblMsgError.Text = ""
        LblMsgFupXls.Text = ""
        LblMsgXlsProsesError.Text = ""
        LblMsgXlsWorksheet.Text = ""
        LblXlsProses.Text = ""
        LblPOSapError.Text = ""
        LblPOSapCompany.Text = ""
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
        vnQuery = "Select PM.OID,PM.CompanyCode,PM.XlsFileName,PM.UploadStartDatetime,PM.UploadEndDatetime,MU.UserName vUploadBy"
        vnQuery += vbCrLf & " From Sys_SsoPOFileXls_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner Join Sys_SsoUser_MA MU with(nolock) on MU.OID=PM.UploadUserOID"
        vnQuery += vbCrLf & "Where UploadSourceOID=" & enuUploadSource.Xls
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If
        vnQuery += vbCrLf & " Order by PM.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvPOSap()
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
        vnQuery = "Select PM.OID,PM.CompanyCode,PM.XlsFileName,PM.UploadStartDatetime,PM.UploadEndDatetime,MU.UserName vUploadBy,case when PM.StatusSuccess=1 then 'Y' else 'N' end vStatusSuccess,StatusMessage"
        vnQuery += vbCrLf & " From Sys_SsoPOFileXls_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner Join Sys_SsoUser_MA MU with(nolock) on MU.OID=PM.UploadUserOID"
        vnQuery += vbCrLf & "Where UploadSourceOID=" & enuUploadSource.SAP_Api
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If
        vnQuery += vbCrLf & " Order by PM.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvPOSap.DataSource = vnDtb
        GrvPOSap.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnXlsUpload_Click(sender As Object, e As EventArgs) Handles BtnXlsUpload.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Upload_Xls) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            Exit Sub
        End If
        If TxtXlsWorksheet.Text = "" Then
            LblMsgXlsWorksheet.Text = "Isi Nama Worksheet"
            Exit Sub
        End If

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "BtnXlsUpload_Click", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        vsXlsFolder = Server.MapPath("~") & "\XlsFolder\"
        vsXlsFileName = vsXlsFolder & stuTransCode.SsoPOPembelian & "_" & Format(Date.Now, "yyyyMMdd_HHmmss ") & FupXls.FileName
        vsTextStream.WriteLine("vnFileName : " & vsXlsFileName)

        vsSheetName = Trim(TxtXlsWorksheet.Text)
        vsTextStream.WriteLine("vnSheetName : " & vsSheetName)

        vsTextStream.WriteLine("FupXls.SaveAs(" & vsXlsFileName & ")...Start")

        FupXls.SaveAs(vsXlsFileName)

        vsTextStream.WriteLine("FupXls.SaveAs(" & vsXlsFileName & ")...End")

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgXlsProsesError.Text = pbMsgError
            LblMsgXlsProsesError.Visible = True

            vsTextStream.WriteLine("Error Open Koneksi SQLServer :")
            vsTextStream.WriteLine(pbMsgError)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("------------------------EOF------------------------")
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            psCreateTable_Sys_SsoPO_Temp(vnSQLConn)

            vnSQLTrans = vnSQLConn.BeginTransaction("xls")
            vnBeginTrans = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------")
            vsTextStream.WriteLine("fsXlsImportData...Start")
            If fsXlsImportData(vnSQLConn, vnSQLTrans) Then
                vsTextStream.WriteLine("fsXlsImportData...End")
                vsTextStream.WriteLine("=======================")
                vsTextStream.WriteLine("")

                vnBeginTrans = False
                vnSQLTrans.Commit()
            Else
                vsTextStream.WriteLine("fsXlsImportData...Gagal")
                vsTextStream.WriteLine("=======================")
                vsTextStream.WriteLine("")

                vnBeginTrans = False
                vnSQLTrans.Rollback()
            End If

            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psDropTable_Sys_SsoPO_Temp(vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("------------------------EOF------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            LblXlsProses.Visible = True

            psFillGrvList()

        Catch ex As Exception
            LblMsgXlsProsesError.Text = ex.Message
            LblMsgXlsProsesError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("PROCESS TERMINATED...ERROR :")
            vsTextStream.WriteLine(ex.Message)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------ERROR-----------------------")
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

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


    Private Function fsXlsImportData(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData = False
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        Const cnColOBL = 0
        Const cnColJURNAL = 1
        Const cnColDATE_ = 2
        Const cnColSUB = 3
        Const cnColBRG = 4
        Const cnColGDGOBL = 5
        Const cnColQTY = 6
        Const cnColNAMA_SUPPLIER = 12
        Const cnColNAMA_BARANG = 13
        Const cnColJOBNAME = 14

        Dim vnColOBL As String
        Dim vnColJURNAL As String
        Dim vnColDATE_ As String
        Dim vnColSUB As String
        Dim vnColBRG As String
        Dim vnColGDGOBL As String
        Dim vnColQTY As String
        Dim vnColNAMA_SUPPLIER As String
        Dim vnColNAMA_BARANG As String
        Dim vnColJOBNAME As String

        Dim vnQuery As String

        vnQuery = "Delete #Sys_SsoPO_Temp"
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Select max(OID) From Sys_SsoPOFileXls_TR with(nolock)"
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)

        Dim vnGetDate As String = fbuGetDateNowSQLTrans(vriSQLConn, vriSQLTrans)
        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) + 1

        vnQuery = "Insert into Sys_SsoPOFileXls_TR"
        vnQuery += vbCrLf & "(OID,CompanyCode,UploadSourceOID,XlsFileName,UploadStartDatetime,UploadUserOID)"
        vnQuery += vbCrLf & "Values"
        vnQuery += vbCrLf & "(" & vnHOID & ",'" & vnCompanyCode & "'," & enuUploadSource.Xls & ",'" & vsXlsFileName & "','" & vnGetDate & "'," & Session("UserOID") & ")"
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        Dim vnPath As String = System.IO.Path.GetFullPath(vsXlsFileName)
        vsTextStream.WriteLine("vnPath : " & vnPath)

        Dim vnXConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vsXlsFileName & ";" & "Extended Properties=Excel 12.0 Xml;" ';IMEX=1;HDR=YES
        vsTextStream.WriteLine("vnXConnStr : " & vnXConnStr)

        Dim vnXConn As New OleDbConnection(vnXConnStr)
        vnXConn.Open()

        Dim vnXCommand As OleDbCommand
        vnXCommand = vnXConn.CreateCommand()
        vnXCommand.CommandText = "Select * From [" & vsSheetName & "$]"
        vsTextStream.WriteLine("")

        Dim vnXReader As OleDbDataReader
        vnXReader = vnXCommand.ExecuteReader

        Dim vnNo As Integer
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnCompanyCode = " & vnCompanyCode)
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")

        Dim vnReturn As Boolean = True
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnColOBL = fbuValStr(vnXReader.Item(cnColOBL))
            vsTextStream.WriteLine(vbCrLf & "vnColOBL " & vnColOBL)

            vnColJURNAL = fbuValStr(vnXReader.Item(cnColJURNAL))
            vsTextStream.WriteLine(vbCrLf & "vnColJURNAL " & vnColJURNAL)

            'vnColDATE_ = fbuFormatDateDMY_To_YMD(fbuValStr(vnXReader.Item(cnColDATE_)))
            vnColDATE_ = Convert.ToDateTime(vnXReader.Item(cnColDATE_)).ToString("yyyy-MM-dd")
            vsTextStream.WriteLine(vbCrLf & "vnColDATE_ " & vnColDATE_)

            vnColSUB = fbuValStr(vnXReader.Item(cnColSUB))
            vsTextStream.WriteLine(vbCrLf & "vnColSUB " & vnColSUB)

            vnColBRG = fbuValStr(vnXReader.Item(cnColBRG))
            vsTextStream.WriteLine(vbCrLf & "vnColBRG " & vnColBRG)

            vnColGDGOBL = fbuFormatString(fbuValStr(vnXReader.Item(cnColGDGOBL)))
            vsTextStream.WriteLine(vbCrLf & "vnColGDGOBL " & vnColGDGOBL)

            vnColQTY = fbuValStr(vnXReader.Item(cnColQTY))
            vsTextStream.WriteLine(vbCrLf & "vnColQTY " & vnColQTY)

            vnColNAMA_SUPPLIER = fbuFormatString(fbuValStr(vnXReader.Item(cnColNAMA_SUPPLIER)))
            vsTextStream.WriteLine(vbCrLf & "vnColNAMA_SUPPLIER " & vnColNAMA_SUPPLIER)

            vnColNAMA_BARANG = fbuFormatString(fbuValStr(vnXReader.Item(cnColNAMA_BARANG)))
            vsTextStream.WriteLine(vbCrLf & "vnColNAMA_BARANG " & vnColNAMA_BARANG)

            vnColJOBNAME = fbuValStr(vnXReader.Item(cnColJOBNAME))
            vsTextStream.WriteLine(vbCrLf & "vnColJOBNAME " & vnColJOBNAME)

            vnQuery = "Insert into #Sys_SsoPO_Temp"
            vnQuery += vbCrLf & "(CompanyCode,PO_NO,JURNAL,PO_DATE,SUB,BRG_ORIG,BRG,GDGCODE,"
            vnQuery += vbCrLf & "QTY,NAMA_SUPPLIER,NAMA_BARANG,JOBNAME,"
            vnQuery += vbCrLf & "POFileXlsOID,UploadSourceOID,UploadDatetime"
            vnQuery += vbCrLf & ")"

            vnQuery += vbCrLf & "Select '" & vnCompanyCode & "'CompanyCode,'" & vnColOBL & "'PO_NO,'" & vnColJURNAL & "'JURNAL,'" & vnColDATE_ & "'PO_DATE,'" & vnColSUB & "'SUB,'" & vnColBRG & "'BRG_ORIG,'" & vnColBRG & "'BRG,'" & vnColGDGOBL & "'GDGCODE,"
            vnQuery += vbCrLf & "'" & vnColQTY & "'QTY,'" & vnColNAMA_SUPPLIER & "'NAMA_SUPPLIER,'" & vnColNAMA_BARANG & "'NAMA_BARANG,'" & vnColJOBNAME & "'JOBNAME,"
            vnQuery += vbCrLf & vnHOID & "," & enuUploadSource.Xls & ",'" & vnGetDate & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        vnQuery = "Delete #Sys_SsoPO_Temp Where isnull(JURNAL,'')<>''"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Delete POD"
        vnQuery += vbCrLf & "       From Sys_SsoPODetail_TR POD"
        vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
        vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO AND ABT.BRG=POD.BRG"
        vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Delete POD"
        vnQuery += vbCrLf & "       From Sys_SsoPO POD"
        vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
        vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO AND ABT.BRG=POD.BRG"
        vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Delete POD"
        vnQuery += vbCrLf & "       From Sys_SsoPOStatus_TR POD"
        vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
        vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO"
        vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Delete POH"
        vnQuery += vbCrLf & "       From Sys_SsoPOHeader_TR POH"
        vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO"
        vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into Sys_SsoPO"
        vnQuery += vbCrLf & "Select *,0 POHOID From #Sys_SsoPO_Temp ABT with(nolock) WHERE NOT ABT.PO_NO+ABT.BRG IN"
        vnQuery += vbCrLf & "	   (Select AB.PO_NO+AB.BRG FROM Sys_SsoPO AB)"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        Dim vnPOHOID As Integer
        vnQuery = "Select isnull(max(OID),0) From Sys_SsoPOHeader_TR with(nolock)"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        vnPOHOID = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        Dim vnDtbPOH As New DataTable
        vnQuery = "Select distinct CompanyCode,PO_NO From Sys_SsoPO Where POHOID=0"
        pbuFillDtbSQLTrans(vnDtbPOH, vnQuery, vriSQLConn, vriSQLTrans)
        For vn = 0 To vnDtbPOH.Rows.Count - 1
            vnPOHOID = vnPOHOID + 1
            vnQuery = "Insert into Sys_SsoPOHeader_TR(OID,CompanyCode,PO_NO,PO_DATE,SUB,NAMA_SUPPLIER,GDGCODE,JOBNAME,TransCode,TransStatus,"
            vnQuery += vbCrLf & "UploadSourceOID,UploadDatetime)"
            vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",CompanyCode,PO_NO,PO_DATE,SUB,NAMA_SUPPLIER,GDGCODE,JOBNAME,'" & stuTransCode.SsoPOPembelian & "'TransCode," & enuTCSPPO.Baru & " TransStatus,"
            vnQuery += vbCrLf & "       UploadSourceOID,UploadDatetime"
            vnQuery += vbCrLf & " From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Update Sys_SsoPO set POHOID=" & vnPOHOID & "Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into Sys_SsoPOStatus_TR(POHOID,TransCode,TransStatus,TransStatusDatetime)"
            vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",'" & stuTransCode.SsoPOPembelian & "'TransCode," & enuTCSPPO.Baru & " TransStatus,'" & vnGetDate & "'"
            vnQuery += vbCrLf & "From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into Sys_SsoPODetail_TR(POHOID,BRG_ORIG,BRG,NAMA_BARANG,QTY)"
            vnQuery += vbCrLf & "Select " & vnPOHOID & ",BRG,BRG,NAMA_BARANG,QTY"
            vnQuery += vbCrLf & "From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        Next

        vsTextStream.WriteLine(vbCrLf & "")
        vsTextStream.WriteLine(vbCrLf & "Commit Transaction")

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData = True
    End Function

    Private Function fsImportDataPO_Step_2(vriGetDate As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsImportDataPO_Step_2 = False
        Try
            Dim vnQuery As String
            vnQuery = "Delete #Sys_SsoPO_Temp Where isnull(JURNAL,'')<>''"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Delete POD"
            vnQuery += vbCrLf & "       From Sys_SsoPODetail_TR POD"
            vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
            vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO AND ABT.BRG=POD.BRG"
            vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Delete POD"
            vnQuery += vbCrLf & "       From Sys_SsoPO POD"
            vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
            vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO AND ABT.BRG=POD.BRG"
            vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Delete POD"
            vnQuery += vbCrLf & "       From Sys_SsoPOStatus_TR POD"
            vnQuery += vbCrLf & "	         Inner Join Sys_SsoPOHeader_TR POH ON POH.OID=POD.POHOID"
            vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO"
            vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Delete POH"
            vnQuery += vbCrLf & "       From Sys_SsoPOHeader_TR POH"
            vnQuery += vbCrLf & "	         Inner Join #Sys_SsoPO_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.PO_NO=POH.PO_NO"
            vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCSPPO.Baru
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into Sys_SsoPO"
            vnQuery += vbCrLf & "Select *,0 POHOID From #Sys_SsoPO_Temp ABT with(nolock) WHERE NOT ABT.PO_NO+ABT.BRG IN"
            vnQuery += vbCrLf & "	   (Select AB.PO_NO+AB.BRG FROM Sys_SsoPO AB)"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            Dim vnPOHOID As Integer
            vnQuery = "Select isnull(max(OID),0) From Sys_SsoPOHeader_TR with(nolock)"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            vnPOHOID = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

            Dim vnDtbPOH As New DataTable
            vnQuery = "Select distinct CompanyCode,PO_NO From Sys_SsoPO Where POHOID=0"
            pbuFillDtbSQLTrans(vnDtbPOH, vnQuery, vriSQLConn, vriSQLTrans)
            For vn = 0 To vnDtbPOH.Rows.Count - 1
                vnPOHOID = vnPOHOID + 1
                vnQuery = "Insert into Sys_SsoPOHeader_TR("
                vnQuery += vbCrLf & "OID,CompanyCode,PO_NO,PO_DATE,SUB,NAMA_SUPPLIER,GDGCODE,JOBNAME,TransCode,TransStatus,"
                vnQuery += vbCrLf & "UploadSourceOID,UploadDatetime,"
                vnQuery += vbCrLf & "SAP_DocEntry,SAP_DocType,SAP_DocumentStatus,SAP_Cancelled)"
                vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",CompanyCode,PO_NO,PO_DATE,SUB,NAMA_SUPPLIER,GDGCODE,JOBNAME,'" & stuTransCode.SsoPOPembelian & "'TransCode," & enuTCSPPO.Baru & " TransStatus,"
                vnQuery += vbCrLf & "       UploadSourceOID,UploadDatetime,"
                vnQuery += vbCrLf & "       SAP_DocEntry,SAP_DocType,SAP_DocumentStatus,SAP_Cancelled"
                vnQuery += vbCrLf & "  From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                vnQuery = "Update Sys_SsoPO set POHOID=" & vnPOHOID & "Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                vnQuery = "Insert into Sys_SsoPOStatus_TR(POHOID,TransCode,TransStatus,TransStatusDatetime)"
                vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",'" & stuTransCode.SsoPOPembelian & "'TransCode," & enuTCSPPO.Baru & " TransStatus,'" & vriGetDate & "'"
                vnQuery += vbCrLf & "From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                vnQuery = "Insert into Sys_SsoPODetail_TR(POHOID,BRG_ORIG,BRG,NAMA_BARANG,QTY,SAP_LineNum)"
                vnQuery += vbCrLf & "Select " & vnPOHOID & ",BRG,BRG,NAMA_BARANG,QTY,SAP_LineNum"
                vnQuery += vbCrLf & "From Sys_SsoPO Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and PO_NO='" & vnDtbPOH.Rows(vn).Item("PO_NO") & "'"
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
            Next

            vsTextStream.WriteLine("fsImportDataPO_Step_2 = True")

            fsImportDataPO_Step_2 = True

        Catch ex As Exception

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("ERROR")
            vsTextStream.WriteLine("fsImportDataPO_Step_2 = False")
            vsTextStream.WriteLine("ex.Message")
            vsTextStream.WriteLine(ex.Message)
            fsImportDataPO_Step_2 = False
        End Try
    End Function

    Private Sub psCreateTable_Sys_SsoPO_Temp(vriSQLConn As SqlConnection)
        Dim vnQuery As String
        vnQuery = "CREATE Table #Sys_SsoPO_Temp("
        vnQuery += vbCrLf & "[CompanyCode] [varchar](15) Not NULL,"
        vnQuery += vbCrLf & "[PO_NO] [varchar](50) Not NULL,"
        vnQuery += vbCrLf & "[JURNAL] [varchar](50) NULL,"
        vnQuery += vbCrLf & "[PO_DATE] [DateTime] Not NULL,"
        vnQuery += vbCrLf & "[SUB] [varchar](50) Not NULL,"
        vnQuery += vbCrLf & "[BRG_ORIG] [varchar](450) Not NULL,"
        vnQuery += vbCrLf & "[BRG] [varchar](450) Not NULL,"
        vnQuery += vbCrLf & "[GDGCODE] [varchar](50) Not NULL,"
        vnQuery += vbCrLf & "[QTY] [numeric](18, 0) Not NULL,"
        vnQuery += vbCrLf & "[NAMA_SUPPLIER] [varchar](50) Not NULL,"
        vnQuery += vbCrLf & "[NAMA_BARANG] [varchar](450) Not NULL,"
        vnQuery += vbCrLf & "[JOBNAME] [varchar](15) Not NULL,"

        vnQuery += vbCrLf & "[SAP_DocEntry] [int],"
        vnQuery += vbCrLf & "[SAP_DocType] [varchar](50),"
        vnQuery += vbCrLf & "[SAP_DocumentStatus] [varchar](50),"
        vnQuery += vbCrLf & "[SAP_Cancelled] [varchar](50),"
        vnQuery += vbCrLf & "[SAP_LineNum] [int],"
        vnQuery += vbCrLf & "[UploadSourceOID] [tinyint] Not NULL,"

        vnQuery += vbCrLf & "[POFileXlsOID] [Int] NULL,"
        vnQuery += vbCrLf & "[UploadDatetime] [DateTime] Not NULL"
        vnQuery += vbCrLf & ")"
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQL(vnQuery, cbuActionNew, vriSQLConn)
    End Sub

    Private Sub psDropTable_Sys_SsoPO_Temp(vriSQLConn As SqlConnection)
        Dim vnQuery As String
        vnQuery = "DROP Table #Sys_SsoPO_Temp"
        pbuExecuteSQL(vnQuery, cbuActionNew, vriSQLConn)
    End Sub
    Protected Sub BtnData_Click(sender As Object, e As EventArgs) Handles BtnData.Click
        PanPOData.Visible = True
        PanPOUpload.Visible = False
        PanPOSAP.Visible = False
    End Sub

    Protected Sub BtnPOUpload_Click(sender As Object, e As EventArgs) Handles BtnPOUpload.Click
        PanPOData.Visible = False
        PanPOUpload.Visible = True
        PanPOSAP.Visible = False
    End Sub

    Protected Sub BtnPOFind_Click(sender As Object, e As EventArgs) Handles BtnPOFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgPOFindError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgPOFindError.Visible = True
            Exit Sub
        End If

        LblMsgPOFindError.Text = ""
        LblMsgPOFindError.Visible = False

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOFindError.Text = pbMsgError
            LblMsgPOFindError.Visible = True
            Exit Sub
        End If

        psFillGrvPOH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvPOH(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String
        Dim vnSupplier As String = fbuFormatString(Trim(TxtPOSupplier.Text))

        If ChkSt_PL_Full.Checked = False And ChkSt_PL_Not.Checked = False And ChkSt_PL_Sebagian.Checked = False Then
            ChkSt_PL_Full.Checked = True
            ChkSt_PL_Not.Checked = True
            ChkSt_PL_Sebagian.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If Not (ChkSt_PL_Full.Checked = True And ChkSt_PL_Not.Checked = True And ChkSt_PL_Sebagian.Checked = True) Then
            If ChkSt_PL_Full.Checked = True Then
                vnCrStatus += "sum(pod.QTY_PL)>=sum(pod.QTY)"
            End If
            If ChkSt_PL_Not.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "sum(pod.QTY_PL)=0"
            End If
            If ChkSt_PL_Sebagian.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "(sum(pod.QTY_PL)>0 and sum(pod.QTY_PL)<sum(pod.QTY))"
            End If
            vnCrStatus = " and OID in (Select pod.POHOID From Sys_SsoPODetail_TR pod with(nolock) group by pod.POHOID having " & vnCrStatus & ")"
        End If

        If ChkPO_Closed.Checked = False And ChkPO_NotClosed.Checked = False Then
            ChkPO_Closed.Checked = False
            ChkPO_NotClosed.Checked = True
        End If

        Dim vnCrClosed As String = ""
        If Not (ChkPO_Closed.Checked = True And ChkPO_NotClosed.Checked = True) Then
            If ChkPO_Closed.Checked = True Then
                vnCrClosed = "=" & enuTCSPPO.Closed
            End If
            If ChkPO_NotClosed.Checked = True Then
                vnCrClosed = "<>" & enuTCSPPO.Closed
            End If
            vnCrClosed = " and poh.TransStatus" & vnCrClosed
        End If

        vnCriteria = "      Where 1=1"
        vnCriteria += vbCrLf & vnCrStatus
        vnCriteria += vbCrLf & vnCrClosed

        If DstPOCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and poh.CompanyCode='" & DstPOCompany.SelectedValue & "'"
        End If
        If Trim(TxtPOSupplier.Text) <> "" Then
            vnCriteria += vbCrLf & "            and (poh.SUB like '%" & vnSupplier & "%' or poh.NAMA_SUPPLIER like '%" & vnSupplier & "%')"
        End If
        If Trim(TxtPONo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and poh.PO_NO like '%" & fbuFormatString(Trim(TxtPONo.Text)) & "%'"
        End If
        If IsDate(TxtPOStart.Text) Then
            vnCriteria += vbCrLf & "            and poh.PO_DATE >= '" & TxtPOStart.Text & "'"
        End If
        If IsDate(TxtPOEnd.Text) Then
            vnCriteria += vbCrLf & "            and poh.PO_DATE <= '" & TxtPOEnd.Text & "'"
        End If

        vnQuery = "Select poh.CompanyCode,poh.PO_NO,convert(varchar(11),poh.PO_DATE,106)vPO_DATE,convert(varchar(11),poh.ETA_DATE,106)vETA_DATE,"
        vnQuery += vbCrLf & "       poh.SUB +' '+poh.NAMA_SUPPLIER vSupplier,"
        vnQuery += vbCrLf & "       poh.vPLExist,poh.vGRExist,"
        vnQuery += vbCrLf & "       st.TransStatusDescr,poh.TransStatus,poh.OID POHOID"
        vnQuery += vbCrLf & "  From fnTbl_SsoPOHeader() poh"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA st with(nolock) on st.TransCode=poh.TransCode and st.TransStatus=poh.TransStatus"

        If vnUserCompanyCode <> "" And DstPOCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=poh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria

        If Trim(TxtPOBrg.Text) <> "" Then
            Dim vnBrg As String = fbuFormatString(Trim(TxtPOBrg.Text))
            vnQuery += vbCrLf & "      and poh.OID in(Select pod.POHOID From Sys_SsoPODetail_TR pod with(nolock) where pod.BRG like '%" & vnBrg & "%' or pod.NAMA_BARANG like '%" & vnBrg & "%')"
        End If
        vnQuery += vbCrLf & "Order by poh.CompanyCode,poh.PO_NO"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvPOH.DataSource = vnDtb
        GrvPOH.DataBind()

        PanPOD.Visible = False
    End Sub
    Private Sub psFillGrvPOD(vriPOHOID As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select poh.CompanyCode,poh.PO_NO,convert(varchar(11),poh.PO_DATE,106)vPO_DATE,poh.SUB,poh.NAMA_SUPPLIER,poh.GDGCODE,"
        vnQuery += vbCrLf & "            pod.BRG,pod.NAMA_BARANG,pod.QTY,pod.QTY_PL,isnull(por.QTY_RCV,0)QTY_RCV,"
        vnQuery += vbCrLf & "            pod.OID vPODOID,pod.POHOID"
        vnQuery += vbCrLf & "       From Sys_SsoPOHeader_TR poh"
        vnQuery += vbCrLf & "            inner join Sys_SsoPODetail_TR pod on pod.POHOID=poh.OID"
        vnQuery += vbCrLf & "            left outer join Sys_SsoPODetailRcv_TR por on por.POHOID=pod.POHOID and por.BRG=pod.BRG"

        If vnUserCompanyCode <> "" And DstPOCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=poh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & "      Where pod.POHOID=" & vriPOHOID
        vnQuery += vbCrLf & "Order by pod.BRG"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvPOD.DataSource = vnDtb
        GrvPOD.DataBind()
    End Sub

    Private Sub psFillGrvPLH(vriPOHOID As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.OID,PM.PLNo,convert(varchar(11),PM.PLDate,106)vPLDate,"
        vnQuery += vbCrLf & "     RC.RcvPONo,convert(varchar(11),RC.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "     ST.TransStatusDescr"

        vnQuery += vbCrLf & "From Sys_SsoPLHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "     inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"

        vnQuery += vbCrLf & "     left outer join Sys_SsoRcvPOHeader_TR RC with(nolock) on RC.RcvPORefOID=PM.OID and RC.RcvRefTypeOID=" & enuRcvType.Pembelian & " and RC.RcvPORefTypeOID=" & enuRcvPOType.Import

        vnQuery += vbCrLf & "Where PM.TransStatus<>" & enuTCPLSP.Cancelled & " and PM.OID in(Select b.PLHOID From Sys_SsoPLDetail_TR b Where b.POHOID=" & vriPOHOID & ")"
        vnQuery += vbCrLf & "Order by PM.PLNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvPLH.DataSource = vnDtb
        GrvPLH.DataBind()

        If GrvPLH.Rows.Count = 0 Then
            psFillGrvPLD(0, vriSQLConn)
        Else
            psFillGrvPLD(GrvPLH.Rows(0).Cells(ensColPLH.OID).Text, vriSQLConn)
        End If
    End Sub

    Private Sub psFillGrvPLD(vriPLHOID As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select pld.OID,pld.BRGCODE,pld.BRGNAME,pld.PLDQty,pld.PLDSet,pld.PLDCtn"
        vnQuery += vbCrLf & "  From Sys_SsoPLDetail_TR pld with(nolock)"

        If vriPLHOID = 0 Then
            vnQuery += vbCrLf & " Where 1=2"
        Else
            vnQuery += vbCrLf & " Where pld.PLHOID=" & vriPLHOID
        End If

        vnQuery += vbCrLf & "Order by pld.BRGCODE"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvPLD.DataSource = vnDtb
        GrvPLD.DataBind()
    End Sub

    Private Sub psFillGrvRcvPOH(vriPOHOID As String, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select PM.OID,PM.RcvPONo,PM.RcvPORefTypeOID,PM.RcvPORefOID,PM.RcvPORefNo,convert(varchar(11),PM.RcvPODate,106)vRcvPODate,"
        vnQuery += vbCrLf & "       PM.RcvPORefNo,RT.RcvPOTypeName,WM.WarehouseName,ST.TransStatusDescr,PM.TransStatus"
        vnQuery += vbCrLf & "  From Sys_SsoRcvPOHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_Warehouse_MA WM with(nolock) on WM.OID=PM.WarehouseOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoRcvPOType_MA RT with(nolock) on RT.OID=PM.RcvPORefTypeOID"
        vnQuery += vbCrLf & "       inner join Sys_SsoTransStatus_MA ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & " Where (PM.RcvPORefTypeOID=" & enuRcvPOType.Import & " and"
        vnQuery += vbCrLf & "                PM.RcvPORefOID in(Select plh.OID From Sys_SsoPLHeader_TR plh"
        vnQuery += vbCrLf & "                                                      inner join Sys_SsoPLDetail_TR pld on pld.PLHOID=plh.OID and plh.TransStatus<>-2 and pld.POHOID=" & vriPOHOID & "))"
        vnQuery += vbCrLf & "       OR"
        vnQuery += vbCrLf & "       (PM.RcvPORefTypeOID=" & enuRcvPOType.Local & " and"
        vnQuery += vbCrLf & "        PM.OID in(Select rcs.RcvPOHOID From Sys_SsoRcvPOScan_TR rcs with(nolock) Where rcs.POHOID=" & vriPOHOID & "))"
        vnQuery += vbCrLf & "Order by PM.RcvPONo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvRcvPOH.DataSource = vnDtb
        GrvRcvPOH.DataBind()

        If GrvRcvPOH.Rows.Count = 0 Then
            psFillGrvRcvPOSumm(0, 0, 0, 1, vriSQLConn)
        Else
            Dim vnGRow As GridViewRow
            vnGRow = GrvRcvPOH.Rows(0)
            psFillGrvRcvPOSumm(vriPOHOID, vnGRow.Cells(ensColRcvPOH.OID).Text, vnGRow.Cells(ensColRcvPOH.TransStatus).Text, vnGRow.Cells(ensColRcvPOH.RcvPORefTypeOID).Text, vriSQLConn)
        End If
    End Sub

    Private Sub psFillGrvRcvPOSumm(vriPOHOID As String, vriRcvPOHOID As String, vriStatus As Integer, vriRcvPORefTypeOID As Byte, vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        If vriStatus >= enuTCRCPO.Receive_Done Then
            vnQuery = "Select rcm.POHOID,PO_NO,rcm.BRGCODE BRG,msb.BRGNAME NAMA_BARANG,rcm.SumPLQty vSumPLQty,rcm.SumPOQty vSumPOQty,0 vSumRetDRealQty,rcm.SumRcvPOScanQty vSumRcvPOScanQty,rcm.RcvPOQty_Total vRcvPOQty_Total,"
            If vriRcvPORefTypeOID = enuRcvPOType.Import Then
                vnQuery += vbCrLf & "    (rcm.SumRcvPOScanQty - rcm.SumPLQty)vQtyVarian"
            Else
                vnQuery += vbCrLf & "    (rcm.SumRcvPOScanQty - rcm.SumPOQty)vQtyVarian"
            End If
            vnQuery += vbCrLf & "From Sys_SsoRcvPOSummaryDone_TR rcm with(nolock)"
            vnQuery += vbCrLf & "     inner join Sys_SsoRcvPOHeader_TR rch with(nolock) on rch.OID=rcm.RcvPOHOID"
            vnQuery += vbCrLf & "	  inner join " & fbuGetDBMaster() & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=rcm.BRGCODE and msb.CompanyCode=rch.RcvPOCompanyCode"
            vnQuery += vbCrLf & "     left outer join Sys_SsoPOHeader_TR poh with(nolock) on poh.OID=rcm.POHOID"
            vnQuery += vbCrLf & "Where rch.OID=" & vriRcvPOHOID

            If vriRcvPORefTypeOID = enuRcvPOType.Import Then
            Else
                vnQuery += vbCrLf & "      and rcm.POHOID=" & vriPOHOID
            End If
            If vriRcvPOHOID = 0 Then
                vnQuery += vbCrLf & "       and 1=2"
            Else
                vnQuery += vbCrLf & "order by case when isnull(rcm.SumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
            End If
        Else
            If vriRcvPORefTypeOID = enuRcvPOType.Import Then
                vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,vSumPLQty,0 vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,0 vRcvPOQty_Total,(vSumRcvPOScanQty - vSumPLQty)vQtyVarian"
                vnQuery += vbCrLf & " From fnTbl_SsoRcvPOImport_SummaryWithOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                vnQuery += vbCrLf & "Where 1=1"
            Else
                vnQuery = "Select POHOID,PO_NO,BRG,NAMA_BARANG,0 vSumPLQty,vSumPOQty,0 vSumRetDRealQty,vSumRcvPOScanQty,vRcvPOQty_Total,(vRcvPOQty_Total - vSumPOQty)vQtyVarian"
                vnQuery += vbCrLf & " From fnTbl_SsoRcvPOLocal_SummaryNonOS(" & vriRcvPOHOID & "," & Session("UserOID") & ")"
                vnQuery += vbCrLf & "Where POHOID=" & vriPOHOID
            End If
            If vriRcvPOHOID = 0 Then
                vnQuery += vbCrLf & "       and 1=2"
            Else
                vnQuery += vbCrLf & "order by case when isnull(vSumRcvPOScanQty,0)=0 then 5 else 1 end,PO_NO,BRG,NAMA_BARANG"
            End If
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvRcvPOSumm.DataSource = vnDtb
        GrvRcvPOSumm.DataBind()
    End Sub
    Private Sub psFillGrvPO_20230526_Orig(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String
        Dim vnSupplier As String = fbuFormatString(Trim(TxtPOSupplier.Text))

        If ChkSt_PL_Full.Checked = False And ChkSt_PL_Not.Checked = False And ChkSt_PL_Sebagian.Checked = False Then
            ChkSt_PL_Full.Checked = True
            ChkSt_PL_Not.Checked = True
            ChkSt_PL_Sebagian.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If Not (ChkSt_PL_Full.Checked = True And ChkSt_PL_Not.Checked = True And ChkSt_PL_Sebagian.Checked = True) Then
            If ChkSt_PL_Full.Checked = True Then
                vnCrStatus += "pod.QTY_PL>=pod.QTY"
            End If
            If ChkSt_PL_Not.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "pod.QTY_PL=0"
            End If
            If ChkSt_PL_Sebagian.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "(pod.QTY_PL>0 and pod.QTY_PL<pod.QTY)"
            End If
            vnCrStatus = " and (" & vnCrStatus & ")"
        End If


        vnCriteria = "      Where 1=1"
        vnCriteria += vbCrLf & vnCrStatus

        If DstPOCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and poh.CompanyCode='" & DstPOCompany.SelectedValue & "'"
        End If
        If Trim(TxtPOSupplier.Text) <> "" Then
            vnCriteria += vbCrLf & "            and (poh.SUB like '%" & vnSupplier & "%' or poh.NAMA_SUPPLIER like '%" & vnSupplier & "%')"
        End If
        If Trim(TxtPONo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and poh.PO_NO like '%" & fbuFormatString(Trim(TxtPONo.Text)) & "%'"
        End If
        If IsDate(TxtPOStart.Text) Then
            vnCriteria += vbCrLf & "            and poh.PO_DATE >= '" & TxtPOStart.Text & "'"
        End If
        If IsDate(TxtPOEnd.Text) Then
            vnCriteria += vbCrLf & "            and poh.PO_DATE <= '" & TxtPOEnd.Text & "'"
        End If

        vnQuery = "Select poh.CompanyCode,poh.PO_NO,convert(varchar(11),poh.PO_DATE,106)vPO_DATE,poh.SUB,poh.NAMA_SUPPLIER,poh.GDGCODE,"
        vnQuery += vbCrLf & "            pod.BRG,pod.NAMA_BARANG,pod.QTY,pod.QTY_PL,"
        vnQuery += vbCrLf & "            pod.OID vPODOID,pod.POHOID"
        vnQuery += vbCrLf & "       From Sys_SsoPOHeader_TR poh"
        vnQuery += vbCrLf & "            inner join Sys_SsoPODetail_TR pod on pod.POHOID=poh.OID"

        If vnUserCompanyCode <> "" And DstPOCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu on mu.CompanyCode=poh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & "Order by poh.CompanyCode,poh.PO_NO,pod.NAMA_BARANG"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvPOD.DataSource = vnDtb
        GrvPOD.DataBind()
    End Sub

    Private Sub GrvPOD_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvPOD.PageIndexChanging
        GrvPOD.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOError.Text = pbMsgError
            LblMsgPOError.Visible = True
            Exit Sub
        End If

        psFillGrvPOD(LblMsgPOHOID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvPOD_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvPOD.SelectedIndexChanged

    End Sub

    Protected Sub GrvPOH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvPOH.SelectedIndexChanged

    End Sub

    Private Sub GrvPOH_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvPOH.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvPOH.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvPOH.Rows(vnIdx)

        If e.CommandName = "PO_NO" Then
            HdfPOHRowIdx.Value = vnIdx

            Dim vnPOHOID As String = vnGRow.Cells(ensColPOH.POHOID).Text
            Dim vnPONo As String = DirectCast(vnGRow.Cells(ensColPOH.PO_NO).Controls(0), LinkButton).Text
            LblMsgPOHOID.Text = vnPOHOID
            HdfPOHOID.Value = LblMsgPOHOID.Text
            HdfPOHStatus.Value = vnGRow.Cells(ensColPOH.TransStatus).Text

            LblMsgPOHNo.Text = vnPONo

            LblMsgPOHSupplier.Text = vnGRow.Cells(ensColPOH.vSupplier).Text
            LblMsgPOHStatus.Text = "Status = " & vnGRow.Cells(ensColPOH.TransStatusDescr).Text
            LblMsgPOHDate.Text = "Tanggal PO = " & vnGRow.Cells(ensColPOH.vPO_DATE).Text
            LblMsgPOHETADate.Text = "ETA = " & fbuValStrHtml(vnGRow.Cells(ensColPOH.vETA_DATE).Text)

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgPOError.Text = pbMsgError
                LblMsgPOError.Visible = True
                Exit Sub
            End If

            If vnGRow.Cells(ensColPOH.TransStatus).Text = enuTCSPPO.Closed Then
                BtnPOHEta.Enabled = False
                BtnPOHClose.Enabled = False

                PanPOHClose.Visible = True

                Dim vnDtb As New DataTable
                Dim vnQuery As String
                vnQuery = "Select usr.UserName + ' ' + convert(varchar(11),ClosedDatetime,106)+' '+convert(varchar(11),ClosedDatetime,108),poh.POCloseNote"
                vnQuery += vbCrLf & "      From Sys_SsoPOHeader_TR poh with(nolock)"
                vnQuery += vbCrLf & "	        inner join Sys_SsoUser_MA usr with(nolock) on usr.OID=poh.ClosedUserOID"
                vnQuery += vbCrLf & "	  Where poh.OID=" & vnPOHOID
                pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

                If vnDtb.Rows.Count > 0 Then
                    LblMsgPOHClose.Text = "Close Info = " & vnDtb.Rows(0).Item(0)
                    TxtPOHCloseNote.Text = vnDtb.Rows(0).Item(1)
                End If
            Else
                BtnPOHEta.Enabled = True
                BtnPOHClose.Enabled = True

                PanPOHClose.Visible = False
            End If
            BtnPOHEta.Visible = BtnPOHEta.Enabled
            BtnPOHClose.Visible = BtnPOHClose.Enabled

            If RdbPOD.SelectedValue = "PO" Then
                psFillGrvPOD(vnPOHOID, vnSQLConn)
            ElseIf RdbPOD.SelectedValue = "PL" Then
                psFillGrvPLH(HdfPOHOID.Value, vnSQLConn)
            ElseIf RdbPOD.SelectedValue = "GR" Then
                psFillGrvRcvPOH(HdfPOHOID.Value, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanPOD.Visible = True
        End If
    End Sub

    Private Sub GrvPOH_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvPOH.PageIndexChanging
        GrvPOH.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOError.Text = pbMsgError
            LblMsgPOError.Visible = True
            Exit Sub
        End If

        psFillGrvPOH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        PanPOD.Visible = False
    End Sub

    Protected Sub BtnPOHEta_Click(sender As Object, e As EventArgs) Handles BtnPOHEta.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgPOError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            Exit Sub
        End If
        If Len(LblMsgPOHETADate.Text) > 6 Then
            TxtPOHEta.Text = Mid(LblMsgPOHETADate.Text, 7)
        Else
            TxtPOHEta.Text = ""
        End If
        TxtPOHEta.Focus()
        psShowPOHEta(True)
    End Sub

    Protected Sub BtnPOHClose_Click(sender As Object, e As EventArgs) Handles BtnPOHClose.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Close_Trans) = False Then
            LblMsgPOError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            Exit Sub
        End If
        psShowPOHClo(True)
    End Sub

    Protected Sub BtnPOHEtaYes_Click(sender As Object, e As EventArgs) Handles BtnPOHEtaYes.Click
        Dim vnUserOID As Integer = Session("UserOID")

        LblMsgPOHEta.Text = ""
        If Not IsDate(TxtPOHEta.Text) Then
            LblMsgPOHEta.Text = "Isi ETA"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOHEta.Text = pbMsgError
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            Dim vnQuery As String
            vnSQLTrans = vnSQLConn.BeginTransaction("upd")
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPOHeader_TR Set ETA_DATE='" & TxtPOHEta.Text & "',ETADatetime=Getdate(),ETAUserOID=" & vnUserOID & " Where OID=" & HdfPOHOID.Value
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing

            GrvPOH.Rows(HdfPOHRowIdx.Value).Cells(ensColPOH.vETA_DATE).Text = TxtPOHEta.Text
            LblMsgPOHETADate.Text = "ETA = " & TxtPOHEta.Text

            psShowPOHEta(False)
        Catch ex As Exception
            LblMsgPOHEta.Text = ex.Message

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If
        End Try

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub BtnPOHEtaNo_Click(sender As Object, e As EventArgs) Handles BtnPOHEtaNo.Click
        psShowPOHEta(False)
    End Sub

    Protected Sub BtnPOHCloNo_Click(sender As Object, e As EventArgs) Handles BtnPOHCloNo.Click
        psShowPOHClo(False)
    End Sub

    Private Sub BtnPOHCloYes_Click(sender As Object, e As EventArgs) Handles BtnPOHCloYes.Click
        Dim vnUserOID As Integer = Session("UserOID")

        LblMsgPOHClo.Text = ""
        If Trim(TxtPOHCloNote.Text) = "" Then
            LblMsgPOHClo.Text = "Isi Close Note"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOHClo.Text = pbMsgError
            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean

        Try
            Dim vnQuery As String
            vnSQLTrans = vnSQLConn.BeginTransaction("upd")
            vnBeginTrans = True

            vnQuery = "Update Sys_SsoPOHeader_TR Set POCloseNote='" & fbuFormatString(Trim(TxtPOHCloNote.Text)) & "',TransStatus=" & enuTCSPPO.Closed & ",ClosedDatetime=Getdate(),ClosedUserOID=" & vnUserOID & " Where OID=" & HdfPOHOID.Value
            pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

            vnSQLTrans.Commit()
            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing

            HdfPOHStatus.Value = enuTCSPPO.Closed
            GrvPOH.Rows(HdfPOHRowIdx.Value).Cells(ensColPOH.TransStatus).Text = enuTCSPPO.Closed

            GrvPOH.Rows(HdfPOHRowIdx.Value).Cells(ensColPOH.TransStatusDescr).Text = "Closed"
            LblMsgPOHStatus.Text = "Status = Closed"

            psShowPOHClo(False)
        Catch ex As Exception
            LblMsgPOHClo.Text = ex.Message

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If
        End Try

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvPLH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvPLH.SelectedIndexChanged

    End Sub

    Protected Sub RdbPOD_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RdbPOD.SelectedIndexChanged
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOHClo.Text = pbMsgError
            Exit Sub
        End If

        If RdbPOD.SelectedValue = "PO" Then
            GrvPOD.Visible = True
            PanPL.Visible = False
            PanGR.Visible = False
            psFillGrvPOD(HdfPOHOID.Value, vnSQLConn)

        ElseIf RdbPOD.SelectedValue = "PL" Then
            GrvPOD.Visible = False
            PanPL.Visible = True
            PanGR.Visible = False
            psFillGrvPLH(HdfPOHOID.Value, vnSQLConn)

        ElseIf RdbPOD.SelectedValue = "GR" Then
            GrvPOD.Visible = False
            PanPL.Visible = False
            PanGR.Visible = True
            psFillGrvRcvPOH(HdfPOHOID.Value, vnSQLConn)

        End If

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub GrvPLH_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvPLH.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvPLH.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvPLH.Rows(vnIdx)

        If e.CommandName = "PLNo" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgPOHClo.Text = pbMsgError
                Exit Sub
            End If

            psFillGrvPLD(GrvPLH.Rows(vnIdx).Cells(ensColPLH.OID).Text, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub GrvRcvPOH_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvRcvPOH.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnRowIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnRowIdx >= GrvRcvPOH.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvRcvPOH.Rows(vnRowIdx)

        If e.CommandName = "RcvPONo" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgPOHClo.Text = pbMsgError
                Exit Sub
            End If

            vnGRow = GrvRcvPOH.Rows(vnRowIdx)
            psFillGrvRcvPOSumm(HdfPOHOID.Value, vnGRow.Cells(ensColRcvPOH.OID).Text, vnGRow.Cells(ensColRcvPOH.TransStatus).Text, vnGRow.Cells(ensColRcvPOH.RcvPORefTypeOID).Text, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub GrvPOSap_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvPOSap.PageIndexChanging
        GrvPOSap.PageIndex = e.NewPageIndex
        psFillGrvPOSap()
    End Sub

    Protected Sub BtnPOSapFind_Click(sender As Object, e As EventArgs) Handles BtnPOSapFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgFindError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgFindError.Visible = True
            Exit Sub
        End If
        psFillGrvPOSap()
    End Sub

    Private Sub BtnPOSapPO_Click(sender As Object, e As EventArgs) Handles BtnPOSapPO.Click
        PanPOData.Visible = True
        PanPOUpload.Visible = False
        PanPOSAP.Visible = False
    End Sub

    Protected Sub BtnPOSap_Click(sender As Object, e As EventArgs) Handles BtnPOSap.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Upload_Xls) = False Then
            LblPOSapError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblPOSapError.Visible = True
            Exit Sub
        End If

        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        psClearMessage()

        If DstPOSapCompany.SelectedValue = "" Then
            LblPOSapCompany.Text = "Pilih Company"
            Exit Sub
        End If

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgXlsProsesError.Text = pbMsgError
            LblMsgXlsProsesError.Visible = True

            Exit Sub
        End If

        Dim vnSQLTrans As SqlTransaction = Nothing
        Dim vnBeginTrans As Boolean
        Dim vnTbTempCreated As Boolean
        Try
            Dim vnHOID As String
            Dim vnQuery As String
            vnQuery = "Select max(OID) From Sys_SsoPOFileXls_TR with(nolock)"
            vnHOID = fbuGetDataNumSQL(vnQuery, vnSQLConn) + 1

            pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "BtnPOSap_Click_SAPApi_GetPO", vnHOID, vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
            vsTextStream.WriteLine("")

            Dim vnDataFileNameOnly As String = ""
            Dim vnDataFileName As String = ""

            pbuCreateDataFile(vsFso_Data, vsTextStream_Data, Session("UserNip"), csModuleName, "BtnPOSap_Click_SAPApi_GetPO", vnHOID, stuFolderName.SAPApiFolder, vsLogFileName, vnDataFileNameOnly, vnDataFileName)

            vsTextStream.WriteLine("vnDataFileNameOnly = " & vnDataFileNameOnly)
            vsTextStream.WriteLine("vnDataFileName = " & vnDataFileName)

            psCreateTable_Sys_SsoPO_Temp(vnSQLConn)
            vnTbTempCreated = True

            vnSQLTrans = vnSQLConn.BeginTransaction("xls")
            vnBeginTrans = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------")
            vsTextStream.WriteLine("PO SAP...Start")

            Dim vnGetDate As String = fbuGetDateNowSQLTrans(vnSQLConn, vnSQLTrans)
            Dim vnCompany As String = DstPOSapCompany.SelectedValue
            vsTextStream.WriteLine("Company Code = " & vnCompany)

            vnQuery = "Insert into Sys_SsoPOFileXls_TR(OID,CompanyCode,XlsFileName,UploadSourceOID,UploadStartDatetime,UploadUserOID,StatusSuccess,StatusMessage)"
            vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vnCompany & "','" & vnDataFileNameOnly & ".txt'," & enuUploadSource.SAP_Api & ",'" & vnGetDate & "'," & vnUserOID & ",0,'Upload Start')"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)

            Dim vnMessage As String = ""
            Dim vnSuccess_GetSAP As Boolean
            If fbuSAPApi_GetPurchaseOrder(vsTextStream, vsTextStream_Data, vnUserOID, vnHOID, DstPOSapCompany.SelectedValue, vnSQLConn, vnSQLTrans, vnMessage) Then
                If fsImportDataPO_Step_2(vnGetDate, vnSQLConn, vnSQLTrans) = True Then
                    vnQuery = "Update Sys_SsoPOFileXls_TR Set StatusSuccess=1,StatusMessage='Success',UploadEndDatetime=getdate() Where OID=" & vnHOID
                    vsTextStream.WriteLine("")
                    vsTextStream.WriteLine("vnQuery")
                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                    vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                    vsTextStream.WriteLine("fbuSAPApi_GetPurchaseOrder...Success")
                    vsTextStream.WriteLine("=======================")
                    vsTextStream.WriteLine("")

                    vnBeginTrans = False
                    vnSQLTrans.Commit()

                    vnSuccess_GetSAP = True
                Else
                    vnSuccess_GetSAP = False
                End If
            Else
                vnSuccess_GetSAP = False
            End If

            If vnSuccess_GetSAP = False Then
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("fbuSAPApi_GetPurchaseOrder...Gagal")
                vsTextStream.WriteLine("=======================")
                vsTextStream.WriteLine("")

                vnBeginTrans = False
                vnSQLTrans.Rollback()

                vnSQLTrans = vnSQLConn.BeginTransaction("xls")
                vnBeginTrans = True

                vnQuery = "Insert into Sys_SsoPOFileXls_TR(OID,CompanyCode,UploadSourceOID,UploadStartDatetime,UploadUserOID,StatusSuccess,StatusMessage)"
                vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vnCompany & "'," & enuUploadSource.SAP_Api & ",getdate()," & vnUserOID & ",0,'" & fbuFormatString(vnMessage) & "')"
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
                vsTextStream.WriteLine("vnQuery")
                vsTextStream.WriteLine(vnQuery)
                pbuExecuteSQLTrans(vnQuery, cbuActionNew, vnSQLConn, vnSQLTrans)
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))

                vnBeginTrans = False
                vnSQLTrans.Commit()
            End If

            vnSQLTrans.Dispose()
            vnSQLTrans = Nothing
            vnBeginTrans = False

            psDropTable_Sys_SsoPO_Temp(vnSQLConn)
            vnTbTempCreated = False

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("------------------------EOF------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            vsTextStream_Data.WriteLine("")
            vsTextStream_Data.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream_Data.WriteLine("------------------------EOF------------------------")
            vsTextStream_Data.Close()
            vsTextStream_Data = Nothing
            vsFso_Data = Nothing

            psFillGrvPOSap()

        Catch ex As Exception
            LblPOSapError.Text = ex.Message
            LblPOSapError.Visible = True

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("PROCESS TERMINATED...ERROR :")
            vsTextStream.WriteLine(ex.Message)
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("-----------------------ERROR-----------------------")
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("------------------------EOF------------------------")

            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            If vnBeginTrans Then
                vnSQLTrans.Rollback()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
            End If
            If vnTbTempCreated Then
                psDropTable_Sys_SsoPO_Temp(vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End Try
    End Sub

    Protected Sub BtnPOGetSAP_Click(sender As Object, e As EventArgs) Handles BtnPOGetSAP.Click
        PanPOData.Visible = False
        PanPOUpload.Visible = False
        PanPOSAP.Visible = True
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub
End Class