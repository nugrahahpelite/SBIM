Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoSalesOrder
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoSalesOrder"

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
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Enum ensColSOH
        CompanyCode = 0
        SalesOrderNo = 1
        vSalesOrderDate = 2
        vSUB = 3
        NAMA_CUSTOMER = 4
        SalesOrderHOID = 5
    End Enum
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Sso"

        If Session("UserName") = "" Then
            Response.Redirect("~/Default.aspx")
        End If
        If Not IsPostBack Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoCustomerSO, vnSQLConn)

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
                pbuFillDstCompany(DstListCompany, True, vnSQLConn)
                pbuFillDstCompany(DstPOCompany, True, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstListCompany, True, vnSQLConn)
                pbuFillDstCompanyByUser(Session("UserOID"), DstPOCompany, True, vnSQLConn)
            End If

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
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
        vnQuery += vbCrLf & " From Sys_SsoSalesOrderFileXls_TR PM"
        vnQuery += vbCrLf & "      inner Join Sys_SsoUser_MA MU on MU.OID=PM.UploadUserOID"
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "Where PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If
        vnQuery += vbCrLf & " Order by PM.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

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
        vsXlsFileName = vsXlsFolder & stuTransCode.SsoCustomerSO & "_" & Format(Date.Now, "yyyyMMdd_HHmmss ") & FupXls.FileName
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

        Const cnColOJL = 0
        Const cnColJURNAL = 1
        Const cnColDATE_ = 2
        Const cnColSUB = 3
        Const cnColBRG = 4
        Const cnColGDGOJL = 5
        Const cnColQTY = 6
        Const cnColLNS = 7
        Const cnColVal = 8
        Const cnColJOB = 9
        Const cnColUnit = 10

        Const cnColNAMA_CUSTOMER = 11
        Const cnColNAMA_BARANG = 12
        Const cnColJOBNAME = 13
        Const cnColUNIQUEIDOJD = 14
        Const cnColALAMAT = 15
        Const cnColNAMA_KOTA = 16
        Const cnColPOCUS = 17
        'Const cnColDELDATE = 18

        Dim vnColOJL As String
        Dim vnColJURNAL As String
        Dim vnColDATE_ As String
        Dim vnColSUB As String
        Dim vnColBRG As String
        Dim vnColGDGOJL As String
        Dim vnColQTY As String
        Dim vnColLNS As String
        Dim vnColVal As String
        Dim vnColJOB As String
        Dim vnColUnit As String

        Dim vnColNAMA_CUSTOMER As String
        Dim vnColNAMA_BARANG As String
        Dim vnColJOBNAME As String
        Dim vnColUNIQUEIDOJD As String
        Dim vnColALAMAT As String
        Dim vnColNAMA_KOTA As String
        Dim vnColPOCUS As String
        'Dim vnColDELDATE As String

        Dim vnQuery As String

        vnQuery = "Delete Sys_SsoSalesOrder_Temp"
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        vnQuery = "Select max(OID) From Sys_SsoSalesOrderFileXls_TR with(nolock)"

        Dim vnGetDate As String = fbuGetDateNowSQLTrans(vriSQLConn, vriSQLTrans)
        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) + 1

        vnQuery = "Insert into Sys_SsoSalesOrderFileXls_TR"
        vnQuery += vbCrLf & "(OID,CompanyCode,XlsFileName,UploadStartDatetime,UploadUserOID)"
        vnQuery += vbCrLf & "Values"
        vnQuery += vbCrLf & "(" & vnHOID & ",'" & vnCompanyCode & "','" & vsXlsFileName & "','" & vnGetDate & "'," & Session("UserOID") & ")"
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

            vnColOJL = fbuValStr(vnXReader.Item(cnColOJL))
            vsTextStream.WriteLine(vbCrLf & "vnColOBL " & vnColOJL)

            vnColJURNAL = fbuValStr(vnXReader.Item(cnColJURNAL))
            vsTextStream.WriteLine(vbCrLf & "vnColJURNAL " & vnColJURNAL)

            'vnColDATE_ = fbuFormatDateDMY_To_YMD(fbuValStr(vnXReader.Item(cnColDATE_)))
            vnColDATE_ = DateTime.Parse(vnXReader.Item(cnColDATE_)).ToString("dd MMM yyyy")
            vsTextStream.WriteLine(vbCrLf & "vnColDATE_ " & vnColDATE_)

            vnColSUB = fbuValStr(vnXReader.Item(cnColSUB))
            vsTextStream.WriteLine(vbCrLf & "vnColSUB " & vnColSUB)

            vnColBRG = fbuValStr(vnXReader.Item(cnColBRG))
            vsTextStream.WriteLine(vbCrLf & "vnColBRG " & vnColBRG)

            vnColGDGOJL = fbuFormatString(fbuValStr(vnXReader.Item(cnColGDGOJL)))
            vsTextStream.WriteLine(vbCrLf & "vnColGDGOBL " & vnColGDGOJL)

            vnColQTY = fbuValNum(vnXReader.Item(cnColQTY))
            vsTextStream.WriteLine(vbCrLf & "vnColQTY " & vnColQTY)

            vnColLNS = fbuValNum(vnXReader.Item(cnColLNS))
            vsTextStream.WriteLine(vbCrLf & "vnColLNS " & vnColLNS)

            vnColVal = fbuValNum(vnXReader.Item(cnColVal))
            vsTextStream.WriteLine(vbCrLf & "vnColVal " & vnColVal)

            vnColJOB = fbuFormatString(fbuValStr(vnXReader.Item(cnColJOB)))
            vsTextStream.WriteLine(vbCrLf & "vnColJOB " & vnColJOB)

            vnColUnit = fbuFormatString(fbuValStr(vnXReader.Item(cnColUnit)))
            vsTextStream.WriteLine(vbCrLf & "vnColUnit " & vnColUnit)

            vnColUNIQUEIDOJD = fbuFormatString(fbuValStr(vnXReader.Item(cnColUNIQUEIDOJD)))
            vsTextStream.WriteLine(vbCrLf & "vnColUNIQUEIDOJD " & vnColUNIQUEIDOJD)

            vnColNAMA_CUSTOMER = fbuFormatString(fbuValStr(vnXReader.Item(cnColNAMA_CUSTOMER)))
            vsTextStream.WriteLine(vbCrLf & "vnColNAMA_CUSTOMER " & vnColNAMA_CUSTOMER)

            vnColNAMA_BARANG = fbuFormatString(fbuValStr(vnXReader.Item(cnColNAMA_BARANG)))
            vsTextStream.WriteLine(vbCrLf & "vnColNAMA_BARANG " & vnColNAMA_BARANG)

            vnColJOBNAME = fbuFormatString(fbuValStr(vnXReader.Item(cnColJOBNAME)))
            vsTextStream.WriteLine(vbCrLf & "vnColJOBNAME " & vnColJOBNAME)

            vnColALAMAT = fbuFormatString(fbuValStr(vnXReader.Item(cnColALAMAT)))
            vsTextStream.WriteLine(vbCrLf & "vnColALAMAT " & vnColALAMAT)

            vnColNAMA_KOTA = fbuValStr(vnXReader.Item(cnColNAMA_KOTA))
            vsTextStream.WriteLine(vbCrLf & "vnColNAMA_KOTA " & vnColNAMA_KOTA)

            vnColPOCUS = fbuFormatString(fbuValStr(vnXReader.Item(cnColPOCUS)))
            vsTextStream.WriteLine(vbCrLf & "vnColPOCUS " & vnColPOCUS)

            'vnColDELDATE = fbuValStr(vnXReader.Item(cnColDELDATE))
            'vsTextStream.WriteLine(vbCrLf & "vnColDELDATE " & vnColDELDATE)

            vnQuery = "Insert into Sys_SsoSalesOrder_Temp"
            vnQuery += vbCrLf & "(CompanyCode,SalesOrderNo,JURNAL,SalesOrderDate,SUB,BRG,GDGOJL,"
            vnQuery += vbCrLf & "QTY,LNS,VAL,JOB,UNIT,UNIQUEIDOJD,NAMA_CUSTOMER,"
            vnQuery += vbCrLf & "NAMA_BARANG,JOBNAME,ALAMAT,NAMA_KOTA,POCUS,DELDATE,"
            vnQuery += vbCrLf & "SalesOrderFileXlsOID,UploadDatetime"
            vnQuery += vbCrLf & ")"

            vnQuery += vbCrLf & "Select '" & vnCompanyCode & "'CompanyCode,'" & vnColOJL & "'SalesOrderNo,'" & vnColJURNAL & "'JURNAL,'" & vnColDATE_ & "'SalesOrderDate,'" & vnColSUB & "'SUB,'" & vnColBRG & "'BRG,'" & vnColGDGOJL & "'GDGOJL,"
            vnQuery += vbCrLf & "'" & vnColQTY & "'QTY,'" & vnColLNS & "'LNS,'" & vnColVal & "'VAL,'" & vnColJOB & "'JOB,'" & vnColUnit & "'UNIT,'" & vnColUNIQUEIDOJD & "'UNIQUEIDOJD,'" & vnColNAMA_CUSTOMER & "'NAMA_CUSTOMER,"
            vnQuery += vbCrLf & "'" & vnColNAMA_BARANG & "'NAMA_BARANG,'" & vnColJOBNAME & "'JOBNAME,'" & vnColALAMAT & "'ALAMAT,'" & vnColNAMA_KOTA & "'NAMA_KOTA,'" & vnColPOCUS & "'POCUS,Null DELDATE,"
            vnQuery += vbCrLf & vnHOID & ",'" & vnGetDate & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        vnQuery = "Delete Sys_SsoSalesOrder_Temp Where isnull(JURNAL,'')<>''"
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        '<---12 Oct 2023 dimatiin Agus
        'vnQuery = "Delete POD"
        'vnQuery += vbCrLf & "       From Sys_SsoSalesOrderDetail_TR POD"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrderHeader_TR POH ON POH.OID=POD.SalesOrderHOID"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrder_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.SalesOrderNo=POH.SalesOrderNo AND ABT.BRG=POD.BRG"
        'vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCCSSO.Baru
        'vsTextStream.WriteLine("vnQuery")
        'vsTextStream.WriteLine(vnQuery)
        'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        'vnQuery = "Delete POD"
        'vnQuery += vbCrLf & "       From Sys_SsoSalesOrder POD"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrderHeader_TR POH ON POH.OID=POD.SalesOrderHOID"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrder_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.SalesOrderNo=POH.SalesOrderNo AND ABT.BRG=POD.BRG"
        'vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCCSSO.Baru
        'vsTextStream.WriteLine("vnQuery")
        'vsTextStream.WriteLine(vnQuery)
        'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        'vnQuery = "Delete POD"
        'vnQuery += vbCrLf & "       From Sys_SsoSalesOrderStatus_TR POD"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrderHeader_TR POH ON POH.OID=POD.SalesOrderHOID"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrder_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.SalesOrderNo=POH.SalesOrderNo"
        'vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCCSSO.Baru
        'vsTextStream.WriteLine("vnQuery")
        'vsTextStream.WriteLine(vnQuery)
        'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        'vnQuery = "Delete POH"
        'vnQuery += vbCrLf & "       From Sys_SsoSalesOrderHeader_TR POH"
        'vnQuery += vbCrLf & "	         Inner Join Sys_SsoSalesOrder_Temp ABT ON ABT.CompanyCode=POH.CompanyCode AND ABT.SalesOrderNo=POH.SalesOrderNo"
        'vnQuery += vbCrLf & "	   Where POH.TransStatus=" & enuTCCSSO.Baru
        'vsTextStream.WriteLine("vnQuery")
        'vsTextStream.WriteLine(vnQuery)
        'pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        '<<==12 Oct 2023 dimatiin Agus

        vnQuery = "Insert into Sys_SsoSalesOrder"
        vnQuery += vbCrLf & "Select *,0 SalesOrderHOID From Sys_SsoSalesOrder_Temp ABT with(nolock) WHERE NOT ABT.SalesOrderNo+ABT.BRG IN"
        vnQuery += vbCrLf & "	   (Select AB.SalesOrderNo+AB.BRG FROM Sys_SsoSalesOrder AB)"
        vsTextStream.WriteLine("vnQuery")
        vsTextStream.WriteLine(vnQuery)
        pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

        Dim vnPOHOID As Integer
        vnQuery = "Select isnull(max(OID),0) From Sys_SsoSalesOrderHeader_TR with(nolock)"
        vnPOHOID = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        Dim vnDtbPOH As New DataTable
        vnQuery = "Select distinct CompanyCode,SalesOrderNo From Sys_SsoSalesOrder Where SalesOrderHOID=0"
        pbuFillDtbSQLTrans(vnDtbPOH, vnQuery, vriSQLConn, vriSQLTrans)
        For vn = 0 To vnDtbPOH.Rows.Count - 1
            vnPOHOID = vnPOHOID + 1
            vnQuery = "Insert into Sys_SsoSalesOrderHeader_TR(OID,CompanyCode,SalesOrderNo,SalesOrderDate,SUB,NAMA_CUSTOMER,ALAMAT,NAMA_KOTA,TransCode,TransStatus)"
            vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",CompanyCode,SalesOrderNo,SalesOrderDate,SUB,NAMA_CUSTOMER,ALAMAT,NAMA_KOTA,'" & stuTransCode.SsoCustomerSO & "'TransCode," & enuTCCSSO.Baru & " TransStatus"
            vnQuery += vbCrLf & "From Sys_SsoSalesOrder Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and SalesOrderNo='" & vnDtbPOH.Rows(vn).Item("SalesOrderNo") & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Update Sys_SsoSalesOrder set SalesOrderHOID=" & vnPOHOID & "Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and SalesOrderNo='" & vnDtbPOH.Rows(vn).Item("SalesOrderNo") & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into Sys_SsoSalesOrderStatus_TR(SalesOrderHOID,TransCode,TransStatus,TransStatusDatetime)"
            vnQuery += vbCrLf & "Select distinct " & vnPOHOID & ",'" & stuTransCode.SsoCustomerSO & "'TransCode," & enuTCCSSO.Baru & " TransStatus,'" & vnGetDate & "'"
            vnQuery += vbCrLf & "From Sys_SsoSalesOrder Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and SalesOrderNo='" & vnDtbPOH.Rows(vn).Item("SalesOrderNo") & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into Sys_SsoSalesOrderDetail_TR(SalesOrderHOID,GDGOJL,BRG,NAMA_BARANG,QTY,LNS,VAL,JOB,UNIT,UNIQUEIDOJD,JOBNAME,POCUS,DELDATE)"
            vnQuery += vbCrLf & "Select " & vnPOHOID & ",GDGOJL,BRG,NAMA_BARANG,QTY,LNS,VAL,JOB,UNIT,UNIQUEIDOJD,JOBNAME,POCUS,DELDATE"
            vnQuery += vbCrLf & "From Sys_SsoSalesOrder Where CompanyCode='" & vnDtbPOH.Rows(vn).Item("CompanyCode") & "' and SalesOrderNo='" & vnDtbPOH.Rows(vn).Item("SalesOrderNo") & "'"
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        Next

        vsTextStream.WriteLine(vbCrLf & "")
        vsTextStream.WriteLine(vbCrLf & "Commit Transaction")

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData = True
    End Function

    Protected Sub BtnData_Click(sender As Object, e As EventArgs) Handles BtnData.Click
        PanSOData.Visible = True
        PanSOUpload.Visible = False
    End Sub

    Protected Sub BtnSOUpload_Click(sender As Object, e As EventArgs) Handles BtnSOUpload.Click
        PanSOData.Visible = False
        PanSOUpload.Visible = True
    End Sub

    Protected Sub BtnSOFind_Click(sender As Object, e As EventArgs) Handles BtnSOFind.Click
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.View_Data) = False Then
            LblMsgSOFindError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgSOFindError.Visible = True
            Exit Sub
        End If

        LblMsgSOFindError.Text = ""
        LblMsgSOFindError.Visible = False

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgSOFindError.Text = pbMsgError
            LblMsgSOFindError.Visible = True
            Exit Sub
        End If

        psFillGrvSOH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvSOH(vriSQLConn As SqlConnection)
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        Dim vnCriteria As String
        Dim vnSupplier As String = fbuFormatString(Trim(TxtSOCustomer.Text))

        If ChkSt_TRB_Full.Checked = False And ChkSt_TRB_Not.Checked = False And ChkSt_TRB_Sebagian.Checked = False Then
            ChkSt_TRB_Full.Checked = True
            ChkSt_TRB_Not.Checked = True
            ChkSt_TRB_Sebagian.Checked = True
        End If

        Dim vnCrStatus As String = ""
        If Not (ChkSt_TRB_Full.Checked = True And ChkSt_TRB_Not.Checked = True And ChkSt_TRB_Sebagian.Checked = True) Then
            If ChkSt_TRB_Full.Checked = True Then
                vnCrStatus += "sum(pod.QTY_TRB)>=sum(pod.QTY)"
            End If
            If ChkSt_TRB_Not.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "sum(pod.QTY_TRB)=0"
            End If
            If ChkSt_TRB_Sebagian.Checked = True Then
                vnCrStatus += IIf(vnCrStatus = "", "", " or ") & "(sum(pod.QTY_TRB)>0 and sum(pod.QTY_TRB)<sum(pod.QTY))"
            End If
            vnCrStatus = " and OID in (Select pod.SalesOrderHOID From Sys_SsoSalesOrderDetail_TR pod with(nolock) group by pod.SalesOrderHOID having " & vnCrStatus & ")"
        End If

        vnCriteria = "      Where 1=1"
        If DstPOCompany.SelectedValue <> "" Then
            vnCriteria += vbCrLf & "            and soh.CompanyCode='" & DstPOCompany.SelectedValue & "'"
        End If
        If Trim(TxtSOCustomer.Text) <> "" Then
            vnCriteria += vbCrLf & "            and (soh.SUB like '%" & vnSupplier & "%' or soh.NAMA_CUSTOMER like '%" & vnSupplier & "%')"
        End If
        If Trim(TxtSONo.Text) <> "" Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderNo like '%" & fbuFormatString(Trim(TxtSONo.Text)) & "%'"
        End If
        If IsDate(TxtSOStart.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate >= '" & TxtSOStart.Text & "'"
        End If
        If IsDate(TxtSOEnd.Text) Then
            vnCriteria += vbCrLf & "            and soh.SalesOrderDate <= '" & TxtSOEnd.Text & "'"
        End If

        vnQuery = "Select soh.CompanyCode,soh.SalesOrderNo,convert(varchar(11),soh.SalesOrderDate,106)vSalesOrderDate,soh.SUB vSUB,soh.NAMA_CUSTOMER,"
        vnQuery += vbCrLf & "            OID SalesOrderHOID,case when abs(SOVoid)=1 then 'Y' else 'N' end vSOVoid,"
        vnQuery += vbCrLf & "            case when abs(SOVoid)=1 then soh.SOVoidNo +'<br />'+convert(varchar(11),SOVoidDatetime,106)+' '+convert(varchar(5),SOVoidDatetime,108)+'<br />'+SOVoidNote else '' end vSOVoid_Info"
        vnQuery += vbCrLf & "       From Sys_SsoSalesOrderHeader_TR soh with(nolock)"

        If vnUserCompanyCode <> "" And DstPOCompany.SelectedValue = "" Then
            vnQuery += vbCrLf & "            inner join Sys_SsoUserCompany_MA mu with(nolock) on mu.CompanyCode=soh.CompanyCode and mu.UserOID=" & Session("UserOID")
        End If

        vnQuery += vbCrLf & vnCriteria
        vnQuery += vbCrLf & vnCrStatus
        vnQuery += vbCrLf & "Order by soh.CompanyCode,soh.SalesOrderNo"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSOH.DataSource = vnDtb
        GrvSOH.DataBind()

        PanSOD.Visible = False
    End Sub
    Private Sub psFillGrvSOD(vriSalesOrderHOID As Integer, vriSQLConn As SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select pod.GDGOJL,pod.BRG,pod.NAMA_BARANG,pod.QTY,pod.QTY_TRB,"
        vnQuery += vbCrLf & "            pod.OID vSalesOrderDOID,pod.SalesOrderHOID"
        vnQuery += vbCrLf & "       From Sys_SsoSalesOrderDetail_TR pod with(nolock)"

        vnQuery += vbCrLf & "Where SalesOrderHOID=" & vriSalesOrderHOID
        vnQuery += vbCrLf & "Order by pod.NAMA_BARANG"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        GrvSOD.DataSource = vnDtb
        GrvSOD.DataBind()
    End Sub
    Private Sub GrvSOD_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvSOD.PageIndexChanging
        GrvSOD.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOError.Text = pbMsgError
            LblMsgPOError.Visible = True
            Exit Sub
        End If

        psFillGrvSOD(LblMsgSOHOID.Text, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Protected Sub GrvSOH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvSOH.SelectedIndexChanged

    End Sub

    Private Sub GrvSOH_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvSOH.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvSOH.Rows(vnIdx)

        If e.CommandName = "SalesOrderNo" Then
            Dim vnSalesOrderHOID As String = vnGRow.Cells(ensColSOH.SalesOrderHOID).Text
            Dim vnSONo As String = DirectCast(vnGRow.Cells(ensColSOH.SalesOrderNo).Controls(0), LinkButton).Text
            LblMsgSOHOID.Text = vnSalesOrderHOID
            LblMsgSODNo.Text = vnSONo

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgPOError.Text = pbMsgError
                LblMsgPOError.Visible = True
                Exit Sub
            End If

            psFillGrvSOD(vnSalesOrderHOID, vnSQLConn)

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            PanSOD.Visible = True
        End If
    End Sub

    Private Sub GrvSOH_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GrvSOH.PageIndexChanging
        GrvSOH.PageIndex = e.NewPageIndex

        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgPOError.Text = pbMsgError
            LblMsgPOError.Visible = True
            Exit Sub
        End If

        psFillGrvSOH(vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing

        PanSOD.Visible = False
    End Sub
End Class