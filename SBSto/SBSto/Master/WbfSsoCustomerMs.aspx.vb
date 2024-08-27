Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class WbfSsoCustomerMs
    Inherits System.Web.UI.Page

    Const csModuleName = "WbfSsoCustomerMs"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("CurrentFolder") = "Master"

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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsCustomer, vnSQLConn)

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

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.CompanyCode,PM.CUSTSUB,PM.CUSTNAME,PM.CUSTPERSON"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstCustomer_MA PM"
        vnQuery += vbCrLf & "Where 1=1"
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If
        If Trim(TxtKriteria.Text) <> "" Then
            vnQuery += vbCrLf & "      and (PM.CUSTSUB like '%" & Trim(TxtKriteria.Text) & "%' OR PM.CUSTNAME like '%" & Trim(TxtKriteria.Text) & "%')"
        End If

        vnQuery += vbCrLf & " Order by PM.CUSTNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psClearMessage()
        LblMsgCompany.Visible = False
        LblXlsProses.Visible = False
        LblXlsWorksheet.Visible = False
        LblMsgXlsProsesError.Visible = False
        LblMsgError.Visible = False
    End Sub
    Protected Sub BtnXlsUpload_Click(sender As Object, e As EventArgs) Handles BtnXlsUpload.Click
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        psClearMessage()

        If DstCompany.SelectedValue = "" Then
            LblMsgCompany.Text = "Pilih Company"
            LblMsgCompany.Visible = True
            Exit Sub
        End If
        If TxtXlsWorksheet.Text = "" Then
            LblXlsWorksheet.Text = "Isi Nama Worksheet"
            LblXlsWorksheet.Visible = True
            Exit Sub
        End If
        If fbuValAccess(ViewState("UGAccess"), stuTrAccessCode.Create_EditDel) = False Then
            LblMsgError.Text = "Akses Error....Anda Tidak Memiliki Akses"
            LblMsgError.Visible = True
            Exit Sub
        End If

        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "BtnXlsUpload_Click", "0", vsLogFileName, vsLogFileNameError, vsLogFileNameErrorSend)
        vsTextStream.WriteLine("Open SQL Connection....Start")

        vsXlsFolder = Server.MapPath("~") & "\XlsFolder\"
        vsXlsFileName = vsXlsFolder & "BRG_" & Format(Date.Now, "yyyyMMdd_HHmmss ") & FupXls.FileName
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

        Const cnCUSTSUB = 0
        Const cnCUSTNAME = 1
        Const cnCUSTPERSON = 2
        Const cnCUSTPHONE = 3
        Const cnCUSTFAX = 4
        Const cnCUSTADDRESS = 5
        Const cnCUSTPOST = 6
        Const cnCUSTNPWP = 7
        Const cnCUSTREM = 8
        Const cnCUSTMAX_ = 9
        Const cnCUSTPRK = 10
        Const cnCUSTPRKUM = 11
        Const cnCUSTREV_PRK = 12
        Const cnCUSTRJL_PRK = 13
        Const cnCUSTDISC_PRK = 14
        Const cnCUSTCHUSR = 15
        Const cnCUSTCHTIME = 16
        Const cnCUSTGROUP_ = 17
        Const cnCUSTKOTA = 18
        Const cnCUSTINTERN_GROUP = 19
        Const cnCUSTPSR = 20
        Const cnCUSTINACTIVE_ = 21
        Const cnCUSTOWNER = 22
        Const cnCUSTADDRESS_OWNER = 23
        Const cnCUSTKOTA_OWNER = 24
        Const cnCUSTEMAIL = 25
        Const cnCUSTJABATAN = 26
        Const cnCUSTNPKP = 27
        Const cnCUSTGRADE = 28
        Const cnCUSTKODERETUR = 29
        Const cnCUSTSLSM = 30
        Const cnCUSTKOLEKTOR = 31
        Const cnCUSTBILLTO = 32
        Const cnCUSTSENDTO = 33
        Const cnCUSTISFPJ = 34
        Const cnCUSTREGION = 35
        Const cnCUSTHR_MSK_FAKTUR = 36
        Const cnCUSTHR_PENAGIHAN = 37
        Const cnCUSTHR_PEMBAYARAN = 38
        Const cnCUSTTERMS = 39
        Const cnCUSTDISC = 40
        Const cnCUSTTAX = 41
        Const cnCUSTCOA = 42
        Const cnCUSTJNSDEFAULT = 43
        Const cnCUSTAGAMA = 44
        Const cnCUSTTGL_LAHIR = 45
        Const cnCUSTKDTR_FP = 46
        Const cnCUSTDESA = 47
        Const cnCUSTSLSMNAME = 48
        Const cnCUSTKOTANAMA = 49

        Dim vnCUSTSUB As String
        Dim vnCUSTNAME As String
        Dim vnCUSTPERSON As String
        Dim vnCUSTPHONE As String
        Dim vnCUSTFAX As String
        Dim vnCUSTADDRESS As String
        Dim vnCUSTPOST As String
        Dim vnCUSTKOTA As String
        Dim vnCUSTKOTANAMA As String

        Dim vnQuery As String
        vnQuery = "Select isnull(max(convert(int,OID)),0)+1 From " & fbuGetDBMaster() & "Sys_MstCustomerUpload_TR"

        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomerUpload_TR("
        vnQuery += vbCrLf & "OID,CustXlsFileName,CustWorkSheetName,UploadDatetime,UploadUserOID)"
        vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vsXlsFileName & "','" & vsSheetName & "',getdate()," & Session("UserOID") & ")"
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
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnCUSTSUB = Trim(fbuValStr(vnXReader.Item(cnCUSTSUB)))

            If vnCUSTSUB <> "" Then
                vnQuery = "Select CUSTSUB From " & fbuGetDBMaster() & "Sys_MstCustomer_MA Where CompanyCode='" & vnCompanyCode & "' and CUSTSUB='" & vnCUSTSUB & "'"
                vsTextStream.WriteLine(vnQuery)
                If fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = "" Then
                    vsTextStream.WriteLine("CUSTSUB NOT EXIST")
                    vnCUSTNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTNAME))))
                    vnCUSTPERSON = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTPERSON)))
                    vnCUSTPHONE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTPHONE))))
                    vnCUSTFAX = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTFAX))))
                    vnCUSTADDRESS = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTADDRESS))))
                    vnCUSTPOST = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTPOST)))
                    vnCUSTKOTA = fbuValStr(vnXReader.Item(cnCUSTKOTA))
                    vnCUSTKOTANAMA = fbuValStr(vnXReader.Item(cnCUSTKOTANAMA))

                    vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomer_MA("
                    vnQuery += vbCrLf & "CompanyCode,CUSTSUB,CUSTNAME,CUSTPERSON,"
                    vnQuery += vbCrLf & "CUSTPHONE,CUSTFAX,CUSTADDRESS,CUSTPOST,"
                    vnQuery += vbCrLf & "CUSTKOTA,CUSTKOTANAMA,CustXlsFileOID)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnCUSTSUB & "' vnCUSTSUB,'" & vnCUSTNAME & "' vnCUSTNAME,'" & vnCUSTPERSON & "' vnCUSTPERSON,"
                    vnQuery += vbCrLf & "'" & vnCUSTPHONE & "' vnCUSTPHONE,'" & vnCUSTFAX & "' vnCUSTFAX,'" & vnCUSTADDRESS & "' vnCUSTADDRESS,'" & vnCUSTPOST & "' vnCUSTPOST,"
                    vnQuery += vbCrLf & "'" & vnCUSTKOTA & "' vnCUSTKOTA,'" & vnCUSTKOTANAMA & "' vnCUSTKOTANAMA,'" & vnHOID & "' vnHOID"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("CUSTSUB ALREADY EXIST")
                End If
            Else
                vsTextStream.WriteLine("CUSTSUB = EMPTY --> EXIT LOOP")
                Exit While
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData = True
    End Function

    Private Function fsXlsImportData_20221212_Orig_Kolom_Lengkap(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData_20221212_Orig_Kolom_Lengkap = False
        Dim vnCompanyCode As String = Session("UserCompanyCode")

        Const cnCUSTSUB = 0
        Const cnCUSTNAME = 1
        Const cnCUSTPERSON = 2
        Const cnCUSTPHONE = 3
        Const cnCUSTFAX = 4
        Const cnCUSTADDRESS = 5
        Const cnCUSTPOST = 6
        Const cnCUSTNPWP = 7
        Const cnCUSTREM = 8
        Const cnCUSTMAX_ = 9
        Const cnCUSTPRK = 10
        Const cnCUSTPRKUM = 11
        Const cnCUSTREV_PRK = 12
        Const cnCUSTRJL_PRK = 13
        Const cnCUSTDISC_PRK = 14
        Const cnCUSTCHUSR = 15
        Const cnCUSTCHTIME = 16
        Const cnCUSTGROUP_ = 17
        Const cnCUSTKOTA = 18
        Const cnCUSTINTERN_GROUP = 19
        Const cnCUSTPSR = 20
        Const cnCUSTINACTIVE_ = 21
        Const cnCUSTOWNER = 22
        Const cnCUSTADDRESS_OWNER = 23
        Const cnCUSTKOTA_OWNER = 24
        Const cnCUSTEMAIL = 25
        Const cnCUSTJABATAN = 26
        Const cnCUSTNPKP = 27
        Const cnCUSTGRADE = 28
        Const cnCUSTKODERETUR = 29
        Const cnCUSTSLSM = 30
        Const cnCUSTKOLEKTOR = 31
        Const cnCUSTBILLTO = 32
        Const cnCUSTSENDTO = 33
        Const cnCUSTISFPJ = 34
        Const cnCUSTREGION = 35
        Const cnCUSTHR_MSK_FAKTUR = 36
        Const cnCUSTHR_PENAGIHAN = 37
        Const cnCUSTHR_PEMBAYARAN = 38
        Const cnCUSTTERMS = 39
        Const cnCUSTDISC = 40
        Const cnCUSTTAX = 41
        Const cnCUSTCOA = 42
        Const cnCUSTJNSDEFAULT = 43
        Const cnCUSTAGAMA = 44
        Const cnCUSTTGL_LAHIR = 45
        Const cnCUSTKDTR_FP = 46
        Const cnCUSTDESA = 47
        Const cnCUSTSLSMNAME = 48
        Const cnCUSTKOTANAMA = 49

        Dim vnCUSTSUB As String
        Dim vnCUSTNAME As String
        Dim vnCUSTPERSON As String
        Dim vnCUSTPHONE As String
        Dim vnCUSTFAX As String
        Dim vnCUSTADDRESS As String
        Dim vnCUSTPOST As String
        Dim vnCUSTNPWP As String
        Dim vnCUSTREM As String
        Dim vnCUSTMAX_ As String
        Dim vnCUSTPRK As String
        Dim vnCUSTPRKUM As String
        Dim vnCUSTREV_PRK As String
        Dim vnCUSTRJL_PRK As String
        Dim vnCUSTDISC_PRK As String
        Dim vnCUSTCHUSR As String
        Dim vnCUSTCHTIME As String
        Dim vnCUSTGROUP_ As String
        Dim vnCUSTKOTA As String
        Dim vnCUSTINTERN_GROUP As String
        Dim vnCUSTPSR As String
        Dim vnCUSTINACTIVE_ As String
        Dim vnCUSTOWNER As String
        Dim vnCUSTADDRESS_OWNER As String
        Dim vnCUSTKOTA_OWNER As String
        Dim vnCUSTEMAIL As String
        Dim vnCUSTJABATAN As String
        Dim vnCUSTNPKP As String
        Dim vnCUSTGRADE As String
        Dim vnCUSTKODERETUR As String
        Dim vnCUSTSLSM As String
        Dim vnCUSTKOLEKTOR As String
        Dim vnCUSTBILLTO As String
        Dim vnCUSTSENDTO As String
        Dim vnCUSTISFPJ As String
        Dim vnCUSTREGION As String
        Dim vnCUSTHR_MSK_FAKTUR As String
        Dim vnCUSTHR_PENAGIHAN As String
        Dim vnCUSTHR_PEMBAYARAN As String
        Dim vnCUSTTERMS As String
        Dim vnCUSTDISC As String
        Dim vnCUSTTAX As String
        Dim vnCUSTCOA As String
        Dim vnCUSTJNSDEFAULT As String
        Dim vnCUSTAGAMA As String
        Dim vnCUSTTGL_LAHIR As String
        Dim vnCUSTKDTR_FP As String
        Dim vnCUSTDESA As String
        Dim vnCUSTSLSMNAME As String
        Dim vnCUSTKOTANAMA As String

        Dim vnQuery As String
        vnQuery = "Select isnull(max(OID),0)+1 From " & fbuGetDBMaster() & "Sys_MstCustomerUpload_TR"

        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomerUpload_TR("
        vnQuery += vbCrLf & "OID,CustXlsFileName,CustWorkSheetName,UploadDatetime,UploadUserOID)"
        vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vsXlsFileName & "','" & vsSheetName & "',getdate()," & Session("UserOID") & ")"
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
        vsTextStream.WriteLine("--------------------Loop...Start--------------------")
        While vnXReader.Read
            vsTextStream.WriteLine("")

            vnNo = vnNo + 1
            vsTextStream.WriteLine("vnNo " & vnNo)

            vnCUSTSUB = Trim(fbuValStr(vnXReader.Item(cnCUSTSUB)))

            If vnCUSTSUB <> "" Then
                vnQuery = "Select CUSTSUB From " & fbuGetDBMaster() & "Sys_MstCustomer_MA Where CompanyCode='" & vnCompanyCode & "' and CUSTSUB='" & vnCUSTSUB & "'"
                vsTextStream.WriteLine(vnQuery)
                If fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = "" Then
                    vsTextStream.WriteLine("CUSTSUB NOT EXIST")
                    vnCUSTNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTNAME))))
                    vnCUSTPERSON = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTPERSON)))
                    vnCUSTPHONE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTPHONE))))
                    vnCUSTFAX = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTFAX))))
                    vnCUSTADDRESS = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTADDRESS))))
                    vnCUSTPOST = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTPOST)))
                    vnCUSTNPWP = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTNPWP))))
                    vnCUSTREM = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTREM)))
                    vnCUSTMAX_ = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTMAX_))))
                    vnCUSTPRK = fbuFormatString(fbuValNullStr(vnXReader.Item(cnCUSTPRK)))
                    vnCUSTPRKUM = fbuValStr(vnXReader.Item(cnCUSTPRKUM))
                    vnCUSTREV_PRK = fbuValStr(vnXReader.Item(cnCUSTREV_PRK))
                    vnCUSTRJL_PRK = fbuValStr(vnXReader.Item(cnCUSTRJL_PRK))
                    vnCUSTDISC_PRK = fbuValStr(vnXReader.Item(cnCUSTDISC_PRK))
                    vnCUSTCHUSR = fbuValStr(vnXReader.Item(cnCUSTCHUSR))
                    vnCUSTCHTIME = fbuValStr(vnXReader.Item(cnCUSTCHTIME))
                    vnCUSTGROUP_ = fbuValStr(vnXReader.Item(cnCUSTGROUP_))
                    vnCUSTKOTA = fbuValStr(vnXReader.Item(cnCUSTKOTA))
                    vnCUSTINTERN_GROUP = fbuValStr(vnXReader.Item(cnCUSTINTERN_GROUP))
                    vnCUSTPSR = fbuValStr(vnXReader.Item(cnCUSTPSR))
                    vnCUSTINACTIVE_ = Trim(fbuValStr(vnXReader.Item(cnCUSTINACTIVE_)))
                    vnCUSTOWNER = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTOWNER))))
                    vnCUSTADDRESS_OWNER = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTADDRESS_OWNER))))
                    vnCUSTKOTA_OWNER = fbuValStr(vnXReader.Item(cnCUSTKOTA_OWNER))
                    vnCUSTEMAIL = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTEMAIL))))
                    vnCUSTJABATAN = Trim(fbuValStr(vnXReader.Item(cnCUSTJABATAN)))
                    vnCUSTNPKP = Trim(fbuValStr(vnXReader.Item(cnCUSTNPKP)))
                    vnCUSTGRADE = fbuValStr(vnXReader.Item(cnCUSTGRADE))
                    vnCUSTKODERETUR = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTKODERETUR)))
                    vnCUSTSLSM = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTSLSM)))
                    vnCUSTKOLEKTOR = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTKOLEKTOR)))
                    vnCUSTBILLTO = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnCUSTBILLTO))))
                    vnCUSTSENDTO = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTSENDTO)))
                    vnCUSTISFPJ = Trim(fbuValStr(vnXReader.Item(cnCUSTISFPJ)))
                    vnCUSTREGION = fbuValStr(vnXReader.Item(cnCUSTREGION))
                    vnCUSTHR_MSK_FAKTUR = fbuValStr(vnXReader.Item(cnCUSTHR_MSK_FAKTUR))
                    vnCUSTHR_PENAGIHAN = fbuValStr(vnXReader.Item(cnCUSTHR_PENAGIHAN))
                    vnCUSTHR_PEMBAYARAN = fbuValStr(vnXReader.Item(cnCUSTHR_PEMBAYARAN))
                    vnCUSTTERMS = fbuValStr(vnXReader.Item(cnCUSTTERMS))
                    vnCUSTDISC = fbuValStr(vnXReader.Item(cnCUSTDISC))
                    vnCUSTTAX = fbuValStr(vnXReader.Item(cnCUSTTAX))
                    vnCUSTCOA = fbuValStr(vnXReader.Item(cnCUSTCOA))
                    vnCUSTJNSDEFAULT = fbuValStr(vnXReader.Item(cnCUSTJNSDEFAULT))
                    vnCUSTAGAMA = fbuValStr(vnXReader.Item(cnCUSTAGAMA))
                    vnCUSTTGL_LAHIR = fbuFormatDateDMY_To_YMD_Null(fbuValStr(vnXReader.Item(cnCUSTTGL_LAHIR)))
                    vnCUSTKDTR_FP = fbuValStr(vnXReader.Item(cnCUSTKDTR_FP))
                    vnCUSTDESA = fbuValStr(vnXReader.Item(cnCUSTDESA))
                    vnCUSTSLSMNAME = fbuFormatString(fbuValStr(vnXReader.Item(cnCUSTSLSMNAME)))
                    vnCUSTKOTANAMA = fbuValStr(vnXReader.Item(cnCUSTKOTANAMA))

                    vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstCustomer_MA("
                    vnQuery += vbCrLf & "CompanyCode,CUSTSUB,CUSTNAME,CUSTPERSON,"
                    vnQuery += vbCrLf & "CUSTPHONE,CUSTFAX,CUSTADDRESS,CUSTPOST,"
                    vnQuery += vbCrLf & "CUSTNPWP,CUSTREM,CUSTMAX_,CUSTPRK,"
                    vnQuery += vbCrLf & "CUSTPRKUM,CUSTREV_PRK,CUSTRJL_PRK,CUSTDISC_PRK,"
                    vnQuery += vbCrLf & "CUSTCHUSR,CUSTCHTIME,CUSTGROUP_,CUSTKOTA,"
                    vnQuery += vbCrLf & "CUSTINTERN_GROUP,CUSTPSR,CUSTINACTIVE_,CUSTOWNER,"
                    vnQuery += vbCrLf & "CUSTADDRESS_OWNER,CUSTKOTA_OWNER,CUSTEMAIL,CUSTJABATAN,"
                    vnQuery += vbCrLf & "CUSTNPKP,CUSTGRADE,CUSTKODERETUR,CUSTSLSM,"
                    vnQuery += vbCrLf & "CUSTKOLEKTOR,CUSTBILLTO,CUSTSENDTO,CUSTISFPJ,"
                    vnQuery += vbCrLf & "CUSTREGION,CUSTHR_MSK_FAKTUR,CUSTHR_PENAGIHAN,CUSTHR_PEMBAYARAN,"
                    vnQuery += vbCrLf & "CUSTTERMS,CUSTDISC,CUSTTAX,CUSTCOA,"
                    vnQuery += vbCrLf & "CUSTJNSDEFAULT,CUSTAGAMA,CUSTTGL_LAHIR,CUSTKDTR_FP,"
                    vnQuery += vbCrLf & "CUSTDESA,CUSTSLSMNAME,CUSTKOTANAMA,CustXlsFileOID)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnCUSTSUB & "' vnCUSTSUB,'" & vnCUSTNAME & "' vnCUSTNAME,'" & vnCUSTPERSON & "' vnCUSTPERSON,"
                    vnQuery += vbCrLf & "'" & vnCUSTPHONE & "' vnCUSTPHONE,'" & vnCUSTFAX & "' vnCUSTFAX,'" & vnCUSTADDRESS & "' vnCUSTADDRESS,'" & vnCUSTPOST & "' vnCUSTPOST,"
                    vnQuery += vbCrLf & "'" & vnCUSTNPWP & "' vnCUSTNPWP,'" & vnCUSTREM & "' vnCUSTREM,'" & vnCUSTMAX_ & "' vnCUSTMAX_,'" & vnCUSTPRK & "' vnCUSTPRK,"
                    vnQuery += vbCrLf & "'" & vnCUSTPRKUM & "' vnCUSTPRKUM,'" & vnCUSTREV_PRK & "' vnCUSTREV_PRK,'" & vnCUSTRJL_PRK & "' vnCUSTRJL_PRK,'" & vnCUSTDISC_PRK & "' vnCUSTDISC_PRK,"
                    vnQuery += vbCrLf & "'" & vnCUSTCHUSR & "' vnCUSTCHUSR,'" & vnCUSTCHTIME & "' vnCUSTCHTIME,'" & vnCUSTGROUP_ & "' vnCUSTGROUP_,'" & vnCUSTKOTA & "' vnCUSTKOTA,"
                    vnQuery += vbCrLf & "'" & vnCUSTINTERN_GROUP & "' vnCUSTINTERN_GROUP,'" & vnCUSTPSR & "' vnCUSTPSR,'" & vnCUSTINACTIVE_ & "' vnCUSTINACTIVE_,'" & vnCUSTOWNER & "' vnCUSTOWNER,"
                    vnQuery += vbCrLf & "'" & vnCUSTADDRESS_OWNER & "' vnCUSTADDRESS_OWNER,'" & vnCUSTKOTA_OWNER & "' vnCUSTKOTA_OWNER,'" & vnCUSTEMAIL & "' vnCUSTEMAIL,'" & vnCUSTJABATAN & "' vnCUSTJABATAN,"
                    vnQuery += vbCrLf & "'" & vnCUSTNPKP & "' vnCUSTNPKP,'" & vnCUSTGRADE & "' vnCUSTGRADE,'" & vnCUSTKODERETUR & "' vnCUSTKODERETUR,'" & vnCUSTSLSM & "' vnCUSTSLSM,"
                    vnQuery += vbCrLf & "'" & vnCUSTKOLEKTOR & "' vnCUSTKOLEKTOR,'" & vnCUSTBILLTO & "' vnCUSTBILLTO,'" & vnCUSTSENDTO & "' vnCUSTSENDTO,'" & vnCUSTISFPJ & "' vnCUSTISFPJ,"
                    vnQuery += vbCrLf & "'" & vnCUSTREGION & "' vnCUSTREGION,'" & vnCUSTHR_MSK_FAKTUR & "' vnCUSTHR_MSK_FAKTUR,'" & vnCUSTHR_PENAGIHAN & "' vnCUSTHR_PENAGIHAN,'" & vnCUSTHR_PEMBAYARAN & "' vnCUSTHR_PEMBAYARAN,"
                    vnQuery += vbCrLf & "'" & vnCUSTTERMS & "' vnCUSTTERMS,'" & vnCUSTDISC & "' vnCUSTDISC,'" & vnCUSTTAX & "' vnCUSTTAX,'" & vnCUSTCOA & "' vnCUSTCOA,"
                    vnQuery += vbCrLf & "'" & vnCUSTJNSDEFAULT & "' vnCUSTJNSDEFAULT,'" & vnCUSTAGAMA & "' vnCUSTAGAMA," & vnCUSTTGL_LAHIR & " vnCUSTTGL_LAHIR,'" & vnCUSTKDTR_FP & "' vnCUSTKDTR_FP,"
                    vnQuery += vbCrLf & "'" & vnCUSTDESA & "' vnCUSTDESA,'" & vnCUSTSLSMNAME & "' vnCUSTSLSMNAME,'" & vnCUSTKOTANAMA & "' vnCUSTKOTANAMA,'" & vnHOID & "' vnHOID"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("CUSTSUB ALREADY EXIST")
                End If
            Else
                vsTextStream.WriteLine("CUSTSUB = EMPTY --> EXIT LOOP")
                Exit While
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData_20221212_Orig_Kolom_Lengkap = True
    End Function
End Class