Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class WbfSsoProductMs
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoProductMs"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Dim vsSheetName As String
    Dim vsXlsFolder As String
    Dim vsXlsFileName As String

    Enum ensColList
        CompanyCode = 0
        BRGCODE = 1
        BRGNAME = 2
        BRGUNIT = 3
        vConfirmNotSN = 4
        vIsSN = 5
        vIsActive = 6
        vHS = 7
    End Enum
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

            ViewState("UGAccess") = fbuGetDtbUGAccess(stuTransCode.SsoMsBarang, vnSQLConn)

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

        If ChkListIsSNNew.Checked Then
        Else
            If ChkListSN_Yes.Checked = False And ChkListSN_No.Checked = False Then
                ChkListSN_Yes.Checked = True
                ChkListSN_No.Checked = True
            End If
            If ChkListActive_Yes.Checked = False And ChkListActive_No.Checked = False Then
                ChkListActive_Yes.Checked = True
                ChkListActive_No.Checked = True
            End If
        End If

        Dim vnCompanyCode As String = Session("UserCompanyCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.CompanyCode,PM.BRGCODE,PM.BRGNAME,PM.BRGUNIT,"

        If ChkListIsSNNew.Checked Then
            vnQuery += vbCrLf & "      'Confirm Not SN'vConfirmNotSN,"
        Else
            vnQuery += vbCrLf & "      ''vConfirmNotSN,"
        End If

        vnQuery += vbCrLf & "      case when abs(IsSN)=0 then 'N' else 'Y' end vIsSN,"
        vnQuery += vbCrLf & "      case when abs(IsActive)=0 then 'N' else 'Y' end vIsActive,'History'vHS"
        vnQuery += vbCrLf & " From " & fbuGetDBMaster() & "Sys_MstBarang_MA PM"
        vnQuery += vbCrLf & "Where 1=1"
        If DstListCompany.SelectedValue <> "" Then
            vnQuery += vbCrLf & "      and PM.CompanyCode='" & DstListCompany.SelectedValue & "'"
        End If
        If Trim(TxtListKriteria.Text) <> "" Then
            vnQuery += vbCrLf & "      and (PM.BRGCODE like '%" & Trim(TxtListKriteria.Text) & "%' OR PM.BRGNAME like '%" & Trim(TxtListKriteria.Text) & "%')"
        End If

        If ChkListIsSNNew.Checked Then
            vnQuery += vbCrLf & "      and PM.CompanyCode+'x'+PM.BRGCODE in(Select b.CompanyCode+'x'+b.BRGCODE From Sys_SsoMstBarangNew_TMP b Where b.isSNUpdated=0)"
        Else
            If ChkListSN_Yes.Checked = True And ChkListSN_No.Checked = False Then
                vnQuery += vbCrLf & "      and abs(IsSN)=1"
            ElseIf ChkListSN_Yes.Checked = False And ChkListSN_No.Checked = True Then
                vnQuery += vbCrLf & "      and abs(IsSN)=0"
            End If
            If ChkListActive_Yes.Checked = True And ChkListActive_No.Checked = False Then
                vnQuery += vbCrLf & "      and abs(IsActive)=1"
            ElseIf ChkListActive_Yes.Checked = False And ChkListActive_No.Checked = True Then
                vnQuery += vbCrLf & "      and abs(IsActive)=0"
            End If
        End If

        vnQuery += vbCrLf & " Order by PM.BRGNAME"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvList.DataSource = vnDtb
        GrvList.DataBind()

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

    Private Sub psFillGrvHS(vriCompanyCode As String, vriBrgCode As String)
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select PM.OID,case when abs(PM.IsSN)=0 then 'N' else 'Y' end vIsSN,"
        vnQuery += vbCrLf & "      case when abs(PM.IsActive)=0 then 'N' else 'Y' end vIsActive,"
        vnQuery += vbCrLf & "      SM.UserName vEditBy,Convert(varchar(11),PM.IsHSDatetime,106)+' '+Convert(varchar(5),PM.IsHSDatetime,108)vEditAt"
        vnQuery += vbCrLf & " From Sys_SsoMstBarang_HS PM"
        vnQuery += vbCrLf & "      inner join Sys_SsoUser_MA SM on SM.OID=PM.IsHSUserOID"
        vnQuery += vbCrLf & "Where PM.CompanyCode='" & vriCompanyCode & "' and PM.BRGCODE='" & vriBrgCode & "'"
        vnQuery += vbCrLf & "order by PM.OID"
        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

        GrvHS.DataSource = vnDtb
        GrvHS.DataBind()

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
            vsTextStream.WriteLine(Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
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

        Const cnBRGCODE = 0
        Const cnBRGNAME = 1
        Const cnBRGQMIN = 2
        Const cnBRGPRK = 3
        Const cnBRGUNIT = 4
        Const cnBRGJENIS = 5
        Const cnBRGPRICE = 6
        Const cnBRGREM = 7
        Const cnBRGDISCONTINUE = 8
        Const cnBRGCHUSR = 9
        Const cnBRGCHTIME = 10
        Const cnBRGGROUP = 11
        Const cnBRGGRAM = 12
        Const cnBRGPANJANG = 13
        Const cnBRGLEBAR = 14
        Const cnBRGTINGGI = 15
        Const cnBRGDIJUAL = 16
        Const cnBRGDIBELI = 17
        Const cnBRGGOL1 = 18
        Const cnBRGGOL2 = 19
        Const cnBRGJNSBRG = 20
        Const cnBRGGOL = 21
        Const cnBRGJNSPRODUK = 22
        Const cnBRGKELBRG = 23
        Const cnBRGLAUNCHING_DATE = 24
        Const cnBRGGDG = 25
        Const cnBRGSPEC = 26
        Const cnBRGPCL = 27
        Const cnBRGBATASED = 28
        Const cnBRGMINPRICE = 29
        Const cnBRGFAST = 30
        Const cnBRGW_SN = 31
        Const cnBRGPRKSJL = 32
        Const cnBRGVOL = 33
        Const cnBRGLOK_REFF = 34
        Const cnBRGHARGAJUAL = 35
        Const cnIsSN = 36

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGUNIT As String
        Dim vnIsSN As String

        Dim vnUserOID As String = Session("UserOID")

        Dim vnQuery As String
        vnQuery = "Select isnull(max(convert(int,OID)),0)+1 From " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR"

        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR("
        vnQuery += vbCrLf & "OID,BrgXlsFileName,BrgWorkSheetName,UploadDatetime,UploadUserOID)"
        vnQuery += vbCrLf & "values(" & vnHOID & ",'" & vsXlsFileName & "','" & vsSheetName & "',getdate()," & vnUserOID & ")"
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

            vnBRGCODE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGCODE))))

            vnIsSN = "0"

            If vnBRGCODE <> "" Then
                vnQuery = "Select BRGCODE From " & fbuGetDBMaster() & "Sys_MstBarang_MA Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBRGCODE & "'"
                vsTextStream.WriteLine(vnQuery)
                vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGNAME))))
                vnBRGUNIT = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGUNIT))))

                If fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = "" Then
                    vsTextStream.WriteLine("BRGCODE NOT EXIST")

                    vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarang_MA("
                    vnQuery += vbCrLf & "CompanyCode,BRGCODE,BRGNAME,BRGUNIT,IsSN,BrgXlsFileOID)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnIsSN & " vnIsSN,'" & vnHOID & "' vnHOID"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                    vsTextStream.WriteLine("BRGCODE NOT EXIST")

                    vnQuery = "Insert into Sys_SsoMstBarangNew_TMP("
                    vnQuery += vbCrLf & "CompanyCode,BRGCODE,BRGNAME,BRGUNIT,IsSN,BrgXlsFileOID,"
                    vnQuery += vbCrLf & "CreationUserOID,CreationDatetime)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnIsSN & " vnIsSN,'" & vnHOID & "' vnHOID,"
                    vnQuery += vbCrLf & vnUserOID & ",getdate()"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("BRGCODE ALREADY EXIST")

                    vnQuery = "Update " & fbuGetDBMaster() & "Sys_MstBarang_MA Set BRGNAME='" & vnBRGNAME & "',BRGUNIT='" & vnBRGUNIT & "',IsSN=" & vnIsSN & " Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBRGCODE & "'"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                End If
            Else
                vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
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

    Private Function fsXlsImportData_20230710_Masih_Pakai_IsSN(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData_20230710_Masih_Pakai_IsSN = False
        Dim vnCompanyCode As String = DstCompany.SelectedValue

        Const cnBRGCODE = 0
        Const cnBRGNAME = 1
        Const cnBRGQMIN = 2
        Const cnBRGPRK = 3
        Const cnBRGUNIT = 4
        Const cnBRGJENIS = 5
        Const cnBRGPRICE = 6
        Const cnBRGREM = 7
        Const cnBRGDISCONTINUE = 8
        Const cnBRGCHUSR = 9
        Const cnBRGCHTIME = 10
        Const cnBRGGROUP = 11
        Const cnBRGGRAM = 12
        Const cnBRGPANJANG = 13
        Const cnBRGLEBAR = 14
        Const cnBRGTINGGI = 15
        Const cnBRGDIJUAL = 16
        Const cnBRGDIBELI = 17
        Const cnBRGGOL1 = 18
        Const cnBRGGOL2 = 19
        Const cnBRGJNSBRG = 20
        Const cnBRGGOL = 21
        Const cnBRGJNSPRODUK = 22
        Const cnBRGKELBRG = 23
        Const cnBRGLAUNCHING_DATE = 24
        Const cnBRGGDG = 25
        Const cnBRGSPEC = 26
        Const cnBRGPCL = 27
        Const cnBRGBATASED = 28
        Const cnBRGMINPRICE = 29
        Const cnBRGFAST = 30
        Const cnBRGW_SN = 31
        Const cnBRGPRKSJL = 32
        Const cnBRGVOL = 33
        Const cnBRGLOK_REFF = 34
        Const cnBRGHARGAJUAL = 35
        Const cnIsSN = 36

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGUNIT As String
        Dim vnIsSN As String

        Dim vnQuery As String
        vnQuery = "Select isnull(max(convert(int,OID)),0)+1 From " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR"

        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR("
        vnQuery += vbCrLf & "OID,BrgXlsFileName,BrgWorkSheetName,UploadDatetime,UploadUserOID)"
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

            vnBRGCODE = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGCODE))))
            vnIsSN = IIf(Trim(fbuValStr(vnXReader.Item(cnIsSN))) = "1", "1", "0")

            If vnBRGCODE <> "" Then
                vnQuery = "Select BRGCODE From " & fbuGetDBMaster() & "Sys_MstBarang_MA Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBRGCODE & "'"
                vsTextStream.WriteLine(vnQuery)
                vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGNAME))))
                vnBRGUNIT = Trim(fbuValStr(vnXReader.Item(cnBRGUNIT)))

                If fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = "" Then
                    vsTextStream.WriteLine("BRGCODE NOT EXIST")

                    vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarang_MA("
                    vnQuery += vbCrLf & "CompanyCode,BRGCODE,BRGNAME,BRGUNIT,IsSN,BrgXlsFileOID)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,'" & vnBRGUNIT & "' vnBRGUNIT," & vnIsSN & " vnIsSN,'" & vnHOID & "' vnHOID"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("BRGCODE ALREADY EXIST")

                    vnQuery = "Update " & fbuGetDBMaster() & "Sys_MstBarang_MA Set BRGNAME='" & vnBRGNAME & "',BRGUNIT='" & vnBRGUNIT & "',IsSN=" & vnIsSN & " Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBRGCODE & "'"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

                End If
            Else
                vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
                Exit While
            End If
        End While
        vsTextStream.WriteLine("--------------------Loop...End--------------------")
        vsTextStream.WriteLine("")
        vnXReader.Close()
        vnXCommand.Dispose()

        vnXConn.Close()

        LblXlsProses.Text = "Upload Data Selesai..." & vnNo & "Data"

        fsXlsImportData_20230710_Masih_Pakai_IsSN = True
    End Function


    Private Function fsXlsImportData_20221212_Orig_Kolom_Lengkap(vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As Boolean
        fsXlsImportData_20221212_Orig_Kolom_Lengkap = False
        Dim vnCompanyCode As String = Session("UserCompanyCode")

        Const cnBRGCODE = 0
        Const cnBRGNAME = 1
        Const cnBRGQMIN = 2
        Const cnBRGPRK = 3
        Const cnBRGUNIT = 4
        Const cnBRGJENIS = 5
        Const cnBRGPRICE = 6
        Const cnBRGREM = 7
        Const cnBRGDISCONTINUE = 8
        Const cnBRGCHUSR = 9
        Const cnBRGCHTIME = 10
        Const cnBRGGROUP = 11
        Const cnBRGGRAM = 12
        Const cnBRGPANJANG = 13
        Const cnBRGLEBAR = 14
        Const cnBRGTINGGI = 15
        Const cnBRGDIJUAL = 16
        Const cnBRGDIBELI = 17
        Const cnBRGGOL1 = 18
        Const cnBRGGOL2 = 19
        Const cnBRGJNSBRG = 20
        Const cnBRGGOL = 21
        Const cnBRGJNSPRODUK = 22
        Const cnBRGKELBRG = 23
        Const cnBRGLAUNCHING_DATE = 24
        Const cnBRGGDG = 25
        Const cnBRGSPEC = 26
        Const cnBRGPCL = 27
        Const cnBRGBATASED = 28
        Const cnBRGMINPRICE = 29
        Const cnBRGFAST = 30
        Const cnBRGW_SN = 31
        Const cnBRGPRKSJL = 32
        Const cnBRGVOL = 33
        Const cnBRGLOK_REFF = 34
        Const cnBRGHARGAJUAL = 35

        Dim vnBRGCODE As String
        Dim vnBRGNAME As String
        Dim vnBRGQMIN As String
        Dim vnBRGPRK As String
        Dim vnBRGUNIT As String
        Dim vnBRGJENIS As String
        Dim vnBRGPRICE As String
        Dim vnBRGREM As String
        Dim vnBRGDISCONTINUE As String
        Dim vnBRGCHUSR As String
        Dim vnBRGCHTIME As String
        Dim vnBRGGROUP As String
        Dim vnBRGGRAM As String
        Dim vnBRGPANJANG As String
        Dim vnBRGLEBAR As String
        Dim vnBRGTINGGI As String
        Dim vnBRGDIJUAL As String
        Dim vnBRGDIBELI As String
        Dim vnBRGGOL1 As String
        Dim vnBRGGOL2 As String
        Dim vnBRGJNSBRG As String
        Dim vnBRGGOL As String
        Dim vnBRGJNSPRODUK As String
        Dim vnBRGKELBRG As String
        Dim vnBRGLAUNCHING_DATE As String
        Dim vnBRGGDG As String
        Dim vnBRGSPEC As String
        Dim vnBRGPCL As String
        Dim vnBRGBATASED As String
        Dim vnBRGMINPRICE As String
        Dim vnBRGFAST As String
        Dim vnBRGW_SN As String
        Dim vnBRGPRKSJL As String
        Dim vnBRGVOL As String
        Dim vnBRGLOK_REFF As String
        Dim vnBRGHARGAJUAL As String

        Dim vnQuery As String
        vnQuery = "Select isnull(max(OID),0)+1 From " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR"

        Dim vnHOID As Integer = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

        vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarangUpload_TR("
        vnQuery += vbCrLf & "OID,BrgXlsFileName,BrgWorkSheetName,UploadDatetime,UploadUserOID)"
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

            vnBRGCODE = Trim(fbuValStr(vnXReader.Item(cnBRGCODE)))

            If vnBRGCODE <> "" Then
                vnQuery = "Select BRGCODE From " & fbuGetDBMaster() & "Sys_MstBarang_MA Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBRGCODE & "'"
                vsTextStream.WriteLine(vnQuery)
                If fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) = "" Then
                    vsTextStream.WriteLine("BRGCODE NOT EXIST")
                    vnBRGNAME = fbuFormatString(Trim(fbuValStr(vnXReader.Item(cnBRGNAME))))
                    vnBRGQMIN = fbuValNum(vnXReader.Item(cnBRGQMIN))
                    vnBRGPRK = Trim(fbuValStr(vnXReader.Item(cnBRGPRK)))
                    vnBRGUNIT = Trim(fbuValStr(vnXReader.Item(cnBRGUNIT)))
                    vnBRGJENIS = Trim(fbuValStr(vnXReader.Item(cnBRGJENIS)))
                    vnBRGPRICE = Val(fbuValNum(vnXReader.Item(cnBRGPRICE)))
                    vnBRGREM = Trim(fbuValStr(vnXReader.Item(cnBRGREM)))
                    vnBRGDISCONTINUE = fbuValNum(vnXReader.Item(cnBRGDISCONTINUE))
                    vnBRGCHUSR = Trim(fbuValStr(vnXReader.Item(cnBRGCHUSR)))
                    vnBRGCHTIME = fbuValNullStr(vnXReader.Item(cnBRGCHTIME))
                    vnBRGGROUP = fbuValNum(vnXReader.Item(cnBRGGROUP))
                    vnBRGGRAM = fbuValNum(vnXReader.Item(cnBRGGRAM))
                    vnBRGPANJANG = fbuValNum(vnXReader.Item(cnBRGPANJANG))
                    vnBRGLEBAR = fbuValNum(vnXReader.Item(cnBRGLEBAR))
                    vnBRGTINGGI = fbuValNum(vnXReader.Item(cnBRGTINGGI))
                    vnBRGDIJUAL = fbuValNum(vnXReader.Item(cnBRGDIJUAL))
                    vnBRGDIBELI = fbuValNum(vnXReader.Item(cnBRGDIBELI))
                    vnBRGGOL1 = fbuValNum(vnXReader.Item(cnBRGGOL1))
                    vnBRGGOL2 = fbuValNum(vnXReader.Item(cnBRGGOL2))
                    vnBRGJNSBRG = fbuValNum(vnXReader.Item(cnBRGJNSBRG))
                    vnBRGGOL = Trim(fbuValStr(vnXReader.Item(cnBRGGOL)))
                    vnBRGJNSPRODUK = Trim(fbuValStr(vnXReader.Item(cnBRGJNSPRODUK)))
                    vnBRGKELBRG = Trim(fbuValStr(vnXReader.Item(cnBRGKELBRG)))
                    vnBRGLAUNCHING_DATE = fbuValNullStr(vnXReader.Item(cnBRGLAUNCHING_DATE))
                    vnBRGGDG = Trim(fbuValStr(vnXReader.Item(cnBRGGDG)))
                    vnBRGSPEC = Trim(fbuValStr(vnXReader.Item(cnBRGSPEC)))
                    vnBRGPCL = Trim(fbuValStr(vnXReader.Item(cnBRGPCL)))
                    vnBRGBATASED = fbuValNum(vnXReader.Item(cnBRGBATASED))
                    vnBRGMINPRICE = fbuValNum(vnXReader.Item(cnBRGMINPRICE))
                    vnBRGFAST = fbuValNum(vnXReader.Item(cnBRGFAST))
                    vnBRGW_SN = fbuValNum(vnXReader.Item(cnBRGW_SN))
                    vnBRGPRKSJL = Trim(fbuValStr(vnXReader.Item(cnBRGPRKSJL)))
                    vnBRGVOL = fbuValNum(vnXReader.Item(cnBRGVOL))
                    vnBRGLOK_REFF = Trim(fbuValStr(vnXReader.Item(cnBRGLOK_REFF)))
                    vnBRGHARGAJUAL = fbuValNum(vnXReader.Item(cnBRGHARGAJUAL))

                    vnQuery = "Insert into " & fbuGetDBMaster() & "Sys_MstBarang_MA("
                    vnQuery += vbCrLf & "CompanyCode,BRGCODE,BRGNAME,"
                    vnQuery += vbCrLf & "BRGQMIN,BRGPRK,BRGUNIT,BRGJENIS,"
                    vnQuery += vbCrLf & "BRGPRICE,BRGREM,BRGDISCONTINUE,BRGCHUSR,"
                    vnQuery += vbCrLf & "BRGCHTIME,BRGGROUP,BRGGRAM,BRGPANJANG,"
                    vnQuery += vbCrLf & "BRGLEBAR,BRGTINGGI,BRGDIJUAL,BRGDIBELI,"
                    vnQuery += vbCrLf & "BRGGOL1,BRGGOL2,BRGJNSBRG,BRGGOL,"
                    vnQuery += vbCrLf & "BRGJNSPRODUK,BRGKELBRG,BRGLAUNCHING_DATE,BRGGDG,"
                    vnQuery += vbCrLf & "BRGSPEC,BRGPCL,BRGBATASED,BRGMINPRICE,BRGFAST,"
                    vnQuery += vbCrLf & "BRGW_SN,BRGPRKSJL,BRGVOL,BRGLOK_REFF,BRGHARGAJUAL,BrgXlsFileOID)"
                    vnQuery += vbCrLf & "Select '" & vnCompanyCode & "' vnCompanyCode,'" & vnBRGCODE & "' vnBRGCODE,'" & vnBRGNAME & "' vnBRGNAME,"
                    vnQuery += vbCrLf & "'" & vnBRGQMIN & "' vnBRGQMIN,'" & vnBRGPRK & "' vnBRGPRK,'" & vnBRGUNIT & "' vnBRGUNIT,'" & vnBRGJENIS & "' vnBRGJENIS,"
                    vnQuery += vbCrLf & "'" & vnBRGPRICE & "' vnBRGPRICE,'" & vnBRGREM & "' vnBRGREM,'" & vnBRGDISCONTINUE & "'vnBRGDISCONTINUE,'" & vnBRGCHUSR & "'vnBRGCHUSR,"
                    vnQuery += vbCrLf & "" & vnBRGCHTIME & " vnBRGCHTIME,'" & vnBRGGROUP & "' vnBRGGROUP,'" & vnBRGGRAM & "' vnBRGGRAM," & vnBRGPANJANG & " vnBRGPANJANG,"
                    vnQuery += vbCrLf & "'" & vnBRGLEBAR & "' vnBRGLEBAR,'" & vnBRGTINGGI & "' vnBRGTINGGI,'" & vnBRGDIJUAL & "' vnBRGDIJUAL,'" & vnBRGDIBELI & "' vnBRGDIBELI,"
                    vnQuery += vbCrLf & "'" & vnBRGGOL1 & "' vnBRGGOL1,'" & vnBRGGOL2 & "' vnBRGGOL2,'" & vnBRGJNSBRG & "' vnBRGJNSBRG,'" & vnBRGGOL & "' vnBRGGOL,"
                    vnQuery += vbCrLf & "'" & vnBRGJNSPRODUK & "' vnBRGJNSPRODUK,'" & vnBRGKELBRG & "' vnBRGKELBRG," & vnBRGLAUNCHING_DATE & " vnBRGLAUNCHING_DATE,'" & vnBRGGDG & "' vnBRGGDG,"
                    vnQuery += vbCrLf & "'" & vnBRGSPEC & "' vnBRGSPEC,'" & vnBRGPCL & "' vnBRGPCL,'" & vnBRGBATASED & "' vnBRGBATASED,'" & vnBRGMINPRICE & "' vnBRGMINPRICE,'" & vnBRGFAST & "'vnBRGFAST,"
                    vnQuery += vbCrLf & "'" & vnBRGW_SN & "' vnBRGW_SN,'" & vnBRGPRKSJL & "' vnBRGPRKSJL,'" & vnBRGVOL & "' vnBRGVOL,'" & vnBRGLOK_REFF & "' vnBRGLOK_REFF,'" & vnBRGHARGAJUAL & "' vnBRGHARGAJUAL,'" & vnHOID & "' vnHOID"

                    vsTextStream.WriteLine(vnQuery)
                    pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
                Else
                    vsTextStream.WriteLine("BRGCODE ALREADY EXIST")
                End If
            Else
                vsTextStream.WriteLine("BRGCODE = EMPTY --> EXIT LOOP")
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

    Protected Sub BtnGenQRCode_Click(sender As Object, e As EventArgs) Handles BtnGenQRCode.Click
        Response.Redirect("WbfSsoProductQR.aspx")
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        Dim vnUserOID As String = Session("UserOID")
        If vnUserOID = "" Then
            Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        End If

        Dim vnCompanyCode As String
        Dim vnBrgCode As String
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)
        vnCompanyCode = vnGRow.Cells(ensColList.CompanyCode).Text
        vnBrgCode = vnGRow.Cells(ensColList.BRGCODE).Text

        If e.CommandName = "vIsSN" Or e.CommandName = "vIsActive" Or e.CommandName = "vConfirmNotSN" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgListError.Text = pbMsgError
                LblMsgListError.Visible = True
                Exit Sub
            End If

            Dim vnSQLTrans As SqlTransaction = Nothing
            Dim vnBeginTrans As Boolean
            Try
                Dim vnDBMaster As String = fbuGetDBMaster()
                Dim vnQuery As String

                vnSQLTrans = vnSQLConn.BeginTransaction("upd")
                vnBeginTrans = True

                If e.CommandName = "vIsSN" Then
                    vnQuery = "Update " & vnDBMaster & "Sys_MstBarang_MA Set isSN=case when isSN=0 then 1 else 0 end Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                    vnQuery = "Update Sys_SsoMstBarangNew_TMP Set isSNUpdated=1 Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                    vnQuery = "insert into Sys_SsoMstBarang_HS"
                    vnQuery += vbCrLf & "(CompanyCode,BRGCODE,isSN,isActive,IsHSUserOID,IsHSDatetime)"
                    vnQuery += vbCrLf & "Select CompanyCode,BRGCODE,isSN,isActive," & vnUserOID & ",getdate() From " & vnDBMaster & "Sys_MstBarang_MA Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                ElseIf e.CommandName = "vIsActive" Then
                    vnQuery = "Update " & vnDBMaster & "Sys_MstBarang_MA Set isActive=case when isActive=0 then 1 else 0 end Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                    vnQuery = "insert into Sys_SsoMstBarang_HS"
                    vnQuery += vbCrLf & "(CompanyCode,BRGCODE,isSN,isActive,IsHSUserOID,IsHSDatetime)"
                    vnQuery += vbCrLf & "Select CompanyCode,BRGCODE,isSN,isActive," & vnUserOID & ",getdate() From " & vnDBMaster & "Sys_MstBarang_MA Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                ElseIf e.CommandName = "vConfirmNotSN" Then
                    vnQuery = "Update Sys_SsoMstBarangNew_TMP Set isSNUpdated=1 Where CompanyCode='" & vnCompanyCode & "' and BRGCODE='" & vnBrgCode & "'"
                    pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vnSQLConn, vnSQLTrans)

                End If

                vnSQLTrans.Commit()
                vnSQLTrans.Dispose()
                vnSQLTrans = Nothing
                vnBeginTrans = False

                If e.CommandName = "vIsSN" Then
                    DirectCast(vnGRow.Cells(ensColList.vIsSN).Controls(0), LinkButton).Text = IIf(DirectCast(vnGRow.Cells(ensColList.vIsSN).Controls(0), LinkButton).Text = "Y", "N", "Y")
                ElseIf e.CommandName = "vIsActive" Then
                    DirectCast(vnGRow.Cells(ensColList.vIsActive).Controls(0), LinkButton).Text = IIf(DirectCast(vnGRow.Cells(ensColList.vIsActive).Controls(0), LinkButton).Text = "Y", "N", "Y")
                Else
                    vnGRow.Visible = False
                End If

            Catch ex As Exception
                LblMsgListError.Text = ex.Message
                LblMsgListError.Visible = True

                If vnBeginTrans Then
                    vnSQLTrans.Rollback()
                    vnSQLTrans.Dispose()
                    vnSQLTrans = Nothing
                End If
            Finally
                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
            End Try

        ElseIf e.CommandName = "vHS" Then
            LblHS.Text = vnCompanyCode & " " & vnBrgCode & "<br />" & vnGRow.Cells(ensColList.BRGNAME).Text
            psFillGrvHS(vnCompanyCode, vnBrgCode)
        End If
    End Sub

    Protected Sub ChkListIsSNNew_CheckedChanged(sender As Object, e As EventArgs) Handles ChkListIsSNNew.CheckedChanged
        If ChkListIsSNNew.Checked Then
            ChkListSN_No.Enabled = False
            ChkListSN_Yes.Enabled = False

            GrvList.Columns(ensColList.vConfirmNotSN).HeaderStyle.CssClass = ""
            GrvList.Columns(ensColList.vConfirmNotSN).ItemStyle.CssClass = ""
        Else
            ChkListSN_No.Enabled = True
            ChkListSN_Yes.Enabled = True

            GrvList.Columns(ensColList.vConfirmNotSN).HeaderStyle.CssClass = "myDisplayNone"
            GrvList.Columns(ensColList.vConfirmNotSN).ItemStyle.CssClass = "myDisplayNone"
        End If

        psFillGrvList()
    End Sub
End Class