Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Module ModSsoClosedXML
    Const csModuleName = "ModSsoClosedXML"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileNameOnly As String
    Dim vsLogFileName As String
    Dim vsLogFileNameError As String

    Public Sub pbuCreateXlsx_SOTally(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTally", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,mc.CompanyName,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTally(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.vSOStockScanVarian!=0"
            End If

            vnQuery += vbCrLf & " order by mb.BRGNAME"
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
            Dim vnSONo As String = ""
            Dim vnCompName As String = ""
            Dim vnWarehouseName As String = ""
            Dim vnSubWhsName As String = ""
            Dim vnSOID As String = ""
            Dim vnSOCutOff As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote As String = ""
            Dim vnSOCloseNote As String = ""

            If vnDtb.Rows.Count = 0 Then
                vnDRow = vnDtb.Rows(0)
            Else
                vnDRow = vnDtb.Rows(0)
                vnSONo = vnDRow.Item("SONo")
                vnCompName = vnDRow.Item("CompanyName")
                vnWarehouseName = vnDRow.Item("WarehouseName")
                vnSubWhsName = vnDRow.Item("vSubWhsName")
                vnSOID = vnDRow.Item("OID")
                vnSOCutOff = vnDRow.Item("SOCutOff")
                vnStatus = vnDRow.Item("TransStatusDescr")
                vnSONote = vnDRow.Item("SONote")
                vnSOCloseNote = fbuValStr(vnDRow.Item("SOCloseNote"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME - TALLY REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME - TALLY REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'NO SO
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo

            'Company
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'SOID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOID

            '<---------------ROW 5
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff

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
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONote

            'Sub Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWhsName

            'CloseNote
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCloseNote")

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 10
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Qty SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumSOScanQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockScanVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockNoteDatetime")
                Next

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 7
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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

    Public Sub pbuCreateXlsx_SOTallyDetail(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTally", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyDetail(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('') stg on stg.vStorageOID=ta.StorageOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.BRGCODE in(Select b.BRGCODE From fnTbl_SsoTally(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "') b Where b.vSOStockScanVarian!=0)"
            End If

            vnQuery += vbCrLf & " order by mb.BRGNAME"
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
            Dim vnSONo As String = ""
            Dim vnCompName As String = ""
            Dim vnWarehouseName As String = ""
            Dim vnSubWhsName As String = ""
            Dim vnSOID As String = ""
            Dim vnSOCutOff As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote As String = ""
            Dim vnSOCloseNote As String = ""

            If vnDtb.Rows.Count = 0 Then
                vnDRow = vnDtb.Rows(0)
            Else
                vnDRow = vnDtb.Rows(0)
                vnSONo = vnDRow.Item("SONo")
                vnCompName = vnDRow.Item("CompanyName")
                vnWarehouseName = vnDRow.Item("WarehouseName")
                vnSubWhsName = vnDRow.Item("vSubWhsName")
                vnSOID = vnDRow.Item("OID")
                vnSOCutOff = vnDRow.Item("SOCutOff")
                vnStatus = vnDRow.Item("TransStatusDescr")
                vnSONote = vnDRow.Item("SONote")
                vnSOCloseNote = fbuValStr(vnDRow.Item("SOCloseNote"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME - DETAIL REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME - DETAIL REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'NO SO
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo

            'Company
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'SOID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOID

            '<---------------ROW 5
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff

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
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONote

            'Sub Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWhsName

            'CloseNote
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCloseNote")

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 10
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Info"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Note Scan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Scan By"

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
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vScanByName")
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
    Public Sub pbuCreateXlsx_SOTallyCompare(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompare", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
            vnQuery += vbCrLf & "            convert(varchar(11),ta.LastCompareDatetime,106)+' '+convert(varchar(8),ta.LastCompareDatetime,108)vLastCompareDatetime"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyCompare(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.vSOScanQtyVarian!=0"
            End If

            vnQuery += vbCrLf & " order by mb.BRGNAME"
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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""
            Dim vnLastCompareDatetime As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("vSubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))

                vsTextStream.WriteLine("15")
                vnLastCompareDatetime = fbuValStr(vnDRow.Item("vLastCompareDatetime"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            'Last Compare
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Last Compare"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnLastCompareDatetime

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 15
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty" & vbCrLf & "Stock 1 - Stock 2"
            vnIXLWorksheet.Columns(vnXCol).Width = "15"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty" & vbCrLf & "Scan 1 - Scan 2"
            vnIXLWorksheet.Columns(vnXCol).Width = "15"

            '<---23 Sep 2023
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty SO 1" & vbCrLf & "(Qty Stock 1 - Qty Scan 1)"
            vnIXLWorksheet.Columns(vnXCol).Width = "22"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty SO 2" & vbCrLf & "(Qty Stock 2 - Qty Scan 2)"
            vnIXLWorksheet.Columns(vnXCol).Width = "22"
            '<==23 Sep 2023

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQtyVarian")

                    '<---23 Sep 2023
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty1") - vnDRow.Item("SOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty2") - vnDRow.Item("SOScanQty2")
                    '<<==23 Sep 2023

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCompareDNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteDatetime")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 12
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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

    Public Sub pbuCreateXlsx_SOTallyCompare_20230923_Bef_Ditambah_Column_Selisih_Reqby_Pak_Mikha(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompare", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT,"
            vnQuery += vbCrLf & "            convert(varchar(11),ta.LastCompareDatetime,106)+' '+convert(varchar(8),ta.LastCompareDatetime,108)vLastCompareDatetime"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyCompare(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.vSOScanQtyVarian!=0"
            End If

            vnQuery += vbCrLf & " order by mb.BRGNAME"
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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""
            Dim vnLastCompareDatetime As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("vSubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))

                vsTextStream.WriteLine("15")
                vnLastCompareDatetime = fbuValStr(vnDRow.Item("vLastCompareDatetime"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            'Last Compare
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Last Compare"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnLastCompareDatetime

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 13
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Scan"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCompareDNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteDatetime")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 10
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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

    Public Sub pbuCreateXlsx_SOTallyCompareDetail(ByRef vriFileName As String, vriTransID As String, vriSOHOID1 As String, vriSOHOID2 As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompareDetail", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyCompareDetail(" & vriTransID & "," & vriSOHOID1 & "," & vriSOHOID2 & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('') stg on stg.vStorageOID=ta.vStorageOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.vSOScanVarian!=0"
            End If
            vnQuery += vbCrLf & " order by mb.BRGNAME,stg.vStorageInfo"

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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""
            Dim vnLastCompareDatetime As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("SubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))

                vsTextStream.WriteLine("15")
                vnLastCompareDatetime = fbuValStr(vnDRow.Item("LastCompareDatetime"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - DETAIL REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - DETAIL REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            'Last Compare
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Last Compare"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnLastCompareDatetime

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 13
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Info"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Scan"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Scan By SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Scan By SO 2"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vScanByName1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vScanByName2")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 6 To 8
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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


    Public Sub pbuCreateXlsx_SOTally_20230203_Orig_Bef_Varian_Only(ByRef vriFileName As String, vriTransID As String, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTally", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,mc.CompanyName,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTally(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"

            vnQuery += vbCrLf & " order by mb.BRGNAME"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "SBSOTally-OID-" & vriTransID & "-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            Dim vnSONo As String = ""
            Dim vnCompName As String = ""
            Dim vnWarehouseName As String = ""
            Dim vnSubWhsName As String = ""
            Dim vnSOID As String = ""
            Dim vnSOCutOff As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote As String = ""
            Dim vnSOCloseNote As String = ""

            If vnDtb.Rows.Count = 0 Then
                vnDRow = vnDtb.Rows(0)
            Else
                vnDRow = vnDtb.Rows(0)
                vnSONo = vnDRow.Item("SONo")
                vnCompName = vnDRow.Item("CompanyName")
                vnWarehouseName = vnDRow.Item("WarehouseName")
                vnSubWhsName = vnDRow.Item("vSubWhsName")
                vnSOID = vnDRow.Item("OID")
                vnSOCutOff = vnDRow.Item("SOCutOff")
                vnStatus = vnDRow.Item("TransStatusDescr")
                vnSONote = vnDRow.Item("SONote")
                vnSOCloseNote = fbuValStr(vnDRow.Item("SOCloseNote"))
            End If

            '<---------------ROW 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMBER BERKAT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 2
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME - TALLY REPORT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'NO SO
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo

            'Company
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'SOID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOID

            '<---------------ROW 5
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff

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
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONote

            'Sub Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWhsName

            'CloseNote
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCloseNote")

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 10
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Qty SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumSOScanQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockScanVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockNoteDatetime")
                Next

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 7
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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

    Public Sub pbuCreateXlsx_SOTallyCompare_20230203_Orig_Bef_Only_Selisih(ByRef vriFileName As String, vriTransID As String, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompare", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode + ' - ' +sw.SubWhsName vSubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyCompare(" & vriTransID & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            vnQuery += vbCrLf & " order by mb.BRGNAME"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "SOTallyCompare-OID-" & vriTransID & "-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("vSubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))
            End If

            '<---------------ROW 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMBER BERKAT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 2
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 13
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Scan"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCompareDNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteDatetime")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 10
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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
    Public Sub pbuCreateXlsx_SOTallyCompare2(ByRef vriFileName As String, vriTransID As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            Dim vnDBMaster As String = fbuGetDBMaster()
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompare", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "	DECLARE @vriUser AS VARCHAR(25)	"
            vnQuery += vbCrLf & "	DECLARE @vriHOID AS INTEGER 	"
            vnQuery += vbCrLf & "	SET @vriUser = '" & HttpContext.Current.Session("UserID") & "'	"
            vnQuery += vbCrLf & "	SET @vriHOID = " & vriTransID & "	"
            vnQuery += vbCrLf & "	Select sh.OID,	"
            vnQuery += vbCrLf & "	       sh.SOCompanyCode,sh.SOWarehouseOID,sh.SOSubWarehouseOID, mc.CompanyName, sm.SubWhsName vSubWhsName, SD.SOCompareDNoteDatetime vLastCompareDateTime,	"
            vnQuery += vbCrLf & "	       sh.SOCompareCloseNote,sh.SOCompareCancelNote,	"
            vnQuery += vbCrLf & "	       sh.SOHOID1,so1.SONo vSONo1,so1.SOCutOff vSOCutOff1,so1.SONote vSONote1,so1.SOCloseNote vSOCloseNote1,so1.SOCancelNote vSOCancelNote1,	"
            vnQuery += vbCrLf & "	       sh.SOHOID2,so2.SONo vSONo2,so2.SOCutOff vSOCutOff2,so2.SONote vSONote2,so2.SOCloseNote vSOCloseNote2,so2.SOCancelNote vSOCancelNote2,	"
            vnQuery += vbCrLf & "	       sh.TransCode,sh.TransStatus,	"
            vnQuery += vbCrLf & "	       st.TransStatusDescr,	"
            vnQuery += vbCrLf & "	       sh.CreationUserOID,sh.CreationDatetime,	"
            vnQuery += vbCrLf & "	       sh.LastCompareUserOID,sh.LastCompareDatetime,	"
            vnQuery += vbCrLf & "	       sh.ClosedUserOID,sh.ClosedDatetime,	"
            vnQuery += vbCrLf & "	       sh.CancelledUserOID,sh.CancelledDatetime,	"
            vnQuery += vbCrLf & "	       sd.OID vDOID,	"
            vnQuery += vbCrLf & "	       sd.BRGCODE,	"
            vnQuery += vbCrLf & "	       sd.SOStockQty1,sd.SOStockQty2,(sd.SOStockQty1-sd.SOStockQty2)vSOStockQtyVarian,	"
            vnQuery += vbCrLf & "	       (sd.SOStockQty1-sd.SOScanQty1)vSOStockScan1,	"
            vnQuery += vbCrLf & "	       sd.SOScanQty1,sd.SOScanQty2,(sd.SOScanQty1-sd.SOScanQty2)vSOScanQtyVarian,	"
            vnQuery += vbCrLf & "	       (sd.SOStockQty2-sd.SOScanQty2)vSOStockScan2,	"
            vnQuery += vbCrLf & "	       (sd.SOStockQty1-sd.SOScanQty1) - (sd.SOStockQty2-sd.SOScanQty2) vSOStockScanAkhir,	"
            vnQuery += vbCrLf & "	       sd.SOCompareDNote,	"
            vnQuery += vbCrLf & "	       sd.SOCompareDNoteUserOID,su.UserName vSOCompareDNoteBy,	"
            vnQuery += vbCrLf & "	       sd.SOCompareDNoteDatetime,	"
            vnQuery += vbCrLf & "	       Convert(varchar(11),sd.SOCompareDNoteDatetime,106)+' '+Convert(varchar(5),sd.SOCompareDNoteDatetime,108) vSOCompareDNoteDatetime,	"
            vnQuery += vbCrLf & "	       Convert(varchar(11),getdate(),106)+' '+Convert(varchar(5),getdate(),108) vPrintDate	"
            vnQuery += vbCrLf & "	       ,@vriUser vPrintUser	"
            vnQuery += vbCrLf & "	  From Sys_SsoSOCompareH_TR sh with(nolock)	"
            vnQuery += vbCrLf & "	       inner join Sys_SsoSOHeader_TR so1 with(nolock) on so1.OID=sh.SOHOID1	"
            vnQuery += vbCrLf & "	       inner join Sys_SsoSOHeader_TR so2 with(nolock) on so2.OID=sh.SOHOID2	"
            vnQuery += vbCrLf & "	       inner join Sys_SsoTransStatus_MA st with(nolock) on st.TransCode=sh.TransCode and st.TransStatus=sh.TransStatus	"
            vnQuery += vbCrLf & "	       inner join Sys_SsoSOCompareD_TR sd with(nolock) on sd.SOCHOID=sh.OID	"
            vnQuery += vbCrLf & "	       left outer join Sys_SsoUser_MA su with(nolock) on su.OID=sd.SOCompareDNoteUserOID	"
            vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=sh.SOCompanyCode	"
            vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sh.SOCompanyCode and mb.BRGCODE=sd.BRGCODE	"
            vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=sh.SOWarehouseOID	"
            vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_SubWarehouse_MA sm with(nolock) on sm.OID=sh.SOSubWarehouseOID	"
            vnQuery += vbCrLf & "	       inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=sh.SOWarehouseOID	"
            vnQuery += vbCrLf & "	   Where sh.OID=@vriHOID	"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       and ta.vSOScanQtyVarian!=0"
            End If

            vnQuery += vbCrLf & " order by mb.BRGNAME"
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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""
            Dim vnLastCompareDatetime As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("vSubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))

                vsTextStream.WriteLine("15")
                vnLastCompareDatetime = fbuValStr(vnDRow.Item("vLastCompareDatetime"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - TALLY REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            'Last Compare
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Last Compare"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnLastCompareDatetime

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 13
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Stock SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock dengan Scan pada SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Scan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock dengan Scan pada SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Stock dengan Scan Antara SO 1 dan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Detail Note Datetime"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOStockQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockScan1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQtyVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockScan2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOStockScanAkhir")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCompareDNote")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteBy")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOCompareDNoteDatetime")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 10
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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

    Public Sub pbuCreateXlsx_SOTallyCompareDetail2(ByRef vriFileName As String, vriTransID As String, vriSOHOID1 As String, vriSOHOID2 As String, vriVarianOnly As Byte, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTallyCompareDetail", vriTransID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "Select ta.*,stg.vStorageInfo,mc.CompanyName,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName,"
            vnQuery += vbCrLf & "            row_number()over(order by mb.BRGNAME)vDSeqNo,mb.BRGNAME,mb.BRGUNIT"
            vnQuery += vbCrLf & "       From fnTbl_SsoTallyCompareDetail(" & vriTransID & "," & vriSOHOID1 & "," & vriSOHOID2 & ",'" & HttpContext.Current.Session("UserID") & "')ta"
            vnQuery += vbCrLf & "	         inner join " & vnDBMaster & "DimCompany mc with(nolock) on mc.CompanyCode=ta.SOCompanyCode"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=ta.SOCompanyCode and mb.BRGCODE=ta.BRGCODE"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=ta.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=ta.SOWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "fnTbl_SsoStorageData('" & HttpContext.Current.Session("UserID") & "') stg on stg.vStorageOID=ta.vStorageOID"
            If vriVarianOnly = 1 Then
                vnQuery += vbCrLf & "       Where ta.vSOScanVarian!=0"
            End If
            vnQuery += vbCrLf & " order by mb.BRGNAME,stg.vStorageInfo"

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
            Dim vnSONo1 As String = ""
            Dim vnSONo2 As String = ""
            Dim vnCompName As String = ""
            Dim vnSubWarehouseName As String = ""
            Dim vnSOCompareID As String = ""
            Dim vnSOCutOff1 As String = ""
            Dim vnSOCutOff2 As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote1 As String = ""
            Dim vnSONote2 As String = ""
            Dim vnSOCloseNote1 As String = ""
            Dim vnSOCloseNote2 As String = ""
            Dim vnSOCloseCompare As String = ""
            Dim vnLastCompareDatetime As String = ""

            vsTextStream.WriteLine("1")

            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("Tidak Ada Data")
                vnDRow = vnDtb.Rows(0)
            Else
                vsTextStream.WriteLine("2")
                vnDRow = vnDtb.Rows(0)

                vsTextStream.WriteLine("3")
                vnSONo1 = vnDRow.Item("vSONo1")

                vsTextStream.WriteLine("3")
                vnSONo2 = vnDRow.Item("vSONo2")

                vsTextStream.WriteLine("4")
                vnCompName = vnDRow.Item("CompanyName")

                vsTextStream.WriteLine("5")
                vnSubWarehouseName = vnDRow.Item("SubWhsName")

                vsTextStream.WriteLine("6")
                vnSOCompareID = vnDRow.Item("OID")

                vsTextStream.WriteLine("7")
                vnSOCutOff1 = vnDRow.Item("vSOCutOff1")

                vsTextStream.WriteLine("8")
                vnSOCutOff2 = vnDRow.Item("vSOCutOff2")

                vsTextStream.WriteLine("9")
                vnStatus = vnDRow.Item("TransStatusDescr")

                vsTextStream.WriteLine("10")
                vnSONote1 = vnDRow.Item("vSONote1")

                vsTextStream.WriteLine("11")
                vnSONote2 = vnDRow.Item("vSONote2")

                vsTextStream.WriteLine("12")
                vnSOCloseNote1 = fbuValStr(vnDRow.Item("vSOCloseNote1"))

                vsTextStream.WriteLine("13")
                vnSOCloseNote2 = fbuValStr(vnDRow.Item("vSOCloseNote2"))

                vsTextStream.WriteLine("14")
                vnSOCloseCompare = fbuValStr(vnDRow.Item("SOCompareCloseNote"))

                vsTextStream.WriteLine("15")
                vnLastCompareDatetime = fbuValStr(vnDRow.Item("LastCompareDatetime"))
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
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - DETAIL REPORT - ALL DATA"
            Else
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK OPNAME COMPARE - DETAIL REPORT - DATA SELISIH"
            End If

            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'Company
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnCompName

            'Gudang
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSubWarehouseName

            'SO Compare ID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Compare  ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCompareID

            '<---------------ROW 5
            'NO SO 1
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo1

            'NO SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSONo2

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnStatus

            '<---------------ROW 6
            'Cutoff SO 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff1

            'Cutoff SO 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Cut Off SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCutOff2

            'Last Compare
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Last Compare"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnLastCompareDatetime

            '<---------------ROW 7
            'SO Close Note 1
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote1

            'SO Close Note 2
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO Close Note 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseNote2

            'SO Compare Close Note
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Compare Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOCloseCompare

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 13
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Satuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Info"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan SO 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih Qty Scan"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Scan By SO 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Scan By SO 2"

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
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("Start Loop")
                Dim vnRow As Integer
                Dim vnBrgCode As String
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vnBrgCode = vnDRow.Item("BRGCODE")

                    vsTextStream.WriteLine("vnRow " & vnRow)
                    vsTextStream.WriteLine("vnBrgCode " & vnBrgCode)

                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSeqNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnBrgCode
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGUNIT")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQty1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanQty2")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSOScanVarian")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vScanByName1")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vScanByName2")
                Next
                vsTextStream.WriteLine("End Loop")
                vsTextStream.WriteLine("")

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 6 To 8
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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
    Public Sub pbuCreateXlsx_MonitoringInvoice(ByRef vriFileName As String, vriDstWarehouse As DropDownList, vriDstCompany As DropDownList, vriInvoiceNo As String, vriNoPicklist As String, vriNoReferensi As String, vriNoPicking As String, vriNoDispatch As String, Chk_Upload As CheckBox, Chk_Picklist As CheckBox, Chk_PickilistPrepared As CheckBox, Chk_Picking As CheckBox, Chk_PickingDone As CheckBox, Chk_Dispatch As CheckBox, Chk_DispatchDone As CheckBox, Chk_DriverConfirm As CheckBox, vriTxtListStart As String, vriTxtListEnd As String, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringInvoice", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)



            Dim vnQuery As String
            Dim vnDtb As New DataTable

            vnQuery = "	select distinct	"
            vnQuery += vbCrLf & "	mj.CompanyCode,whs.WarehouseName, whs.OID, DATEDIFF(HOUR,mj.UploadDatetime,skh.BackDatetime) as [Durasi_Start_to_End]	"
            vnQuery += vbCrLf & "	, mj.NO_NOTA, mj.TANGGAL, mj.KODE_CUST, mj.CUSTOMER, mj.UploadDatetime	"
            vnQuery += vbCrLf & "	, pch.PCLNo, pch.PCLDate, pch.PCLScheduleDate, pch.CreationDatetime as [Time_Create_Picklist], usr_pch.UserName, pch.PreparedDatetime	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pch.CreationDatetime) as [Durasi_Upload_to_Create_Picklist]	"
            vnQuery += vbCrLf & "	, pck.PCKNo, pck.PCKDate, pck.CreationDatetime 'Picking_Created_Date_Time', pck.PickDoneDatetime, pch.PCLRefHOID, pch.PCLRefHNo	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pch.CreationDatetime,pck.PickDoneDatetime) as [Durasi_Picklist_Created_to_Picking_Done]	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pck.PickDoneDatetime) as [Durasi_Upload_to_Picking_Done]	"
            vnQuery += vbCrLf & "	, dsh.DSPNo, dsh.DSPDate, dsh.CreationDatetime 'Dispatch_Created_Date_Time', dsh.DispatchDoneDatetime 'Dispatch_Created_Date'	"
            vnQuery += vbCrLf & "	, dsh.DriverConfirmDatetime 	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pck.PickDoneDatetime,dsh.DriverConfirmDatetime) as [Durasi_Picking_Done_to_Dispatch]	"
            vnQuery += vbCrLf & "	,drv.DcmDriverName,skh.BackDatetime, dsh.CancelledDatetime, pch.TransCode , pch.TransStatus, sstsm.TransStatusDescr	"
            vnQuery += vbCrLf & "	from 	"
            vnQuery += vbCrLf & "	(select	"
            vnQuery += vbCrLf & "	CompanyCode, WarehouseOID, NO_NOTA, TANGGAL, KODE_CUST, CUSTOMER, max(UploadDatetime) as uploadDatetime	"
            vnQuery += vbCrLf & "	from [dbo].[Sys_DcmJUAL]	"
            vnQuery += vbCrLf & "	group by CompanyCode, WarehouseOID, NO_NOTA, TANGGAL, KODE_CUST, CUSTOMER	"
            vnQuery += vbCrLf & "	) as mj	"
            vnQuery += vbCrLf & "	left join [SB_WMS].dbo.[Sys_SsoPCLHeader_TR] pch on pch.PCLRefHNo = mj.NO_NOTA	"
            vnQuery += vbCrLf & "	left join [SB_WMS].dbo.Sys_SsoUser_MA usr_pch on usr_pch.OID = pch.CreationUserOID	"
            vnQuery += vbCrLf & "	left join [SB_DATAWH].dbo.[Sys_Warehouse_MA] whs on whs.OID = mj.WarehouseOID	"
            vnQuery += vbCrLf & "	left join [SB_WMS].dbo.[Sys_SsoPCKHeader_TR] pck on pck.PCLHOID = pch.OID	"
            vnQuery += vbCrLf & "	left join [SB_WMS].dbo.[Sys_SsoDSPPick_TR] dsp on dsp.PCKHOID= pck.OID	"
            vnQuery += vbCrLf & "	left join [SB_WMS].dbo.[Sys_SsoDSPHeader_TR] dsh on dsh.OID= dsp.DSPHOID	"
            vnQuery += vbCrLf & "	left join dbo.Sys_DcmDriver_MA drv on drv.OID= dsh.DcmSchDriverOID	"
            vnQuery += vbCrLf & "	left join Sys_DcmScheduleDetail_TR skd on skd.NotaNo=mj.NO_NOTA and skd.SchDTypeOID=1	"
            vnQuery += vbCrLf & "	left join Sys_DcmScheduleHeader_TR skh on skh.OID=skd.DcmSchHOID	"
            vnQuery += vbCrLf & "	LEFT JOIN SB_WMS.dbo.Sys_SsoTransStatus_MA sstsm ON pch.TransCode = sstsm.TransCode AND pch.TransStatus = sstsm.TransStatus	"
            vnQuery += vbCrLf & "Where 1=1 and LEFT(NO_NOTA,1) <> 'P' "
            If IsDate(vriTxtListStart) Then
                vnQuery += vbCrLf & "            and mj.TANGGAL >= '" & vriTxtListStart & "'"
            End If
            If IsDate(vriTxtListEnd) Then
                vnQuery += vbCrLf & "            and mj.TANGGAL <= '" & vriTxtListEnd & "'"
            End If
            If Val(vriDstWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & " and mj.WarehouseOID = " & vriDstWarehouse.SelectedValue & " "
            End If
            If Val(vriDstCompany.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and mj.CompanyCode = '" & vriDstCompany.SelectedValue & "'"
            End If
            If Trim(vriInvoiceNo) <> "" Then
                vnQuery += vbCrLf & " and mj.NO_NOTA like '%" & fbuFormatString(Trim(vriInvoiceNo)) & "%'"
            End If
            If Trim(vriNoPicklist) <> "" Then
                vnQuery += vbCrLf & " and pch.PCLNo like '%" & fbuFormatString(Trim(vriNoPicklist)) & "%'"
            End If
            If Trim(vriNoReferensi) <> "" Then
                vnQuery += vbCrLf & " and pch.PCLRefHNo like '%" & fbuFormatString(Trim(vriNoReferensi)) & "%'"
            End If
            If Trim(vriNoPicking) <> "" Then
                vnQuery += vbCrLf & " and pck.PCKNo like '%" & fbuFormatString(Trim(vriNoPicking)) & "%'"
            End If
            If Trim(vriNoDispatch) <> "" Then
                vnQuery += vbCrLf & " and dsh.DSPNo like '%" & fbuFormatString(Trim(vriNoDispatch)) & "%'"
            End If

            If Chk_Upload.Checked = True Then
                vnQuery += vbCrLf & " and mj.UploadDatetime is not null "

            Else
                vnQuery += vbCrLf & " and mj.UploadDatetime is null "
            End If
            If Chk_Picklist.Checked = True Then
                vnQuery += vbCrLf & " and pch.creationdatetime is not null "

            Else
                vnQuery += vbCrLf & " and pch.creationdatetime is null "
            End If

            If Chk_PickilistPrepared.Checked = True Then
                vnQuery += vbCrLf & " and pch.PreparedDatetime is not null "

            Else
                vnQuery += vbCrLf & " and pch.PreparedDatetime is null "
            End If
            If Chk_Picking.Checked = True Then
                vnQuery += vbCrLf & " and pck.CreationDatetime is not null "

            Else
                vnQuery += vbCrLf & " and pck.CreationDatetime is null "
            End If
            If Chk_PickingDone.Checked = True Then
                vnQuery += vbCrLf & " and pck.PickDoneDatetime is not null "

            Else
                vnQuery += vbCrLf & " and pck.PickDoneDatetime is null "
            End If

            If Chk_Dispatch.Checked = True Then
                vnQuery += vbCrLf & " and dsh.CreationDatetime is not null "

            Else
                vnQuery += vbCrLf & " and dsh.CreationDatetime is null "
            End If
            If Chk_DispatchDone.Checked = True Then
                vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is not null "

            Else
                vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is null "
            End If
            If Chk_DriverConfirm.Checked = True Then
                vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is not null "

            Else
                vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is null "
            End If
            vnQuery += vbCrLf & " ORDER BY mj.TANGGAL DESC  "


            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "ProgressNota-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT MONITORING INVOICE"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriDstCompany.SelectedValue

            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Monitoring Invoice"
            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriTxtListStart & " s/d " & vriTxtListEnd



            vnXCol = 7
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vriTxtListStart, 1, Len(vriTxtListEnd) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "TransCode"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Start to End"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Invoice No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Invoice Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Upload Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. PickList"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal PickList"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Schedule Date"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Referensi"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No ReF id"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Prepared Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Upload to Create Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picking Created Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Pick Done Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Picking Done to Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Dispatch Created Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Confirm Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Back Date Time"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransStatusDescr")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Start_to_End")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("NO_NOTA")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TANGGAL")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("KODE_CUST")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("UploadDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLScheduleDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PreparedDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Upload_to_Create_Picklist")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCKNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Picking_Created_Date_Time")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PickDoneDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DSPNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DSPDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Picking_Done_to_Dispatch")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Dispatch_Created_Date_Time")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DriverConfirmDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DCMDriverName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BackDatetime")

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

    Public Sub pbuCreateXlsx_MonStock(ByRef vriFileName As String, vriCompanyCode As String, vriWarehouseOID As Integer, vriTxtStart As String, vriSQLConn As SqlConnection)

        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SOTally", vriWarehouseOID, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCrNotaNo As String = ""
            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")

            vnQuery = "	DECLARE @CompanyCode AS VARCHAR(100)	"
            vnQuery += vbCrLf & "	SET @CompanyCode = '" & vriCompanyCode & "'	"
            vnQuery += vbCrLf & "	DECLARE @Warehouse AS INT	"
            vnQuery += vbCrLf & "	SET @Warehouse = " & vriWarehouseOID & "	"
            vnQuery += vbCrLf & "	DECLARE @tanggal AS DATETIME	"
            vnQuery += vbCrLf & "	SET @tanggal = '" & vriTxtStart & "'	"
            vnQuery += vbCrLf & "	Select tb.WarehouseName, tb.CompanyCode,tb.vRcvPODate, tb.BRGCODE, tb.BRGNAME,	"
            vnQuery += vbCrLf & "	SUM(tb.QtyOnHand)vOnHand,	"
            vnQuery += vbCrLf & "	((SUM(tb.QtyOnPicking) + SUM(TB.QtyOnDispatch) )) vStockOnPicking,	"
            vnQuery += vbCrLf & "	((SUM(tb.QtyOnHand))-(SUM(tb.QtyOnPicking) + SUM(TB.QtyOnDispatch) )) vSelisih	"
            vnQuery += vbCrLf & "	From(	"
            vnQuery += vbCrLf & "	Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,	"
            vnQuery += vbCrLf & "	     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,	"
            vnQuery += vbCrLf & "	     case when pm.StorageStagIO=0 then ''	"
            vnQuery += vbCrLf & "	          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,	"
            vnQuery += vbCrLf & "	     pm.vStorageOID,isnull(sm.OID,0)vStockCardOID,	"
            vnQuery += vbCrLf & "	     pm.vStorageInfoHtml,mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,	"
            vnQuery += vbCrLf & "	     sm.TransCode, pch.TransStatus,tm.TransName,sm.TransOID,	"
            vnQuery += vbCrLf & "	     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,	"
            vnQuery += vbCrLf & "	     sm.TransQty,	"
            vnQuery += vbCrLf & "	      ss.QtyOnHand,	"
            vnQuery += vbCrLf & "	  ss.QtyOnKarantina,	"
            vnQuery += vbCrLf & "	  (ss.QtyOnPutaway + ss.QtyOnPutawayWh + ss.QtyOnPutawayWh + ss.QtyOnPutawayKr 	"
            vnQuery += vbCrLf & "	  + ss.QtyOnPutawayDtw + ss.QtyOnPutawayDty + ss.QtyOnPutawayPtv 	"
            vnQuery += vbCrLf & "	  + ss.QtyOnPutawayDsw + ss.QtyOnPutawayDsy	"
            vnQuery += vbCrLf & "	  ) vQtyOnPutaway,	"
            vnQuery += vbCrLf & "	  (ss.QtyOnMovement + ss.QtyOnMovementWh + ss.QtyOnSgo) vQtyMovement,	"
            vnQuery += vbCrLf & "	  ss.QtyOnPickList,	"
            vnQuery += vbCrLf & "	  ss.QtyOnPicking,	"
            vnQuery += vbCrLf & "	  ss.QtyOnDispatch	"
            vnQuery += vbCrLf & "	 From SB_DATAWH.dbo.fnTbl_SsoStorageInfo('')pm	"
            vnQuery += vbCrLf & "	      inner join Sys_SsoStockCard_TR sm with(nolock) on sm.StorageOID=pm.vStorageOID	"
            vnQuery += vbCrLf & "	      inner join Sys_SsoRcvPOHeader_TR rc with(nolock) on rc.OID=sm.RcvPOHOID	"
            vnQuery += vbCrLf & "	      inner join Sys_SsoTransName_MA tm with(nolock) on tm.TransCode=sm.TransCode	"
            vnQuery += vbCrLf & "	      inner join SB_DATAWH.dbo.Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE	"
            vnQuery += vbCrLf & "	  INNER JOIN Sys_SsoStorageStock_MA ss with(nolock) on ss.StorageOID=pm.vStorageOID AND ss.BRGCODE = sm.BRGCODE AND ss.CompanyCode = mb.CompanyCode AND ss.RcvPOHOID = rc.OID	"
            vnQuery += vbCrLf & "	  INNER JOIN Sys_SsoPCLHeader_TR PCH with(nolock) ON PCH.PCLRefHOID = SM.RcvPOHOID	"
            vnQuery += vbCrLf & "	Where 1=1 	"
            vnQuery += vbCrLf & "	            and mb.CompanyCode= @CompanyCode	"
            vnQuery += vbCrLf & "	AND pm.WarehouseOID = @Warehouse	"
            vnQuery += vbCrLf & "	AND rc.RcvPODate <= @tanggal	"

            vnQuery += vbCrLf & "	)tb 	"
            vnQuery += vbCrLf & "	GROUP BY tb.WarehouseName, tb.CompanyCode, tb.BRGCODE, tb.BRGNAME, tb.QtyOnHand,tb.vRcvPODate	"
            vnQuery += vbCrLf & "	Order by tb.vRcvPODate, tb.WarehouseName, tb.CompanyCode, tb.BRGCODE DESC	"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = vriFileName & "-OID-" & vriWarehouseOID & "-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            Dim vnSONo As String = ""
            Dim vnCompName As String = ""
            Dim vnWarehouseName As String = ""
            Dim vnSubWhsName As String = ""
            Dim vnSOID As String = ""
            Dim vnSOCutOff As String = ""
            Dim vnStatus As String = ""
            Dim vnSONote As String = ""
            Dim vnSOCloseNote As String = ""

            If vnDtb.Rows.Count = 0 Then
                vnDRow = vnDtb.Rows(0)
            Else
                vnDRow = vnDtb.Rows(0)

                vnCompName = vnDRow.Item("CompanyCode")
                vnWarehouseName = vnDRow.Item("WarehouseName")
                vnSubWhsName = vnDRow.Item("vSubWhsName")

            End If

            '<---------------ROW 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMBER BERKAT"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 2
            vnXRow = vnXRow + 1
            vnXCol = 1


            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STOCK - MONITORING STOCK - ALL DATA"



            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            '<---------------ROW 4
            'NO SO
            vnXRow = vnXRow + 2
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "NO SO"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            'Company
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            'SOID
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnSOID

            '<---------------ROW 5
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            'Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnWarehouseName

            'Status
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            '<---------------ROW 6
            'Cutoff
            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            'Sub Warehouse
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sub Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            'CloseNote
            vnXCol = vnXCol + 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Close Note"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = ""

            '---------------------------------------------------------------
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnColCount As Byte = 10
            Dim vnRowIdxHead As Byte = vnXRow + 2
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Penerimaan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Brg Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Brg Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "vOnHand"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "vStockOnPicking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "vSelisih"

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
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WareHouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRcvPODate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vOnHand")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStockOnPicking")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSelisih")
                Next

                For vnRow = vnRowIdxHead + 1 To vnDtb.Rows.Count - 1
                    For vnXCol = 5 To 7
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).Style.NumberFormat.Format = "#,###"
                        vnIXLWorksheet.Row(vnRow).Cell(vnXCol).DataType = XLCellValues.Number
                    Next
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
    Public Sub pbuCreateXlsx_StockCard(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList,
                                       vriHdfListRcvPOHOID As String, vriTxtListBrgCode As TextBox, ChkStorageOID As CheckBox, vriTxtStorageOID As TextBox,
                                       vriRdbListStagging As RadioButtonList, vriDstListBuilding As DropDownList, vriDstListLantai As DropDownList,
                                       vriDstListZona As DropDownList, vriDstListStorageType As DropDownList,
                                       vriTxtListRackN_Start As TextBox, vriTxtListRackY_SeqNo As TextBox, vriTxtListRackY_Column As TextBox, vriTxtListRackY_Level As TextBox,
                                       vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_StockCard", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnCrBrgCode As String = fbuFormatString(Trim(vriTxtListBrgCode.Text))
            Dim vnCrRcvPOHOID As String = vriHdfListRcvPOHOID

            Dim vnQuery As String

            Dim vnStorageOID As String

            If ChkStorageOID.Checked Then
                vnStorageOID = vriTxtStorageOID.Text
            Else
                vnQuery = "Select vStorageOID From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & HttpContext.Current.Session("UserID") & "')pm"
                vnQuery += vbCrLf & "inner join Sys_SsoStockCard_TR sm with(nolock) on sm.StorageOID=pm.vStorageOID and sm.BRGCODE ='" & vnCrBrgCode & "'"
                vnQuery += vbCrLf & "Where pm.WarehouseOID=" & vriWarehouse.SelectedValue
                vnQuery += vbCrLf & "      and pm.BuildingOID=" & vriDstListBuilding.SelectedValue
                vnQuery += vbCrLf & "      and pm.LantaiOID=" & vriDstListLantai.SelectedValue
                vnQuery += vbCrLf & "      and pm.ZonaOID=" & vriDstListZona.SelectedValue
                vnQuery += vbCrLf & "      and pm.StorageTypeOID=" & vriDstListStorageType.SelectedValue

                If vriDstListStorageType.SelectedValue = enuStorageType.Floor Then
                    vnQuery += vbCrLf & "      and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(vriTxtListRackN_Start.Text)) & "'"
                ElseIf vriDstListStorageType.SelectedValue = enuStorageType.Rack Then
                    vnQuery += vbCrLf & "      and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(vriTxtListRackY_SeqNo.Text)) & "'"
                    vnQuery += vbCrLf & "      and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(vriTxtListRackY_Column.Text)) & "'"
                    vnQuery += vbCrLf & "      and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(vriTxtListRackY_Level.Text)) & "'"
                ElseIf vriDstListStorageType.SelectedValue = enuStorageType.Staging Then
                    vnQuery += vbCrLf & "      and pm.StorageStagIO=" & vriRdbListStagging.SelectedValue
                End If
                vnStorageOID = fbuGetDataNumSQL(vnQuery, vriSQLConn)
            End If

            Dim vnFromCriteria As String

            vnFromCriteria = " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & HttpContext.Current.Session("UserID") & "')pm"
            vnFromCriteria += vbCrLf & "      inner join Sys_SsoStockCard_TR sm with(nolock) on sm.StorageOID=pm.vStorageOID"
            vnFromCriteria += vbCrLf & "      inner join Sys_SsoRcvPOHeader_TR rc with(nolock) on rc.OID=sm.RcvPOHOID"
            vnFromCriteria += vbCrLf & "      inner join Sys_SsoTransName_MA tm with(nolock) on tm.TransCode=sm.TransCode"
            vnFromCriteria += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
            vnFromCriteria += vbCrLf & "Where 1=1"

            vnFromCriteria += vbCrLf & "            and mb.CompanyCode='" & vriCompany.SelectedValue & "'"
            vnFromCriteria += vbCrLf & "            and mb.BRGCODE='" & vnCrBrgCode & "'"
            vnFromCriteria += vbCrLf & "            and sm.RcvPOHOID=" & vnCrRcvPOHOID
            vnFromCriteria += vbCrLf & "            and pm.vStorageOID=" & vnStorageOID

            Dim vnDtb As New DataTable
            vnQuery = "Select * From("
            vnQuery += vbCrLf & "Select pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,"
            vnQuery += vbCrLf & "     pm.StorageTypeName,pm.vIsRack,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,pm.StorageNumber,"
            vnQuery += vbCrLf & "     case when pm.StorageStagIO=0 then ''"
            vnQuery += vbCrLf & "          when pm.StorageStagIO=1 then 'IN' else 'OUT' end vStorageStagIO,"
            vnQuery += vbCrLf & "     pm.vStorageOID,isnull(sm.OID,0)vStockCardOID,"
            vnQuery += vbCrLf & "     pm.vStorageInfoHtml,mb.CompanyCode,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
            vnQuery += vbCrLf & "     sm.TransCode,tm.TransName,sm.TransOID,"
            vnQuery += vbCrLf & "     convert(varchar(11),sm.CreationDatetime,106)+' '+convert(varchar(8),sm.CreationDatetime,108)vCreationDatetime,"
            vnQuery += vbCrLf & "     sm.TransQty,"
            vnQuery += vbCrLf & "     0 vQtyOnHand"
            vnQuery += vbCrLf & vnFromCriteria

            vnQuery += vbCrLf & "UNION"

            vnQuery += vbCrLf & "Select ''WarehouseName,''BuildingName,''LantaiDescription,''ZonaName,"
            vnQuery += vbCrLf & "     ''StorageTypeName,''vIsRack,''StorageSequenceNumber,''StorageColumn,''StorageLevel,''StorageNumber,"
            vnQuery += vbCrLf & "     ''vStorageStagIO,"
            vnQuery += vbCrLf & "     Null vStorageOID,Null vStockCardOID,"
            vnQuery += vbCrLf & "     ''vStorageInfoHtml,''CompanyCode,'',''vRcvPODate,''BRGCODE,''BRGNAME,"
            vnQuery += vbCrLf & "     ''TransCode,'TOTAL'TransName,Null TransOID,"
            vnQuery += vbCrLf & "     ''vCreationDatetime,"
            vnQuery += vbCrLf & "     sum(sm.TransQty)TransQty,"
            vnQuery += vbCrLf & "     dbo.fnSsoGet_StorageStock_QtyOnHand_ByKey('" & vnStorageOID & "','" & vriCompany.SelectedValue & "','" & vnCrBrgCode & "'," & vnCrRcvPOHOID & ")vQtyOnHand"
            vnQuery += vbCrLf & vnFromCriteria

            vnQuery += vbCrLf & ")tb Order by case when isnull(vStockCardOID,0)=0 then 19 else case when TransCode='" & stuTransCode.SsoStockOB & "' then 4 else 5 end end,vStockCardOID"

            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Stock Card-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT STOCK CARD"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Penerimaan "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriHdfListRcvPOHOID & ""

            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriTxtListBrgCode.Text & ""

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Building"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Lantai"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Zona"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Typa"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Rack"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sequence Number"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Column"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Level"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Number"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Stage"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Stock Card OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Info"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Penerimaan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Penerimaan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Trans Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Trans Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Trans ID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Create Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity On Hand"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BuildingName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("LantaiDescription")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("ZonaName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageTypeName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vIsRack")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageSequenceNumber")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageColumn")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfoHtml")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageLevel")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageNumber")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageStagIO")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfoHtml")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPONo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRcvPODate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vCreationDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQtyOnHand")




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
    Public Sub pbuCreateXlsx_MonitoringPickList1(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList,
                                        vriPCLNo As TextBox, vriRefNo As TextBox, vriPCKNo As TextBox, vriDSPNo As TextBox, vriStartDate As TextBox, vriEndDate As TextBox, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringPickList1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable

            vnQuery = " SELECT * FROM (  Select pcl.OID vPCLHOID,pcl.PCLNo,pcl.PCLCompanyCode,msc.SchDTypeName,pcl.PCLRefHNo,"
            vnQuery += vbCrLf & "		        pcl.WarehouseOID,mwh.WarehouseName,"
            vnQuery += vbCrLf & "		        convert(varchar(11),pcl.CreationDatetime,106)+'<br />'+convert(varchar(11),pcl.CreationDatetime,108)vCreationDatetime,"
            vnQuery += vbCrLf & "		        convert(varchar(11),pcl.PreparedDatetime,106)+'<br />'+convert(varchar(11),pcl.PreparedDatetime,108)vPreparedDatetime,"
            vnQuery += vbCrLf & "		        sts_pcl.TransStatusDescr vTransStatusDescr_PCL,"
            vnQuery += vbCrLf & "		        pck.PCKNo,convert(varchar(11),pck.PCKDate,106)vPCKDate,pck.StorageOID,sti.vStorageInfoHtml,sts_pck.TransStatusDescr vTransStatusDescr_PCK,"
            vnQuery += vbCrLf & "		        dsp.DSPNo,convert(varchar(11),dsp.DSPDate,106)vDSPDate,mdr.DcmDriverName,mvh.VehicleNo,dsp.vTransStatusDescr_DSP,dsp.vCreateUser_DSP,"
            vnQuery += vbCrLf & "		        dsr.DSRNo,convert(varchar(11),dsr.DSRDate,106)vDSRDate,dsr.vTransStatusDescr_DSR,dsr.vCreateUser_DSR,"
            vnQuery += vbCrLf & "		        sgo.SGONo,convert(varchar(11),sgo.SGODate,106)vSGODate,sgo.vTransStatusDescr_SGO,sgo.vCreateUser_SGO,sgo_asal.vStorageInfo_Wh_Bd_Lt vStgOut_Asal,sgo_dest.vStorageInfo_Wh_Bd_Lt vStgOut_Dest,"
            vnQuery += vbCrLf & "		        pcl.CreationDatetime"
            vnQuery += vbCrLf & "		 From Sys_SsoPCLHeader_TR pcl with(nolock)"
            vnQuery += vbCrLf & "	          inner join " & vnDBDcm & "Sys_DcmSchDType_MA msc with(nolock) on msc.OID=pcl.SchDTypeOID"
            vnQuery += vbCrLf & "	          inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=pcl.WarehouseOID"
            vnQuery += vbCrLf & "	          inner join Sys_SsoTransStatus_MA sts_pcl with(nolock) on sts_pcl.TransCode=pcl.TransCode and sts_pcl.TransStatus=pcl.TransStatus"
            vnQuery += vbCrLf & "	          left outer join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=pcl.OID"
            vnQuery += vbCrLf & "	          left outer join Sys_SsoTransStatus_MA sts_pck with(nolock) on sts_pck.TransCode=pck.TransCode and sts_pck.TransStatus=pck.TransStatus"

            vnQuery += vbCrLf & "	          left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sti on sti.vStorageOID=pck.StorageOID"

            vnQuery += vbCrLf & "	          left outer join fnTbl_SsoDSPHeader_Pick() dsp on dsp.PCKHOID=pck.OID"
            vnQuery += vbCrLf & "	          left outer join fnTbl_SsoDSRHeader_Pick() dsr on dsr.PCKHOID=pck.OID"

            vnQuery += vbCrLf & "	          left outer join fnTbl_SsoSGOHeader_Pick() sgo on sgo.PCKHOID=pck.OID"
            vnQuery += vbCrLf & "	          left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sgo_asal on sgo_asal.vStorageOID=sgo.StorageOID"
            vnQuery += vbCrLf & "	          left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sgo_dest on sgo_dest.vStorageOID=sgo.StorageOID_Dest"

            vnQuery += vbCrLf & "	          left outer join " & vnDBDcm & "Sys_DcmDriver_MA mdr with(nolock) on mdr.OID=dsp.DcmSchDriverOID"
            vnQuery += vbCrLf & "	          left outer join " & vnDBDcm & "Sys_DcmVehicle_MA mvh with(nolock) on mvh.OID=dsp.DcmVehicleOID"
            vnQuery += vbCrLf & "	    Where pcl.SchDTypeOID in (" & enuSchDType.Invoice & "," & enuSchDType.TRB & "," & enuSchDType.Perintah_Kirim_DO_Titip & ")"
            vnQuery += vbCrLf & "	    UNION ALL"
            vnQuery += vbCrLf & "	    Select pcl.OID vPCLHOID,pcl.PCLNo,pcl.PCLCompanyCode,msc.SchDTypeName,pcl.PCLRefHNo,"
            vnQuery += vbCrLf & "	           pcl.WarehouseOID,mwh.WarehouseName,"
            vnQuery += vbCrLf & "		       convert(varchar(11),pcl.CreationDatetime,106)+'<br />'+convert(varchar(11),pcl.CreationDatetime,108)vCreationDatetime,"
            vnQuery += vbCrLf & "		       convert(varchar(11),pcl.PreparedDatetime,106)+'<br />'+convert(varchar(11),pcl.PreparedDatetime,108)vPreparedDatetime,"
            vnQuery += vbCrLf & "	           sts_pcl.TransStatusDescr vTransStatusDescr_PCL,"
            vnQuery += vbCrLf & "	           pck.PCKNo,convert(varchar(11),pck.PCKDate,106)vPCKDate,pck.StorageOID,sti.vStorageInfoHtml,sts_pck.TransStatusDescr vTransStatusDescr_PCK,"
            vnQuery += vbCrLf & "		       ptw.vPtwNo DSPNo,convert(varchar(11),ptw.vPtwDate,106)vDSPDate,''DcmDriverName,''VehicleNo,ptw.vTransStatusDescr_Ptw,ptw.vCreateUser_Ptw,"
            vnQuery += vbCrLf & "		       dsr.DSRNo,convert(varchar(11),dsr.DSRDate,106)vDSRDate,dsr.vTransStatusDescr_DSR,dsr.vCreateUser_DSR,"
            vnQuery += vbCrLf & "		       sgo.SGONo,convert(varchar(11),sgo.SGODate,106)vSGODate,sgo.vTransStatusDescr_SGO,sgo.vCreateUser_SGO,sgo_asal.vStorageInfo_Wh_Bd_Lt vStgOut_Asal,sgo_dest.vStorageInfo_Wh_Bd_Lt vStgOut_Dest,"
            vnQuery += vbCrLf & "		       pcl.CreationDatetime"
            vnQuery += vbCrLf & "	      From Sys_SsoPCLHeader_TR pcl with(nolock)"
            vnQuery += vbCrLf & "	           inner join " & vnDBDcm & "Sys_DcmSchDType_MA msc with(nolock) on msc.OID=pcl.SchDTypeOID"
            vnQuery += vbCrLf & "	           inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=pcl.WarehouseOID"
            vnQuery += vbCrLf & "	           inner join Sys_SsoTransStatus_MA sts_pcl with(nolock) on sts_pcl.TransCode=pcl.TransCode and sts_pcl.TransStatus=pcl.TransStatus"

            vnQuery += vbCrLf & "	           left outer join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID=pcl.OID"
            vnQuery += vbCrLf & "	           left outer join fnTbl_SsoDSRHeader_Pick() dsr on dsr.PCKHOID=pck.OID"
            vnQuery += vbCrLf & "	           left outer join Sys_SsoTransStatus_MA sts_pck with(nolock) on sts_pck.TransCode=pck.TransCode and sts_pck.TransStatus=pck.TransStatus"

            vnQuery += vbCrLf & "	           left outer join fnTbl_SsoSGOHeader_Pick() sgo on sgo.PCKHOID=pck.OID"
            vnQuery += vbCrLf & "	           left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sgo_asal on sgo_asal.vStorageOID=sgo.StorageOID"
            vnQuery += vbCrLf & "	           left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sgo_dest on sgo_dest.vStorageOID=sgo.StorageOID_Dest"

            vnQuery += vbCrLf & "	           left outer join " & vnDBMaster & "fnTbl_SsoStorageInfo(0) sti on sti.vStorageOID=pck.StorageOID"
            vnQuery += vbCrLf & "	           left outer join fnTbl_SsoDT_Ptw_Header() ptw on ptw.PCKHOID=pck.OID"
            vnQuery += vbCrLf & "		 Where pcl.SchDTypeOID = " & enuSchDType.DO_Titip & "  ) tb "
            vnQuery += vbCrLf & " Where 1 = 1 "

            If Val(vriWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "      And WarehouseOID =" & vriWarehouse.SelectedValue
            End If
            If Val(vriCompany.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "      And CompanyCode =" & vriCompany.SelectedValue
            End If
            If IsDate(vriStartDate.Text) Then
                vnQuery += vbCrLf & "            and vCreationDatetime >= '" & vriStartDate.Text & "'"
            End If
            If IsDate(vriEndDate.Text) Then
                vnQuery += vbCrLf & "            and vCreationDatetime <= '" & vriEndDate.Text & "'"
            End If

            If Trim(vriPCLNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and PCLNo like '%" & fbuFormatString(Trim(vriPCLNo.Text)) & "%'"
            End If
            If Trim(vriRefNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and PCLRefHNo like '%" & fbuFormatString(Trim(vriRefNo.Text)) & "%'"
            End If
            If Trim(vriPCKNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and isnull(PCKNo,'') like '%" & fbuFormatString(Trim(vriPCKNo.Text)) & "%'"
            End If
            If Trim(vriDSPNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and isnull(DSPNo,'') like '%" & fbuFormatString(Trim(vriDSPNo.Text)) & "%'"
            End If

            vnQuery += vbCrLf & " Order by CreationDatetime DESC"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Monitoring Picklist-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT MONITORING PICKLIST"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PERIODE "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriStartDate.Text & " and " & vriEndDate.Text

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Jenis"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Referensi"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Creation"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Prepared"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Location"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Vehicle No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Dispatch By"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Dispatch Receive"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Move Antar Staging"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status Move Antar Staging"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Staging Out Asal"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Staging Out Tujuan"
            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLCompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SchDTypeName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vCreationDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPreparedDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTransStatusDescr_PCL")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCKNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPCKDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("StorageOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfoHtml")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTransStatusDescr_PCK")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DSPNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDSRDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTransStatusDescr_DSR")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vCreateUser_DSR")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SGONo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSGODate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTransStatusDescr_SGO")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vCreateUser_SGO")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStgOut_Asal")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStgOut_Dest")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CreationDatetime")

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

    Public Sub pbuCreateXlsx_SummaryStock1(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList, vriStartDate As TextBox, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SummaryStock1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable
            Dim vnCutOfDate As String = fbuGetCutOfDate(vriCompany.SelectedValue, vriSQLConn)

            Dim vnParam As String
            vnParam = "'" & vriCompany.SelectedValue & "'," & vriWarehouse.SelectedValue & ",'" & vnCutOfDate & "','" & vriStartDate.Text & "'"

            vnQuery = "spSsoInvNotPickDone_ByCompanyWhsDate_Table " & vnParam
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Summary Stock-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT SUMMARY STOCK"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PERIODE "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriStartDate.Text & ""

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity Stock Card"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity Invoice Belum Picking Done"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity TRB Belum Picking Done"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Quantity Akhir"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("")
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vsTextStream.WriteLine(vnDRow.Item("BRGCODE"))

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQty_StockCard")
                    vsTextStream.WriteLine("vQty_StockCard = " & vnDRow.Item("vQty_StockCard"))

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQtyInv_Belum_PickingDone")
                    vsTextStream.WriteLine("vQtyInv_Belum_PickingDone = " & vnDRow.Item("vQtyInv_Belum_PickingDone"))

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQtyTRB_Belum_PickingDone")
                    vsTextStream.WriteLine("vQtyTRB_Belum_PickingDone = " & vnDRow.Item("vQtyTRB_Belum_PickingDone"))

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQty_StockCard_Winacc")
                    vsTextStream.WriteLine("vQty_StockCard_Winacc = " & vnDRow.Item("vQty_StockCard_Winacc"))
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
    Public Sub pbuCreateXlsx_SummaryPutaway(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList,
                                        vriBrgCode As TextBox, vriStartDate As TextBox, vriEndDate As TextBox, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SummarySKK", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable
            vnQuery = "Select ptw.TransCode,stn.TransName,ptw.vPtwCompanyCode,ptw.vPtwNo,"
            vnQuery += vbCrLf & "      convert(varchar(11),ptw.vPtwDate,106)vPtwDate,"
            vnQuery += vbCrLf & "      mwh.WarehouseName,mwh_d.WarehouseName vWarehouseName_Dest,"
            vnQuery += vbCrLf & "      ptw.PCKNo,ptw.PCLRefHNo,sts.TransStatusDescr,"
            vnQuery += vbCrLf & "      ptw.RcvPOHOID,ptw.RcvPONo,"
            vnQuery += vbCrLf & "      ptw.BRGCODE,msb.BRGNAME,ptw.vSumPtwScan1Qty,ptw.vPtwReceiveQty,ptw.vSumPtwScan2Qty,ptw.CreationDatetime"
            vnQuery += vbCrLf & "      From fnTbl_SsoPutaway_Summary('" & HttpContext.Current.Session("UserID") & "') ptw"
            vnQuery += vbCrLf & "			inner join Sys_SsoTransName_MA stn with(nolock) on stn.TransCode=ptw.TransCode"
            vnQuery += vbCrLf & "			inner join Sys_SsoTransStatus_MA sts with(nolock) on sts.TransCode=ptw.TransCode and sts.TransStatus=ptw.TransStatus"
            vnQuery += vbCrLf & "           inner join " & vnDBMaster & "Sys_MstBarang_MA msb with(nolock) on msb.BRGCODE=ptw.BRGCODE and msb.CompanyCode=ptw.vPtwCompanyCode"
            vnQuery += vbCrLf & "			inner join " & vnDBMaster & "Sys_Warehouse_MA mwh with(nolock) on mwh.OID=ptw.WarehouseOID"
            vnQuery += vbCrLf & "			left outer join " & vnDBMaster & "Sys_Warehouse_MA mwh_d with(nolock) on mwh_d.OID=ptw.WarehouseOID_Dest"
            vnQuery += vbCrLf & "Where 1=1"
            vnQuery += vbCrLf & "            and msb.CompanyCode='" & vriCompany.SelectedValue & "'"
            vnQuery += vbCrLf & "            and msb.BRGCODE like '%" & vriBrgCode.Text & "%' and msb.BRGNAME like '%" & vriBrgCode.Text & "%'"

            If Val(vriWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & "            and ptw.WarehouseOID=" & vriWarehouse.SelectedValue
            End If

            If IsDate(vriStartDate.Text) Then
                vnQuery += vbCrLf & "            and ptw.vPtwDate >= '" & vriStartDate.Text & "'"
            End If
            If IsDate(vriEndDate.Text) Then
                vnQuery += vbCrLf & "            and ptw.vPtwDate <= '" & vriEndDate.Text & "'"
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

            vnFNm = "Summary Putaway-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT MONITORING PICKLIST"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PERIODE "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriStartDate.Text & " and " & vriEndDate.Text

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "TransCode"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Transaksi Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "vPtwNo"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Gudang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Gudang Tujuan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No.Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No.Invoice"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "RcvPOHOID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No.Penerimaan"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan 1"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Diterima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Scan 2"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Creation Date time"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwCompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vWarehouseName_Dest")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCKNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransStatusDescr")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPOHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumPtwScan1Qty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPtwReceiveQty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vSumPtwScan2Qty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CreationDatetime")


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
    Public Sub pbuCreateXlsx_MonDoTitip(ByRef vriFileName As String, vriUserOID As String, vriCompany As DropDownList,
                                        vriChkVarianOnly As CheckBox, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SummarySKK", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")
            Dim vnUserOID As String = HttpContext.Current.Session("UserOID")
            Dim vnQuery As String
            Dim vnDtb As New DataTable
            vnQuery = "Select sm.vCompanyCode,sm.vKodeBarang,mb.BRGNAME,sm.vTotal_QtySisaInvoice,sm.vTotal_QtySisaStock,sm.vTotal_QtySelisih"
            vnQuery += vbCrLf & "From ("
            vnQuery += vbCrLf & "Select case when isnull(tb1.CompanyCode,'')='' then tb2.CompanyCode else tb1.CompanyCode end vCompanyCode,"
            vnQuery += vbCrLf & "       case when isnull(tb1.KodeBarang,'')='' then tb2.BRGCODE else tb1.KodeBarang end vKodeBarang,"
            vnQuery += vbCrLf & "	    isnull(tb1.vTotal_QtySisaInvoice,0)vTotal_QtySisaInvoice,"
            vnQuery += vbCrLf & "	    isnull(tb2.vTotal_QtySisaStock,0) vTotal_QtySisaStock,"
            vnQuery += vbCrLf & "	    isnull(tb1.vTotal_QtySisaInvoice,0) - isnull(tb2.vTotal_QtySisaStock,0) vTotal_QtySelisih"
            vnQuery += vbCrLf & "  From " & vnDBDcm & "fnTbl_DcmDOTitip_SisaInvoice() tb1"
            vnQuery += vbCrLf & "       full join fnTbl_SsoDOTitip_SisaStorageStock()tb2 on tb2.CompanyCode=tb1.CompanyCode and tb2.BRGCODE=tb1.KodeBarang"
            vnQuery += vbCrLf & ")sm"
            vnQuery += vbCrLf & "       inner join " & vnDBMaster & "Sys_MstBarang_MA mb with(nolock) on mb.CompanyCode=sm.vCompanyCode and mb.BRGCODE=sm.vKodeBarang"

            If vnUserCompanyCode = "" Then
            Else
                vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=sm.vCompanyCode and uc.UserOID=" & vnUserOID
            End If

            vnQuery += vbCrLf & " Where 1=1"
            vnQuery += vbCrLf & "            and sm.vCompanyCode='" & vriCompany.SelectedValue & "'"

            If vriChkVarianOnly.Checked Then
                vnQuery += vbCrLf & "       and isnull(sm.vTotal_QtySisaInvoice,0) - isnull(sm.vTotal_QtySisaStock,0)<>0"
            End If

            vnQuery += vbCrLf & "Order by sm.vCompanyCode,sm.vKodeBarang"
            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Monitoring_DO_Titip-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT MONITORING PICKLIST"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            If vriChkVarianOnly.Checked Then
                vnXCol = 5
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tampilkan HANYA Data Selisih "

            End If



            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sisa Invoice"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Sisa Stock DO Titip"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Selisih"


            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vCompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vKodeBarang")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotal_QtySisaInvoice")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotal_QtySisaStock")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotal_QtySelisih")



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

    Public Sub pbuCreateXlsx_MonInvoice(ByRef vriFileName As String, vriUserOID As String, vriCompany As DropDownList, vriWarehouse As DropDownList,
                                        vriInvoiceNo As TextBox, vriPCLNo As TextBox, vriRefNo As TextBox, vriPCKNo As TextBox, vriDSPNo As TextBox, vriStartDate As TextBox, vriEndDate As TextBox,
                                        Chk_Upload As CheckBox, Chk_Picklist As CheckBox, Chk_PickilistPrepared As CheckBox, Chk_Picking As CheckBox,
                                        Chk_PickingDone As CheckBox, vriChkVarianOnly As CheckBox, Chk_Dispatch As CheckBox, Chk_DispatchDone As CheckBox,
                                        Chk_DriverConfirm As CheckBox,
                                        vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_SummarySKK", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")
            Dim vnUserOID As String = HttpContext.Current.Session("UserOID")
            Dim vnQuery As String
            Dim vnDtb As New DataTable



            Dim Transsum As Integer = 0
            Dim vnCompany As String = fbuFormatString(Trim(vriCompany.SelectedValue))

            vnQuery = "	select distinct	"
            vnQuery += vbCrLf & "	mj.CompanyCode,whs.WarehouseName, whs.OID, DATEDIFF(HOUR,mj.UploadDatetime,skh.BackDatetime) as [Durasi_Start_to_End]	"
            vnQuery += vbCrLf & "	, mj.NO_NOTA, mj.TANGGAL, mj.KODE_CUST, mj.CUSTOMER, mj.UploadDatetime	"
            vnQuery += vbCrLf & "	, pch.PCLNo, pch.PCLDate, pch.PCLScheduleDate, pch.CreationDatetime as [Time_Create_Picklist], usr_pch.UserName, pch.PreparedDatetime	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pch.CreationDatetime) as [Durasi_Upload_to_Create_Picklist]	"
            vnQuery += vbCrLf & "	, pck.PCKNo, pck.PCKDate, pck.CreationDatetime 'Picking_Created_Date_Time', pck.PickDoneDatetime, pch.PCLRefHOID, pch.PCLRefHNo	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pch.CreationDatetime,pck.PickDoneDatetime) as [Durasi_Picklist_Created_to_Picking_Done]	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,mj.UploadDatetime,pck.PickDoneDatetime) as [Durasi_Upload_to_Picking_Done]	"
            vnQuery += vbCrLf & "	, dsh.DSPNo, dsh.DSPDate, dsh.CreationDatetime 'Dispatch_Created_Date_Time', dsh.DispatchDoneDatetime 'Dispatch_Created_Date'	"
            vnQuery += vbCrLf & "	, dsh.DriverConfirmDatetime 	"
            vnQuery += vbCrLf & "	, DATEDIFF(MINUTE,pck.PickDoneDatetime,dsh.DriverConfirmDatetime) as [Durasi_Picking_Done_to_Dispatch]	"
            vnQuery += vbCrLf & "	,drv.DcmDriverName,skh.BackDatetime, dsh.CancelledDatetime, pch.TransCode , pch.TransStatus, sstsm.TransStatusDescr	"
            vnQuery += vbCrLf & "	from 	"

            vnQuery += vbCrLf & "	(select	"
            vnQuery += vbCrLf & "	ju.CompanyCode, ju.WarehouseOID, ju.NO_NOTA, ju.TANGGAL, ju.KODE_CUST, ju.CUSTOMER, max(ju.UploadDatetime) as uploadDatetime	"
            vnQuery += vbCrLf & "	from " & vnDBDcm & "Sys_DcmJUAL	ju with(nolock)"

            If vnUserCompanyCode = "" Then
            Else
                vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=ju.CompanyCode and uc.UserOID=" & HttpContext.Current.Session("UserOID")
            End If

            vnQuery += vbCrLf & "	group by CompanyCode, WarehouseOID, NO_NOTA, TANGGAL, KODE_CUST, CUSTOMER	"
            vnQuery += vbCrLf & "	) as mj	"

            vnQuery += vbCrLf & "	left join Sys_SsoPCLHeader_TR pch with(nolock) on pch.PCLRefHNo = mj.NO_NOTA	"
            vnQuery += vbCrLf & "	left join Sys_SsoUser_MA usr_pch with(nolock) on usr_pch.OID = pch.CreationUserOID	"
            vnQuery += vbCrLf & "	left join " & vnDBMaster & "Sys_Warehouse_MA whs with(nolock) on whs.OID = mj.WarehouseOID	"
            vnQuery += vbCrLf & "	left join Sys_SsoPCKHeader_TR pck with(nolock) on pck.PCLHOID = pch.OID	"
            vnQuery += vbCrLf & "	left join Sys_SsoDSPPick_TR dsp with(nolock) on dsp.PCKHOID= pck.OID	"
            vnQuery += vbCrLf & "	left join Sys_SsoDSPHeader_TR dsh with(nolock) on dsh.OID= dsp.DSPHOID	"
            vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmDriver_MA drv with(nolock) on drv.OID= dsh.DcmSchDriverOID	"
            vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmScheduleDetail_TR skd with(nolock) on skd.NotaNo=mj.NO_NOTA and skd.SchDTypeOID=1	"
            vnQuery += vbCrLf & "	left join " & vnDBDcm & "Sys_DcmScheduleHeader_TR skh with(nolock) on skh.OID=skd.DcmSchHOID	"
            vnQuery += vbCrLf & "	LEFT JOIN Sys_SsoTransStatus_MA sstsm with(nolock) ON pch.TransCode = sstsm.TransCode AND pch.TransStatus = sstsm.TransStatus	"
            vnQuery += vbCrLf & "Where 1=1 and LEFT(NO_NOTA,1) <> 'P'"
            If IsDate(vriStartDate.Text) Then
                vnQuery += vbCrLf & "            and mj.TANGGAL >= '" & vriStartDate.Text & "'"
            End If
            If IsDate(vriEndDate.Text) Then
                vnQuery += vbCrLf & "            and mj.TANGGAL <= '" & vriEndDate.Text & "'"
            End If
            If Val(vriWarehouse.SelectedValue) > 0 Then
                vnQuery += vbCrLf & " and mj.WarehouseOID = " & vriWarehouse.SelectedValue & " "
            End If
            If Val(vriWarehouse.SelectedValue) > 0 Then

                vnQuery += vbCrLf & "            and mj.CompanyCode = '" & vnCompany & "'"
            End If
            If Trim(vriInvoiceNo.Text) <> "" Then
                vnQuery += vbCrLf & " and mj.NO_NOTA like '%" & fbuFormatString(Trim(vriInvoiceNo.Text)) & "%'"
            End If
            If Trim(vriPCLNo.Text) <> "" Then
                vnQuery += vbCrLf & " and pch.PCLNo like '%" & fbuFormatString(Trim(vriPCLNo.Text)) & "%'"
            End If
            If Trim(vriRefNo.Text) <> "" Then
                vnQuery += vbCrLf & " and pch.PCLRefHNo like '%" & fbuFormatString(Trim(vriRefNo.Text)) & "%'"
            End If
            If Trim(vriPCKNo.Text) <> "" Then
                vnQuery += vbCrLf & " and pck.PCKNo like '%" & fbuFormatString(Trim(vriPCKNo.Text)) & "%'"
            End If
            If Trim(vriDSPNo.Text) <> "" Then
                vnQuery += vbCrLf & " and dsh.DSPNo like '%" & fbuFormatString(Trim(vriDSPNo.Text)) & "%'"
            End If

            If Chk_Upload.Checked = True Then
                vnQuery += vbCrLf & " and mj.UploadDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and mj.UploadDatetime is null "
            End If
            If Chk_Picklist.Checked = True Then
                vnQuery += vbCrLf & " and pch.creationdatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and pch.creationdatetime is null "
            End If

            If Chk_PickilistPrepared.Checked = True Then
                vnQuery += vbCrLf & " and pch.PreparedDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and pch.PreparedDatetime is null "
            End If
            If Chk_Picking.Checked = True Then
                vnQuery += vbCrLf & " and pck.CreationDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and pck.CreationDatetime is null "
            End If
            If Chk_PickingDone.Checked = True Then
                vnQuery += vbCrLf & " and pck.PickDoneDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and pck.PickDoneDatetime is null "
            End If

            If Chk_Dispatch.Checked = True Then
                vnQuery += vbCrLf & " and dsh.CreationDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and dsh.CreationDatetime is null "
            End If
            If Chk_DispatchDone.Checked = True Then
                vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and dsh.DispatchDoneDatetime is null "
            End If
            If Chk_DriverConfirm.Checked = True Then
                vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is not null "
                Transsum = Transsum + 1
            Else
                vnQuery += vbCrLf & " and dsh.DriverConfirmDatetime is null "
            End If
            vnQuery += vbCrLf & " ORDER BY mj.TANGGAL DESC  "


            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Monitoring_Invoice-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT MONITORING STATUS INVOICE"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            If vriChkVarianOnly.Checked Then
                vnXCol = 5
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tampilkan HANYA Data Selisih "

            End If



            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "TransCode"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Start to End"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Invoice No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Invoice Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer Code"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Upload Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. PickList"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal PickList"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Schedule Date"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No Referensi"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No ReF id"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Prepared Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SDurasi Upload to Create Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picking Created Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Pick Done Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Start to End"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "TNo. Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Durasi Picking Done to Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Dispatch Created Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Confirm Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Back Date Time"


            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TransStatusDescr")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Start_to_End")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("NO_NOTA")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TANGGAL")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("KODE_CUST")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CUSTOMER")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("UploadDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLScheduleDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PCLRefHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PreparedDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Upload_to_Create_Picklist")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Picking_Created_Date_Time")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PickDoneDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DSPNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DSPDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Durasi_Picking_Done_to_Dispatch")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Dispatch_Created_Date_Time")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DriverConfirmDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("DCMDriverName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BackDatetime")



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

    Public Sub pbuCreateXlsx_StockInfo1(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList, vriListBuilding As DropDownList,
                                        vriListLantai As DropDownList, vriListZona As DropDownList, vriListStorageType As DropDownList, RdbListStagging As RadioButtonList,
                                        vriChkStorageOID As CheckBox,
                                        vriTxtListRcvNo As TextBox,
                                        vriTxtListRackY_SeqNo As TextBox,
                                        vriTxtListRackY_Column As TextBox,
                                        vriTxtListRackY_Level As TextBox,
                                        vriTxtListRackN_Start As TextBox,
                                        vriTxtListRackN_End As TextBox,
                                        vriTxtListBrgCode As TextBox, vriTxtListBrgName As TextBox,
                                        vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringPickList1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable


            Dim vnCrBrgCode As String = fbuFormatString(Trim(vriTxtListBrgCode.Text))
            Dim vnCrBrgName As String = fbuFormatString(Trim(vriTxtListBrgName.Text))


            vnQuery = "Select "
            vnQuery += vbCrLf & "     isnull(sm.OID,0)vStorageStockOID,pm.vStorageOID,pm.vStorageInfoHtml,"
            vnQuery += vbCrLf & "     mb.CompanyCode,sm.RcvPOHOID,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
            vnQuery += vbCrLf & "     sm.QtyOnHand,"
            vnQuery += vbCrLf & "     sm.vQtyAvailable,"
            vnQuery += vbCrLf & "     sm.QtyOnPutaway,sm.QtyOnPutawayWh,sm.QtyOnMovement,sm.QtyOnMovementWh,"
            vnQuery += vbCrLf & "     sm.QtyOnPickList,sm.QtyOnPicking,sm.QtyOnSgo,sm.QtyOnDispatch,sm.QtyOnKarantina,sm.QtyOnPutawayKr,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayDtw,sm.QtyOnPutawayDty,sm.QtyOnPutawayPtv,sm.QtyOnPutawayDsw,sm.QtyOnPutawayDsy"
            vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & HttpContext.Current.Session("UserID") & "') pm"
            vnQuery += vbCrLf & "      inner join fnTbl_SsoStorageStock() sm on sm.StorageOID=pm.vStorageOID"
            vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"
            vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
            vnQuery += vbCrLf & "Where 1=1"

            vnQuery += vbCrLf & "            and mb.CompanyCode='" & vriCompany.SelectedValue & "'"
            vnQuery += vbCrLf & "            and mb.BRGCODE like '%" & vnCrBrgCode & "%' and mb.BRGNAME like '%" & vnCrBrgName & "%'"
            If Trim(vriTxtListRcvNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and rc.RcvPONo like '%" & Trim(vriTxtListRcvNo.Text) & "%'"
            End If

            If vriChkStorageOID.Checked Then
                vnQuery += vbCrLf & "            and pm.vStorageOID=" & Val(vriChkStorageOID.Text)
            Else
                If Val(vriWarehouse.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.WarehouseOID=" & vriWarehouse.SelectedValue
                End If
                If Val(vriListBuilding.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.BuildingOID=" & vriListBuilding.SelectedValue
                End If
                If Val(vriListLantai.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.LantaiOID=" & vriListLantai.SelectedValue
                End If
                If Val(vriListZona.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.ZonaOID=" & vriListZona.SelectedValue
                End If
                If Val(vriListStorageType.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.StorageTypeOID=" & vriListStorageType.SelectedValue
                End If

                If vriListStorageType.SelectedValue = enuStorageType.Rack Then
                    If Trim(vriTxtListRackY_SeqNo.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(vriTxtListRackY_SeqNo.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackY_Column.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(vriTxtListRackY_Column.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackY_Level.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(vriTxtListRackY_Level.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,mb.BRGCODE"

                ElseIf vriListStorageType.SelectedValue = enuStorageType.Floor Then
                    If Trim(vriTxtListRackN_Start.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(vriTxtListRackN_Start.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackN_End.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')<='" & fbuFormatString(Trim(vriTxtListRackN_End.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageNumber"

                ElseIf vriListStorageType.SelectedValue = enuStorageType.Staging Then
                    vnQuery += vbCrLf & "            and pm.StorageStagIO=" & RdbListStagging.SelectedValue
                    vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
                Else
                    vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
                End If
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Stock Info-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT STOCK INFO"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "DATA "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriTxtListBrgCode.Text & " and " & vriTxtListBrgName.Text

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Stock OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Location"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "OID Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Nama Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Hand"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Available"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Putaway<br />Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Movement"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Movement<br />Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Moving Antar Stg Out"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Karantina"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Karantina On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty DO Titip On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty DO Titip On Putaway Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Pick Void On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Penerimaan Dispatch On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Penerimaan Dispatch On Putaway Antar Wh"
            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageStockOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfoHtml")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPOHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPONo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRcvPODate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnHand")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQtyAvailable")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutaway")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayWh")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnMovement")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnMovementWh")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPickList")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPicking")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnSgo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnDispatch")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnKarantina")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayKr")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDtw")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayPtv")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDsw")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDsy")

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

    Public Sub pbuCreateXlsx_StockInfo1_20231006(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList, vriListBuilding As DropDownList,
                                        vriListLantai As DropDownList, vriListZona As DropDownList, vriListStorageType As DropDownList, RdbListStagging As RadioButtonList,
                                        vriChkStorageOID As CheckBox,
                                        vriTxtListRcvNo As TextBox,
                                        vriTxtListRackY_SeqNo As TextBox,
                                        vriTxtListRackY_Column As TextBox,
                                        vriTxtListRackY_Level As TextBox,
                                        vriTxtListRackN_Start As TextBox,
                                        vriTxtListRackN_End As TextBox,
                                        vriTxtListBrgCode As TextBox, vriTxtListBrgName As TextBox,
                                        vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringPickList1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable


            Dim vnCrBrgCode As String = fbuFormatString(Trim(vriTxtListBrgCode.Text))
            Dim vnCrBrgName As String = fbuFormatString(Trim(vriTxtListBrgName.Text))


            vnQuery = "Select "
            vnQuery += vbCrLf & "     isnull(sm.OID,0)vStorageStockOID,pm.vStorageOID,pm.vStorageInfoHtml,"
            vnQuery += vbCrLf & "     mb.CompanyCode,sm.RcvPOHOID,rc.RcvPONo,convert(varchar(11),rc.RcvPODate,106)vRcvPODate,mb.BRGCODE,mb.BRGNAME,"
            vnQuery += vbCrLf & "     sm.QtyOnHand,"
            vnQuery += vbCrLf & "     sm.vQtyAvailable,"
            vnQuery += vbCrLf & "     sm.QtyOnPutaway,sm.QtyOnPutawayWh,sm.QtyOnMovement,sm.QtyOnMovementWh,"
            vnQuery += vbCrLf & "     sm.QtyOnPickList,sm.QtyOnPicking,sm.QtyOnSgo,sm.QtyOnDispatch,sm.QtyOnKarantina,sm.QtyOnPutawayKr,"
            vnQuery += vbCrLf & "     sm.QtyOnPutawayDtw,sm.QtyOnPutawayDty,sm.QtyOnPutawayPtv,sm.QtyOnPutawayDsw,sm.QtyOnPutawayDsy"
            vnQuery += vbCrLf & " From " & vnDBMaster & "fnTbl_SsoStorageInfo('" & HttpContext.Current.Session("UserID") & "') pm"
            vnQuery += vbCrLf & "      inner join fnTbl_SsoStorageStock() sm on sm.StorageOID=pm.vStorageOID"
            vnQuery += vbCrLf & "      left outer join Sys_SsoRcvPOHeader_TR rc on rc.OID=sm.RcvPOHOID"
            vnQuery += vbCrLf & "      inner join " & vnDBMaster & "Sys_MstBarang_MA mb on mb.CompanyCode=sm.CompanyCode and mb.BRGCODE=sm.BRGCODE"
            vnQuery += vbCrLf & "Where 1=1"

            vnQuery += vbCrLf & "            and mb.CompanyCode='" & vriCompany.SelectedValue & "'"
            vnQuery += vbCrLf & "            and mb.BRGCODE like '%" & vnCrBrgCode & "%' and mb.BRGNAME like '%" & vnCrBrgName & "%'"
            If Trim(vriTxtListRcvNo.Text) <> "" Then
                vnQuery += vbCrLf & "            and rc.RcvPONo like '%" & Trim(vriTxtListRcvNo.Text) & "%'"
            End If

            If vriChkStorageOID.Checked Then
                vnQuery += vbCrLf & "            and pm.vStorageOID=" & Val(vriChkStorageOID.Text)
            Else
                If Val(vriWarehouse.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.WarehouseOID=" & vriWarehouse.SelectedValue
                End If
                If Val(vriListBuilding.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.BuildingOID=" & vriListBuilding.SelectedValue
                End If
                If Val(vriListLantai.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.LantaiOID=" & vriListLantai.SelectedValue
                End If
                If Val(vriListZona.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.ZonaOID=" & vriListZona.SelectedValue
                End If
                If Val(vriListStorageType.SelectedValue) > 0 Then
                    vnQuery += vbCrLf & "            and pm.StorageTypeOID=" & vriListStorageType.SelectedValue
                End If

                If vriListStorageType.SelectedValue = enuStorageType.Rack Then
                    If Trim(vriTxtListRackY_SeqNo.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageSequenceNumber,'')='" & fbuFormatString(Trim(vriTxtListRackY_SeqNo.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackY_Column.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageColumn,'')='" & fbuFormatString(Trim(vriTxtListRackY_Column.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackY_Level.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageLevel,'')='" & fbuFormatString(Trim(vriTxtListRackY_Level.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageSequenceNumber,pm.StorageColumn,pm.StorageLevel,mb.BRGCODE"

                ElseIf vriListStorageType.SelectedValue = enuStorageType.Floor Then
                    If Trim(vriTxtListRackN_Start.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')>='" & fbuFormatString(Trim(vriTxtListRackN_Start.Text)) & "'"
                    End If
                    If Trim(vriTxtListRackN_End.Text) <> "" Then
                        vnQuery += vbCrLf & "            and isnull(pm.StorageNumber,'')<='" & fbuFormatString(Trim(vriTxtListRackN_End.Text)) & "'"
                    End If
                    vnQuery += vbCrLf & " Order by pm.WarehouseName,pm.BuildingName,pm.LantaiDescription,pm.ZonaName,pm.StorageNumber"

                ElseIf vriListStorageType.SelectedValue = enuStorageType.Staging Then
                    vnQuery += vbCrLf & "            and pm.StorageStagIO=" & RdbListStagging.SelectedValue
                    vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
                Else
                    vnQuery += vbCrLf & " Order by rc.RcvPODate,rc.RcvPONo,mb.BRGCODE"
                End If
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Stock Info-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT STOCK INFO"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriCompany.SelectedValue

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriWarehouse.SelectedItem


            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "DATA "
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vriTxtListBrgCode.Text & " and " & vriTxtListBrgName.Text

            'Dim vnCrCB_Header As String = ""
            'For vn = 0 To vriChlCaraBayar.Items.Count - 1
            '    vnCrCB_Header += vriChlCaraBayar.Items(vn).Text & "=" & IIf(vriChlCaraBayar.Items(vn).Selected = True, "Y", "N") & ", "
            'Next

            vnXCol = 7
            'vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = Mid(vnCrCB_Header, 1, Len(vnCrCB_Header) - 2)
            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Stock OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage OID"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Storage Location"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Company"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "OID Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Terima"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Barang"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Hand"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Available"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Putaway<br />Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Movement"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Movement<br />Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Moving Antar Stg Out"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty On Karantina"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Karantina On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty DO Titip On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty DO Titip On Putaway Antar Wh"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Pick Void On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Penerimaan Dispatch On Putaway"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Qty Penerimaan Dispatch On Putaway Antar Wh"
            vnXCol = vnXCol + 1

            For vnXCol = 1 To 9
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
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageStockOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vStorageInfoHtml")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPOHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("RcvPONo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRcvPODate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGCODE")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("BRGNAME")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnHand")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vQtyAvailable")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutaway")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayWh")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnMovement")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnMovementWh")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPickList")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPicking")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnSgo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnDispatch")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnKarantina")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayKr")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDtw")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDty")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayPtv")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDsw")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("QtyOnPutawayDsy")

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
    Public Sub pbuCreateXlsx_OrderStatus(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As DropDownList, vriCompany As DropDownList,
                                        vriRdlListPickType As RadioButtonList, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringPickList1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable


            Dim vnCompanyCode As String = vriCompany.SelectedValue
            Dim vnWarehouse As String = vriWarehouse.SelectedValue

            Dim vnWarehouseName As String = fbuGetWarehouseName(vnWarehouse, vriSQLConn)

            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")
            Dim vnUserWarehouseCode As String = HttpContext.Current.Session("UserWarehouseCode")

            Dim vnCrDate As String = ""

            Dim vnCrStatus As String = ""

            vnQuery = "	SELECT PM.[CompanyCode],PM.[WarehouseName],[Order Status] vOrderStatus,[Ref No] vRefNo	"
            vnQuery += vbCrLf & "	  ,[TANGGAL],[Priority] ,[DO Titip] vDoTitip	"
            vnQuery += vbCrLf & "	  ,[KODE_CUST],[CUSTOMER],[uploadDatetime]	"
            vnQuery += vbCrLf & "	  ,[Picklist No] vPicklistNo,[Picklist Date] vPickListDate,[PL Created by]vPLCreate,[PreparedDatetime],[PL Status]	"
            vnQuery += vbCrLf & "	  ,[Picking No],[Picking Created Date] vPickingCreate,[Picking Done] vPickingDone	"
            vnQuery += vbCrLf & "	  ,[Dispatch/Putaway No] vDispatchNo,[Confirm Dispatch/Putaway Date] vDispatchConfirm,[Driver Confirm Date] vDriverConfirm	"
            vnQuery += vbCrLf & "	  ,[Driver Name] vDriverName,[Driver Return Time] vDriverReturn	"

            vnQuery += vbCrLf & " FROM " & fbuGetDBDcm() & "vOrderTracing PM"
            If vnUserCompanyCode <> "" Then
                vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & vriUserOID
            End If
            If vnUserWarehouseCode <> "" Then
                vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA mw with(nolock) on mw.WarehouseName=PM.WarehouseName"
                vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=mw.OID and uw.UserOID=" & vriUserOID
            End If

            vnQuery += vbCrLf & "Where 1=1"

            'vnQuery += vbCrLf & vnCrStatus

            If vriCompany.SelectedIndex > 0 Then
                vnQuery += vbCrLf & "            and PM.CompanyCode = '" & vnCompanyCode & "'"
            End If
            If vriWarehouse.SelectedIndex > 0 Then
                vnQuery += vbCrLf & "            and PM.WarehouseName = '" & vnWarehouseName & "'"
            End If
            If vriRdlListPickType.SelectedValue = 1 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') in ('Baru','Prepared','On Picking'))"

            ElseIf vriRdlListPickType.SelectedValue = 2 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') != 'Cancelled' and isnull([PL Status],'') != 'Void') and ([Picking Created Date] is not null) and ([Picking Done] is not null) and ([Confirm Dispatch/Putaway Date] is NULL) and ([Driver Return Time] is null)  "

            ElseIf vriRdlListPickType.SelectedValue = 3 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and isnull([Picklist No],'')=''"
            Else
                vnQuery += vbCrLf & ""
            End If
            vnQuery += vbCrLf & "Order by PM.TANGGAL Desc"

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Order Status-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT ORDER STATUS"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = IIf(vriCompany.SelectedValue = "", "ALL", vriCompany.SelectedValue)

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = IIf(vriWarehouse.SelectedValue = "0", "ALL", vriWarehouse.SelectedItem.Text)

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "JENIS REPORT"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = UCase(vriRdlListPickType.SelectedItem.Text)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 5
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "CompanyCode"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Order Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Ref No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Priority"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Do Titip"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Upload"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Create"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PL Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Prepared Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picking Create"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picking Done"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Dispatch Confirm"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Confirm"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Return"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 23
                vnIXLWorksheet.Row(vnRowIdxHead).Cell(vnXCol).Style.Fill.BackgroundColor = XLColor.AshGrey

            Next

            vnXRow = vnRowIdxHead
            vsTextStream.WriteLine("Proses : Mengisi Data...")
            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("TIDAK ADA DATA")
                vnXRow = vnXRow + 1
                vnXCol = 1
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol + 1).Value = "TIDAK ADA DATA"
            Else
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vOrderStatus")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRefNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TANGGAL")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Priority")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDoTitip")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("KODE_CUST")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CUSTOMER")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("uploadDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPicklistNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickListDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPLCreate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PreparedDateTime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PL Status")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Picking No")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickingCreate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickingDone")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDispatchNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDispatchConfirm")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverConfirm")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverReturn")


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
    Public Sub pbuCreateXlsx_OrderStatus2(ByRef vriFileName As String, vriUserOID As String, vriWarehouse As CheckBoxList, vriCompany As CheckBoxList,
                                        vriRdlListPickType As RadioButtonList, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_MonitoringPickList1", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnCrCB As String = ""

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnDBDcm As String = fbuGetDBDcm()
            Dim vnQuery As String
            Dim vnDtb As New DataTable

            Dim vnCompanyCode As String = vriCompany.SelectedValue
            Dim vnWarehouse As String = vriWarehouse.SelectedValue

            Dim vnUserCompanyCode As String = HttpContext.Current.Session("UserCompanyCode")
            Dim vnUserWarehouseCode As String = HttpContext.Current.Session("UserWarehouseCode")

            Dim vnCrDate As String = ""

            Dim vnCrStatus As String = ""
            Dim vnCrWarehouse As String = "''"
            Dim vnCrWhAll As Boolean = False
            Dim vnTwarehouse As String = ""

            For vn = 0 To vriWarehouse.Items.Count - 1
                If vriWarehouse.Items(vn).Selected Then
                    If vriWarehouse.Items(vn).Value = 0 Then
                        vnCrWhAll = True
                    Else
                        vnCrWarehouse += ",'" & fbuGetWarehouseName(vriWarehouse.Items(vn).Value, vriSQLConn) & "'"
                        vnTwarehouse += "" & fbuGetWarehouseName(vriWarehouse.Items(vn).Value, vriSQLConn) & ""
                    End If
                End If
            Next
            If vnCrWhAll = True Then
                vnCrWarehouse = ""
                vnTwarehouse = "ALL"
            Else
                If vnCrWarehouse = "''" Then
                    vnCrWarehouse = ""
                    vnTwarehouse = "ALL"
                Else
                    vnCrWarehouse = " and PM.WarehouseName  IN (" & vnCrWarehouse & ")"
                End If
            End If

            Dim j As Integer = 0
            Dim vnCrCompany As String = "''"
            Dim vnCrComAll As Boolean = False
            Dim vnTCompany As String = "''"

            For vn = 0 To vriCompany.Items.Count - 1
                If vriCompany.Items(vn).Selected Then
                    If vriCompany.Items(vn).Value = "" Then
                        vnCrComAll = True
                    Else
                        vnCrCompany += ",'" & vriCompany.Items(vn).Value & "'"
                        vnTCompany += "" & vriCompany.Items(vn).Value & ""
                    End If
                End If
            Next
            If vnCrComAll = True Then
                vnCrCompany = ""
            Else
                If vnCrCompany = "''" Then
                    vnCrCompany = ""
                    vnTCompany = "ALL"
                Else
                    vnCrCompany = " and PM.CompanyCode  IN (" & vnCrCompany & ")"
                End If
            End If

            vnQuery = "	SELECT PM.[CompanyCode],PM.[WarehouseName],[Order Status] vOrderStatus,[Ref No] vRefNo	"
            vnQuery += vbCrLf & "	  ,[TANGGAL],[Priority] ,[DO Titip] vDoTitip	"
            vnQuery += vbCrLf & "	  ,[KODE_CUST],[CUSTOMER],[uploadDatetime]	"
            vnQuery += vbCrLf & "	  ,[Picklist No] vPicklistNo,[Picklist Date] vPickListDate,[PL Created by]vPLCreate,[PreparedDatetime],[PL Status]	"
            vnQuery += vbCrLf & "	  ,[Picking No],[Picking Created Date] vPickingCreate,[Picking Done] vPickingDone	"
            vnQuery += vbCrLf & "	  ,[Dispatch/Putaway No] vDispatchNo,[Confirm Dispatch/Putaway Date] vDispatchConfirm,[Driver Confirm Date] vDriverConfirm	"
            vnQuery += vbCrLf & "	  ,[Driver Name] vDriverName,[Driver Return Time] vDriverReturn	"

            vnQuery += vbCrLf & " FROM " & fbuGetDBDcm() & "vOrderTracing PM"
            If vnUserCompanyCode <> "" Then
                vnQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=PM.CompanyCode and uc.UserOID=" & vriUserOID
            End If
            If vnUserWarehouseCode <> "" Then
                vnQuery += vbCrLf & "     inner join " & fbuGetDBMaster() & "Sys_Warehouse_MA mw with(nolock) on mw.WarehouseName=PM.WarehouseName"
                vnQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=mw.OID and uw.UserOID=" & vriUserOID
            End If

            vnQuery += vbCrLf & "Where 1=1"

            If vriRdlListPickType.SelectedValue = 1 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') in ('Baru','Prepared','On Picking'))"

            ElseIf vriRdlListPickType.SelectedValue = 2 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and (isnull([PL Status],'') != 'Cancelled' and isnull([PL Status],'') != 'Void') and ([Picking Created Date] is not null) and ([Picking Done] is not null) and ([Confirm Dispatch/Putaway Date] is NULL) and ([Driver Return Time] is null)  "

            ElseIf vriRdlListPickType.SelectedValue = 3 Then
                vnQuery += vbCrLf & "            and (isnull([Order Status],'') !='Cancelled') and isnull([Picklist No],'')=''"
            Else
                vnQuery += vbCrLf & ""
            End If
            vnQuery += vbCrLf & vnCrWarehouse & vnCrCompany & " Order by PM.TANGGAL Desc"

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("vnQuery")
            vsTextStream.WriteLine(vnQuery)
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Order Status-" & HttpContext.Current.Session("UserOID") & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "REPORT ORDER STATUS"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnTCompany

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnTwarehouse


            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "JENIS REPORT"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = UCase(vriRdlListPickType.SelectedItem.Text)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            Dim vnRowIdxHead As Byte = 5
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "CompanyCode"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Warehouse"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Order Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Ref No"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Priority"

            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Do Titip"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Kode Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Customer"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Upload"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picklist"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Picklist Create"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "PL Status"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Prepared Date Time"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Picking"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picking Create"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Tanggal Picking Done"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "No. Dispatch"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Dispatch Confirm"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Confirm"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Name"
            vnXCol = vnXCol + 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Driver Return"

            vnXCol = vnXCol + 1

            For vnXCol = 1 To 23
                vnIXLWorksheet.Row(vnRowIdxHead).Cell(vnXCol).Style.Fill.BackgroundColor = XLColor.AshGrey
            Next

            vnXRow = vnRowIdxHead
            vsTextStream.WriteLine("Proses : Mengisi Data...")
            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("TIDAK ADA DATA")
                vnXRow = vnXRow + 1
                vnXCol = 1
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol + 1).Value = "TIDAK ADA DATA"
            Else
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vOrderStatus")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vRefNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("TANGGAL")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Priority")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDoTitip")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("KODE_CUST")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("CUSTOMER")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("uploadDatetime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPicklistNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickListDate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPLCreate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PreparedDateTime")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("PL Status")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("Picking No")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickingCreate")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPickingDone")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDispatchNo")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDispatchConfirm")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverConfirm")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vDriverReturn")
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
    Public Sub pbuCreateXlsx_SOStatus(ByRef vriFileName As String, vriUserOID As String, vriUserCompanyCode As String, vriUserWarehouseCode As String, vriChkSt_Closed As CheckBox, vriChkSt_NotClosed As CheckBox, vriSQLConn As SqlConnection)
        Try
            pbuCreateLogFile(vsFso, vsTextStream, HttpContext.Current.Session("UserNip"), csModuleName, "pbuCreateXlsx_StockCard", 0, vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)

            Dim vnDBMaster As String = fbuGetDBMaster()
            Dim vnQuery As String

            vnQuery = "Select so.*,wh.WarehouseName,sw.SubWhsCode,sw.SubWhsName"
            vnQuery += vbCrLf & "       From fnTbl_SsoSOStatus('" & vriUserOID & "')so"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_SubWarehouse_MA sw with(nolock) on sw.OID=so.SOSubWarehouseOID"
            vnQuery += vbCrLf & "			 inner join " & vnDBMaster & "Sys_Warehouse_MA wh with(nolock) on wh.OID=so.SOWarehouseOID"
            vnQuery += vbCrLf & "       Where 1=1"

            If vriUserCompanyCode = "" Then
            Else
                vbuCrpQuery += vbCrLf & "     inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=so.SOCompanyCode and uc.UserOID=" & vriUserOID
            End If
            If vriUserWarehouseCode = "" Then
            Else
                vbuCrpQuery += vbCrLf & "     inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=so.SOWarehouseOID and uw.UserOID=" & vriUserOID
            End If

            If vriChkSt_Closed.Checked = True And vriChkSt_NotClosed.Checked = False Then
                vnQuery += vbCrLf & "             and so.TransStatus=" & enuTCSSOH.Closed
            ElseIf vriChkSt_Closed.Checked = False And vriChkSt_NotClosed.Checked = True Then
                vnQuery += vbCrLf & "             and so.TransStatus!=" & enuTCSSOH.Closed
            End If

            vnQuery += vbCrLf & " order by so.SOHOID"

            Dim vnDtb As New DataTable
            pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine(Date.Now)
            vsTextStream.WriteLine("Jumlah Data = " & vnDtb.Rows.Count)

            Dim vnFNm As String

            vnFNm = "Summary SO Status-" & vriUserOID & "-" & Format(Date.Now, "yyyyMMdd_HHmmss")
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
            Dim vnXCol As Integer = 1

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Header Report...")
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SUMMARY STATUS STOCK OPNAME"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.FontSize = "15"
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.Font.Bold = True

            vnXRow = vnXRow + 1
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "STATUS PER " & Format(Date.Now, "dd MMM yyyy HH:mm")

            Dim vnRowIdxHead As Byte = 4
            vnXRow = vnRowIdxHead
            vnXCol = 1
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO OID"
            vnXCol = 2
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "COMPANY"

            vnXCol = 3
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "WAREHOUSE"
            vnXCol = 4
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "SO NOTE"

            vnXCol = 5
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Item" & vbCrLf & "In System"
            vnXCol = 6
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Qty" & vbCrLf & "In System"

            vnXCol = 7
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Item" & vbCrLf & "Scanned"
            vnXCol = 8
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Percentage Item" & vbCrLf & "Scanned"

            vnXCol = 9
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Qty" & vbCrLf & "Scanned"
            vnXCol = 10
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Percentage Qty" & vbCrLf & "Scanned"

            vnXCol = 11
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Total Item" & vbCrLf & "Selisih"
            vnXCol = 12
            vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = "Percentage Item" & vbCrLf & "Selisih"

            '-----------------------------------------------------------------------------

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Proses : Mempersiapkan Column Header dan Column Format...")

            For vnXCol = 1 To 12
                vnIXLWorksheet.Row(vnRowIdxHead).Cell(vnXCol).Style.Fill.BackgroundColor = XLColor.LightGreen
            Next

            vnIXLWorksheet.Column(2).Width = "12"
            vnIXLWorksheet.Column(3).Width = "16"
            vnIXLWorksheet.Column(4).Width = "40"
            For vnXCol = 5 To 12
                vnIXLWorksheet.Column(2).Width = "12"
            Next

            vnXRow = vnRowIdxHead
            vsTextStream.WriteLine("Proses : Mengisi Data...")
            If vnDtb.Rows.Count = 0 Then
                vsTextStream.WriteLine("TIDAK ADA DATA")
                vnXRow = vnXRow + 1
                vnXCol = 1
                vnIXLWorksheet.Row(vnXRow).Cell(vnXCol + 1).Value = "TIDAK ADA DATA"

                vnIXLWorksheet.Range(vnRowIdxHead, 1, vnXRow, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                vnIXLWorksheet.Range(vnRowIdxHead, 1, vnXRow, 12).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin)
            Else
                Dim vnDRow As DataRow
                Dim vnRow As Integer
                For vnRow = 0 To vnDtb.Rows.Count - 1
                    vnDRow = vnDtb.Rows(vnRow)
                    vsTextStream.WriteLine("Row " & vnRow & " " & vnDtb.Rows(vnRow).Item(1))
                    vnXRow = vnXRow + 1
                    vnXCol = 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOHOID")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SOCompanyCode")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("WarehouseName")
                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("SONote")

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotalItem_System")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "#,##0"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotalQty_System")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "#,##0"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotalItem_Scanned")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "#,##0"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPercentageItem_Scanned_Xls")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "0.00%"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotalQty_Scanned")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "#,##0"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPercentageQty_Scanned_Xls")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "0.00%"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vTotalItem_Selisih")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "#,##0"

                    vnXCol = vnXCol + 1
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Value = vnDRow.Item("vPercentageItem_Selisih_Xls")
                    vnIXLWorksheet.Row(vnXRow).Cell(vnXCol).Style.NumberFormat.Format = "0.00%"
                Next

                vnIXLWorksheet.Range(vnRowIdxHead, 1, vnXRow, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                vnIXLWorksheet.Range(vnRowIdxHead, 1, vnXRow, 12).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin)
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

End Module
