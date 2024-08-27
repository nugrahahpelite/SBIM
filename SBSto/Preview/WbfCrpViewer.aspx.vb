Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Public Class WbfCrpViewer
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfCrpViewer"
    Const csTNoPrefix = "CAS"

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsLogFileNameOnly As String
    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If vbuPreviewOnClose <> "1" Then
            vbuCrpPreviewType = Request.QueryString("vqCrpPreviewType")
            If vbuCrpPreviewType = stuCrpPreviewType.ByQueryPopwin Then
                pbPrevieCrpWebQueryPopWin()
            ElseIf vbuCrpPreviewType = stuCrpPreviewType.ByDataTablePopwin Then
                pbPrevieCrpWebDtbPopWin()
            ElseIf vbuCrpPreviewType = stuCrpPreviewType.ByDataTable Then
                pbPrevieCrpWebDtb()
            Else
                pbPrevieCrpWebQuery()
            End If
        End If
    End Sub

    Sub pbPrevieCrpWebDtb()
        Try
            Dim vnRptFolder = Server.MapPath("~/CrpFiles/")
            Dim vnRpd As ReportDocument
            vnRpd = New ReportDocument
            vnRpd.Load(vnRptFolder & vbuCrpFileName)

            If vbuCrpQuery1 <> "" Then
                vnRpd.OpenSubreport(vbuCrpSubReport1)
                vnRpd.Subreports(vbuCrpSubReport1).SetDataSource(vbuCrpDtb1)

                If vbuCrpQuery2 <> "" Then
                    vnRpd.OpenSubreport(vbuCrpSubReport2)
                    vnRpd.Subreports(vbuCrpSubReport2).SetDataSource(vbuCrpDtb2)

                    If vbuCrpQuery3 <> "" Then
                        vnRpd.OpenSubreport(vbuCrpSubReport3)
                        vnRpd.Subreports(vbuCrpSubReport3).SetDataSource(vbuCrpDtb3)

                        If vbuCrpQuery4 <> "" Then
                            vnRpd.OpenSubreport(vbuCrpSubReport4)
                            vnRpd.Subreports(vbuCrpSubReport4).SetDataSource(vbuCrpDtb4)
                        End If
                    End If
                End If
            End If

            vnRpd.SetDataSource(vbuCrpDtb)

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            If vbuCrpShowGroupTree Then
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If Session("SessionPreview") = "Crp" Then
                Crv.ReportSource = vnRpd
            ElseIf Session("SessionPreview") = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, False, "")
            ElseIf Session("SessionPreview") = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, Response, False, "")
            End If

            vbuCrpDtb.Dispose()
            vbuCrpDtb1.Dispose()
            vbuCrpDtb2.Dispose()
            vbuCrpDtb3.Dispose()
            vbuCrpDtb4.Dispose()
        Catch ex As Exception
            pbMsgError = ex.Message
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True
            psCreateLogFile("pbPrevieCrpWebDtb", ex.Message)
        End Try
    End Sub

    Sub pbPrevieCrpWebQuery()
        Try
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            Dim vnRptFolder = Server.MapPath("~/CrpFiles/")
            Dim vnDtb As New DataTable
            Dim vnDtb1 As New DataTable
            Dim vnDtb2 As New DataTable
            Dim vnDtb3 As New DataTable
            Dim vnDtb4 As New DataTable

            pbuFillDtbSQL(vnDtb, vbuCrpQuery, vnSQLConn)

            Dim vnRpd As ReportDocument
            vnRpd = New ReportDocument
            vnRpd.Load(vnRptFolder & vbuCrpFileName)

            If vbuCrpQuery1 <> "" Then
                pbuFillDtbSQL(vnDtb1, vbuCrpQuery1, vnSQLConn)
                vnRpd.OpenSubreport(vbuCrpSubReport1)
                vnRpd.Subreports(vbuCrpSubReport1).SetDataSource(vnDtb1)

                If vbuCrpQuery2 <> "" Then
                    pbuFillDtbSQL(vnDtb2, vbuCrpQuery2, vnSQLConn)
                    vnRpd.OpenSubreport(vbuCrpSubReport2)
                    vnRpd.Subreports(vbuCrpSubReport2).SetDataSource(vnDtb2)

                    If vbuCrpQuery3 <> "" Then
                        pbuFillDtbSQL(vnDtb3, vbuCrpQuery3, vnSQLConn)
                        vnRpd.OpenSubreport(vbuCrpSubReport3)
                        vnRpd.Subreports(vbuCrpSubReport3).SetDataSource(vnDtb3)

                        If vbuCrpQuery4 <> "" Then
                            pbuFillDtbSQL(vnDtb4, vbuCrpQuery4, vnSQLConn)
                            vnRpd.OpenSubreport(vbuCrpSubReport4)
                            vnRpd.Subreports(vbuCrpSubReport4).SetDataSource(vnDtb4)
                        End If
                    End If
                End If
            End If

            vnRpd.SetDataSource(vnDtb)

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            If vbuCrpShowGroupTree Then
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If Session("SessionPreview") = "Crp" Then
                Crv.ReportSource = vnRpd
            ElseIf Session("SessionPreview") = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, False, "")
            ElseIf Session("SessionPreview") = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, Response, False, "")
            End If

            vnDtb.Reset()
            vnDtb1.Reset()
            vnDtb2.Reset()
            vnDtb3.Reset()
            vnDtb4.Reset()

            vnRpd.Close()
            vnRpd.Dispose()

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

        Catch ex As Exception
            pbMsgError = ex.Message
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True
            psCreateLogFile("pbPrevieCrpWebQuery", ex.Message)
        End Try
    End Sub

    Sub pbPrevieCrpWebQueryPopWin()
        vbuCrpFileName = Request.QueryString("vqCrpFileName")
        pbuCreateLogFile(vsFso, vsTextStream, Session("UserNip"), csModuleName, "pbPrevieCrpWebQueryPopWin", Mid(vbuCrpFileName, 1, Len(vbuCrpFileName) - 4), vsLogFileNameOnly, vsLogFileName, vsLogFileNameError)
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

        Dim vnRpd As ReportDocument
        vnRpd = New ReportDocument

        Try
            vbuCrpPreview = Request.QueryString("vqCrpPreview")
            vsTextStream.WriteLine("vbuCrpPreview " & vbuCrpPreview)

            vsTextStream.WriteLine("vbuCrpFileName " & vbuCrpFileName)

            'vbuCrpSubReport1 = Request.QueryString("vqCrpSubReport1")
            'vsTextStream.WriteLine("vbuCrpSubReport1 " & vbuCrpSubReport1)

            'vbuCrpSubReport2 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpSubReport2 " & vbuCrpSubReport2)

            'vbuCrpSubReport3 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpSubReport3 " & vbuCrpSubReport3)

            'vbuCrpSubReport4 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpSubReport4 " & vbuCrpSubReport4)

            'vbuCrpQuery = Request.QueryString("vqCrpQuery")
            'vsTextStream.WriteLine("vbuCrpQuery " & vbuCrpQuery)

            'vbuCrpQuery = Replace(vbuCrpQuery, "bnsrp", "%")
            'vsTextStream.WriteLine("vbuCrpQuery " & vbuCrpQuery)

            'vbuCrpQuery1 = Request.QueryString("vqCrpQuery1")
            'vsTextStream.WriteLine("vbuCrpQuery1 " & vbuCrpQuery1)

            'vbuCrpQuery1 = Replace(vbuCrpQuery1, "bnsrp", "%")
            'vsTextStream.WriteLine("vbuCrpQuery1 " & vbuCrpQuery1)

            'vbuCrpQuery2 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpQuery2 " & vbuCrpQuery2)

            'vbuCrpQuery3 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpQuery3 " & vbuCrpQuery3)

            'vbuCrpQuery4 = Request.QueryString("")
            'vsTextStream.WriteLine("vbuCrpQuery4 " & vbuCrpQuery4)

            Dim vnRptFolder = Server.MapPath("~/CrpFiles/")
            vsTextStream.WriteLine("vnRptFolder " & vnRptFolder)

            Dim vnDtb As New DataTable
            Dim vnDtb1 As New DataTable
            Dim vnDtb2 As New DataTable
            Dim vnDtb3 As New DataTable
            Dim vnDtb4 As New DataTable

            pbuFillDtbSQL(vnDtb, vbuCrpQuery, vnSQLConn)

            'Dim vnRpd As ReportDocument
            'vnRpd = New ReportDocument

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Load Report " & vnRptFolder & vbuCrpFileName & "...Start")
            vnRpd.Load(vnRptFolder & vbuCrpFileName)
            vsTextStream.WriteLine("Load Report " & vnRptFolder & vbuCrpFileName & "...End")
            vsTextStream.WriteLine("")

            If vbuCrpQuery1 <> "" Then
                vsTextStream.WriteLine("")
                vsTextStream.WriteLine("vbuCrpQuery1 <> Empty")

                pbuFillDtbSQL(vnDtb1, vbuCrpQuery1, vnSQLConn)
                vnRpd.OpenSubreport(vbuCrpSubReport1)
                vnRpd.Subreports(vbuCrpSubReport1).SetDataSource(vnDtb1)

                If vbuCrpQuery2 <> "" Then
                    vsTextStream.WriteLine("")
                    vsTextStream.WriteLine("vbuCrpQuery2 <> Empty")

                    pbuFillDtbSQL(vnDtb2, vbuCrpQuery2, vnSQLConn)
                    vnRpd.OpenSubreport(vbuCrpSubReport2)
                    vnRpd.Subreports(vbuCrpSubReport2).SetDataSource(vnDtb2)

                    If vbuCrpQuery3 <> "" Then
                        vsTextStream.WriteLine("")
                        vsTextStream.WriteLine("vbuCrpQuery3 <> Empty")

                        pbuFillDtbSQL(vnDtb3, vbuCrpQuery3, vnSQLConn)
                        vnRpd.OpenSubreport(vbuCrpSubReport3)
                        vnRpd.Subreports(vbuCrpSubReport3).SetDataSource(vnDtb3)

                        If vbuCrpQuery4 <> "" Then
                            vsTextStream.WriteLine("")
                            vsTextStream.WriteLine("vbuCrpQuery4 <> Empty")

                            pbuFillDtbSQL(vnDtb4, vbuCrpQuery4, vnSQLConn)
                            vnRpd.OpenSubreport(vbuCrpSubReport4)
                            vnRpd.Subreports(vbuCrpSubReport4).SetDataSource(vnDtb4)
                        End If
                    End If
                End If
            End If

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("SetDataSource...Start")
            vnRpd.SetDataSource(vnDtb)
            vsTextStream.WriteLine("SetDataSource...End")
            vsTextStream.WriteLine("")

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            If vbuCrpShowGroupTree Then
                vsTextStream.WriteLine("0.1")
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                vsTextStream.WriteLine("0.2")
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If

            vsTextStream.WriteLine("1")
            vnDtb.Reset()

            vsTextStream.WriteLine("2")
            vnDtb1.Reset()

            vsTextStream.WriteLine("3")
            vnDtb2.Reset()

            vsTextStream.WriteLine("4")
            vnDtb3.Reset()

            vsTextStream.WriteLine("5")
            vnDtb4.Reset()

            vsTextStream.WriteLine("6")

            pbuClearCrpSession()

            vsTextStream.WriteLine("7")

            If vbuCrpPreview = "Crp" Then
                vsTextStream.WriteLine("0.3")
                Crv.ReportSource = vnRpd
            ElseIf vbuCrpPreview = "Pdf" Then
                vsTextStream.WriteLine("0.4")
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, False, "")
            ElseIf vbuCrpPreview = "Xlsx" Then
                vsTextStream.WriteLine("0.5")
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, Response, False, "")
            End If

            vnRpd.Close()
            vsTextStream.WriteLine("8")

            vnRpd.Dispose()

            vsTextStream.WriteLine("9")

            vnRpd = Nothing

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            vsTextStream.WriteLine("")
            vsTextStream.WriteLine("Show Report Berhasil")
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

        Catch ex As Exception
            pbMsgError = ex.Message
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

            vnRpd.Close()
            vnRpd.Dispose()
            vnRpd = Nothing

            vsTextStream.WriteLine("ERROR")
            vsTextStream.WriteLine(ex.Message)
            vsTextStream.WriteLine("Process End           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
            vsTextStream.WriteLine("---------------EOF-------------------------")
            vsTextStream.Close()
            vsTextStream = Nothing
            vsFso = Nothing

            If pbMsgError Like "Load report failed*" Then
                FileCopy(vsLogFileName, vsLogFileNameError)
            End If
        End Try
    End Sub
    Sub pbPrevieCrpWebQueryPopWin_20221208_Bef_Log()
        Try
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            vbuCrpPreview = Request.QueryString("vqCrpPreview")
            vbuCrpFileName = Request.QueryString("vqCrpFileName")
            vbuCrpSubReport1 = Request.QueryString("vqCrpSubReport1")
            vbuCrpSubReport2 = Request.QueryString("")
            vbuCrpSubReport3 = Request.QueryString("")
            vbuCrpSubReport4 = Request.QueryString("")

            vbuCrpQuery = Request.QueryString("vqCrpQuery")
            vbuCrpQuery = Replace(vbuCrpQuery, "bnsrp", "%")

            vbuCrpQuery1 = Request.QueryString("vqCrpQuery1")
            vbuCrpQuery1 = Replace(vbuCrpQuery1, "bnsrp", "%")

            vbuCrpQuery2 = Request.QueryString("")
            vbuCrpQuery3 = Request.QueryString("")
            vbuCrpQuery4 = Request.QueryString("")

            Dim vnRptFolder = Server.MapPath("~/CrpFiles/")
            Dim vnDtb As New DataTable
            Dim vnDtb1 As New DataTable
            Dim vnDtb2 As New DataTable
            Dim vnDtb3 As New DataTable
            Dim vnDtb4 As New DataTable

            pbuFillDtbSQL(vnDtb, vbuCrpQuery, vnSQLConn)

            Dim vnRpd As ReportDocument
            vnRpd = New ReportDocument
            vnRpd.Load(vnRptFolder & vbuCrpFileName)

            If vbuCrpQuery1 <> "" Then
                pbuFillDtbSQL(vnDtb1, vbuCrpQuery1, vnSQLConn)
                vnRpd.OpenSubreport(vbuCrpSubReport1)
                vnRpd.Subreports(vbuCrpSubReport1).SetDataSource(vnDtb1)

                If vbuCrpQuery2 <> "" Then
                    pbuFillDtbSQL(vnDtb2, vbuCrpQuery2, vnSQLConn)
                    vnRpd.OpenSubreport(vbuCrpSubReport2)
                    vnRpd.Subreports(vbuCrpSubReport2).SetDataSource(vnDtb2)

                    If vbuCrpQuery3 <> "" Then
                        pbuFillDtbSQL(vnDtb3, vbuCrpQuery3, vnSQLConn)
                        vnRpd.OpenSubreport(vbuCrpSubReport3)
                        vnRpd.Subreports(vbuCrpSubReport3).SetDataSource(vnDtb3)

                        If vbuCrpQuery4 <> "" Then
                            pbuFillDtbSQL(vnDtb4, vbuCrpQuery4, vnSQLConn)
                            vnRpd.OpenSubreport(vbuCrpSubReport4)
                            vnRpd.Subreports(vbuCrpSubReport4).SetDataSource(vnDtb4)
                        End If
                    End If
                End If
            End If

            vnRpd.SetDataSource(vnDtb)

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            If vbuCrpShowGroupTree Then
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If vbuCrpPreview = "Crp" Then
                Crv.ReportSource = vnRpd
            ElseIf vbuCrpPreview = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, False, "")
            ElseIf vbuCrpPreview = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, Response, False, "")
            End If

            vnDtb.Reset()
            vnDtb1.Reset()
            vnDtb2.Reset()
            vnDtb3.Reset()
            vnDtb4.Reset()

            pbuClearCrpSession()

            vnRpd.Close()
            vnRpd.Dispose()

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

        Catch ex As Exception
            pbMsgError = ex.Message
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True
            psCreateLogFile("pbPrevieCrpWebQueryPopWin", ex.Message)
        End Try
    End Sub

    Sub pbPrevieCrpWebDtbPopWin()
        Try
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            vbuCrpPreview = Request.QueryString("vqCrpPreview")
            vbuCrpFileName = Request.QueryString("vqCrpFileName")

            Dim vnRptFolder = Server.MapPath("~/CrpFiles/")
            Dim vnDtb As New DataTable
            Dim vnDtb1 As New DataTable
            Dim vnDtb2 As New DataTable
            Dim vnDtb3 As New DataTable
            Dim vnDtb4 As New DataTable

            'If vbuCrpFileName = stuCrpName.CrpAbsenceDataEmpByOrg Then
            '    'pbuPayFillSalary(vnDtb, True, Request.QueryString("vqCrpPeriodOID"), Request.QueryString("vqCrpPeriodCompanyOID"), Request.QueryString("vqCrpPeriodCompany"), Request.QueryString("vqCrpBranchOID"), Request.QueryString("vqCrpEmpName"), Request.QueryString("vqCrpEmpNIP"), True, vnSQLConn)
            'End If

            Dim vnRpd As ReportDocument
            vnRpd = New ReportDocument
            vnRpd.Load(vnRptFolder & vbuCrpFileName)

            If Request.QueryString("vqCrpSubReportCount") > 0 Then
                vbuCrpSubReport1 = Request.QueryString("vqCrpSubReport1")

                pbuFillDtbSQL(vnDtb1, vbuCrpQuery1, vnSQLConn)
                vnRpd.OpenSubreport(vbuCrpSubReport1)
                vnRpd.Subreports(vbuCrpSubReport1).SetDataSource(vnDtb1)

                If vbuCrpQuery2 <> "" Then
                    pbuFillDtbSQL(vnDtb2, vbuCrpQuery2, vnSQLConn)
                    vnRpd.OpenSubreport(vbuCrpSubReport2)
                    vnRpd.Subreports(vbuCrpSubReport2).SetDataSource(vnDtb2)

                    If vbuCrpQuery3 <> "" Then
                        pbuFillDtbSQL(vnDtb3, vbuCrpQuery3, vnSQLConn)
                        vnRpd.OpenSubreport(vbuCrpSubReport3)
                        vnRpd.Subreports(vbuCrpSubReport3).SetDataSource(vnDtb3)

                        If vbuCrpQuery4 <> "" Then
                            pbuFillDtbSQL(vnDtb4, vbuCrpQuery4, vnSQLConn)
                            vnRpd.OpenSubreport(vbuCrpSubReport4)
                            vnRpd.Subreports(vbuCrpSubReport4).SetDataSource(vnDtb4)
                        End If
                    End If
                End If
            End If

            vnRpd.SetDataSource(vnDtb)

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            If vbuCrpShowGroupTree Then
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                Crv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If vbuCrpPreview = "Crp" Then
                Crv.ReportSource = vnRpd
            ElseIf vbuCrpPreview = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, False, "")
            ElseIf vbuCrpPreview = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, Response, False, "")
            End If

            vnDtb.Reset()
            vnDtb1.Reset()
            vnDtb2.Reset()
            vnDtb3.Reset()
            vnDtb4.Reset()

            pbuClearCrpSession()

            vnRpd.Close()
            vnRpd.Dispose()

            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing

        Catch ex As Exception
            pbMsgError = ex.Message
            LblMsgError.Text = ex.Message
            LblMsgError.Visible = True
            psCreateLogFile("pbPrevieCrpWebDtbPopWin", ex.Message)
        End Try
    End Sub

    Private Sub psCreateLogFile(vriProcessName As String, vriMessage As String)
        Dim vsTextStream As Scripting.TextStream
        Dim vsFso As Scripting.FileSystemObject

        Dim vsProcessDate As String
        Dim vsLogFolder As String
        Dim vsLogFileName As String

        vsLogFolder = Server.MapPath("~") & "\WebLog\"
        vsProcessDate = Format(Date.Now, "yyMMdd_HHmmss")

        vsLogFileName = vsLogFolder & "SBPay_Preview_Error_" & vsProcessDate & "_" & vriProcessName & ".log"

        vsTextStream = Nothing
        vsFso = CreateObject("Scripting.FileSystemObject")
        vsTextStream = vsFso.OpenTextFile(vsLogFileName, Scripting.IOMode.ForWriting, True)

        vsTextStream.WriteLine("SB - PREVIEW ERROR")
        vsTextStream.WriteLine("Process Start           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vsTextStream.WriteLine("Process Name            : " & vriProcessName)
        vsTextStream.WriteLine("Request.UserHostAddress : " & Request.UserHostAddress)
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine(vriMessage)

        vsTextStream.Close()
        vsTextStream = Nothing
        vsFso = Nothing
    End Sub

End Class