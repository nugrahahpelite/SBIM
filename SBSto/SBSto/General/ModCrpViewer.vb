Imports CrystalDecisions.CrystalReports.Engine
Imports System.Web.UI.Page
Imports System.Data.SqlClient

Module ModCrpViewer
    Public vbuCrpFileName As String
    Public vbuCrpSubReport1 As String
    Public vbuCrpSubReport2 As String
    Public vbuCrpSubReport3 As String
    Public vbuCrpSubReport4 As String

    Public vbuCrpQuery As String
    Public vbuCrpQuery1 As String
    Public vbuCrpQuery2 As String
    Public vbuCrpQuery3 As String
    Public vbuCrpQuery4 As String

    Public vbuPrevPage As String
    Public vbuCrpShowGroupTree As Boolean

    Public vbuCrpPreviewType As String
    Public vbuCrpPreview As String

    Public vbuPreviewOnClose As String

    Public Const cbuPopwinCrp = "'popwin', 'width=1450, height=850, left=15, top=50'"

    Public Structure stuCrpPreviewType
        Const ByDataTable = "DataTable"
        Const ByQuery = "Query"
        Const ByQueryPopwin = "QueryPopwin"
        Const ByDataTablePopwin = "DataTablePopwin"
    End Structure

    Public vbuCrpDtb As DataTable
    Public vbuCrpDtb1 As DataTable
    Public vbuCrpDtb2 As DataTable
    Public vbuCrpDtb3 As DataTable
    Public vbuCrpDtb4 As DataTable

    Public Function fbuOpenCrpViewer(vriRootFolder As String, vriParam As String) As String
        'fbuOpenCrpViewer = "window.open(""" & vriRootFolder & "/Preview/WbfCrpViewerPopUp.aspx?" & vriParam & """, " & cbuPopwin & ");"
        fbuOpenCrpViewer = "window.open(""" & vriRootFolder & "Preview/WbfCrpViewer.aspx?" & vriParam & """, " & cbuPopwinCrp & ");"
    End Function

    Public Function fbuOpenCrpViewerPopUp(vriRootFolder As String, vriParam As String) As String
        fbuOpenCrpViewerPopUp = "window.open(""" & vriRootFolder & "Preview/WbfCrpViewerPopUp.aspx?" & vriParam & """, " & cbuPopwinCrp & ");"
    End Function

    Public Function fbuOpenSmkViewer(vriRootFolder As String, vriParam As String) As String
        fbuOpenSmkViewer = "window.open(""" & vriRootFolder & "/Preview/WbfSmkViewer.aspx?" & vriParam & """, " & cbuPopwinCrp & ");"
    End Function

    Public Sub pbuClearCrpSession()
        HttpContext.Current.Session("CrpFileName") = ""
        HttpContext.Current.Session("CrpSubReport1") = ""
        HttpContext.Current.Session("CrpSubReport2") = ""
        HttpContext.Current.Session("CrpSubReport3") = ""
        HttpContext.Current.Session("CrpSubReport4") = ""

        HttpContext.Current.Session("CrpQuery") = ""
        HttpContext.Current.Session("CrpQuery1") = ""
        HttpContext.Current.Session("CrpQuery2") = ""
        HttpContext.Current.Session("CrpQuery3") = ""
        HttpContext.Current.Session("CrpQuery4") = ""
    End Sub

    Public Function fbuPrevieCrpWebDtb(vriCrv As CrystalDecisions.Web.CrystalReportViewer) As Boolean
        fbuPrevieCrpWebDtb = False
        Try

            Dim vnRptFolder = HttpContext.Current.Server.MapPath("~/CrpFiles/")
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
            If vbuCrpShowGroupTree Then
                vriCrv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                vriCrv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If HttpContext.Current.Session("SessionPreview") = "Crp" Then
                vriCrv.ReportSource = vnRpd
            ElseIf HttpContext.Current.Session("SessionPreview") = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, HttpContext.Current.Response, False, "")
            ElseIf HttpContext.Current.Session("SessionPreview") = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, HttpContext.Current.Response, False, "")
            End If

            vbuCrpDtb.Dispose()
            vbuCrpDtb1.Dispose()
            vbuCrpDtb2.Dispose()
            vbuCrpDtb3.Dispose()
            vbuCrpDtb4.Dispose()

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            fbuPrevieCrpWebDtb = True
        Catch ex As Exception
            pbMsgError = ex.Message
        End Try
    End Function

    Public Function fbuPrevieCrpWebQuery(vriCrv As CrystalDecisions.Web.CrystalReportViewer, vnSQLConn As SqlConnection) As Boolean
        fbuPrevieCrpWebQuery = False
        Try
            Dim vnRptFolder = HttpContext.Current.Server.MapPath("~/CrpFiles/")
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
            If vbuCrpShowGroupTree Then
                vriCrv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.GroupTree
            Else
                vriCrv.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None
            End If
            If HttpContext.Current.Session("SessionPreview") = "Crp" Then
                vriCrv.ReportSource = vnRpd
            ElseIf HttpContext.Current.Session("SessionPreview") = "Pdf" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, HttpContext.Current.Response, False, "")
            ElseIf HttpContext.Current.Session("SessionPreview") = "Xlsx" Then
                vnRpd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, HttpContext.Current.Response, False, "")
            End If

            vnDtb.Reset()
            vnDtb1.Reset()
            vnDtb2.Reset()
            vnDtb3.Reset()
            vnDtb4.Reset()

            vbuCrpQuery = ""
            vbuCrpQuery1 = ""
            vbuCrpQuery2 = ""
            vbuCrpQuery3 = ""
            vbuCrpQuery4 = ""

            vnRpd.Close()
            vnRpd.Dispose()

            fbuPrevieCrpWebQuery = True
        Catch ex As Exception
            pbMsgError = ex.Message
        End Try
    End Function
End Module
