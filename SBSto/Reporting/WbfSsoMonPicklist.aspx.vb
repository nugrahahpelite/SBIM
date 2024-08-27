Imports System.Data.SqlClient
Imports System.IO
Public Class WbfSsoMonPicklist
    Inherits System.Web.UI.Page
    Const csModuleName = "WbfSsoMonPicklist"

    Enum ensColList
        CompanyCode = 0
        vDcmSchLeaveDate = 1
        vCaraBayar_TransCode = 2
        vCaraBayar_TransStatus = 3
        vCaraBayar = 4
        vTotalHarga = 5
    End Enum

    Private Sub psDefaultDisplay()

    End Sub
    Protected Overrides ReadOnly Property PageStatePersister As PageStatePersister
        Get
            Return New SessionPageStatePersister(Me)
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")

        Session("CurrentFolder") = "Reporting"
        If Not IsPostBack Then
            psDefaultDisplay()
            'TxtListStart.Text = Format(DateAdd(DateInterval.Day, -2, Date.Now), "dd MMM yyyy")
            'TxtListEnd.Text = Format(Date.Now, "dd MMM yyyy")

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If

            If Session("UserCompanyCode") = "" Then
                pbuFillDstCompany(DstCompany, False, vnSQLConn)
            Else
                pbuFillDstCompanyByUser(Session("UserOID"), DstCompany, False, vnSQLConn)
            End If
            pbuFillDstWarehouse(DstListWarehouse, True, vnSQLConn)
            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Private Sub psFillGrvList()
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If

        Dim vnCrCB As String = ""

        Dim vnCrDate As String = ""

        If IsDate(TxtListStart.Text) Then
            vnCrDate += vbCrLf & "            and cast(pcl.CreationDatetime as date) >= '" & TxtListStart.Text & "'"
        End If
        If IsDate(TxtListEnd.Text) Then
            vnCrDate += vbCrLf & "            and cast(pcl.CreationDatetime as date) <= '" & TxtListEnd.Text & "'"
        End If

        Dim vnQuery As String
        Dim vnDtb As New DataTable

        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")
        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnDBDcm As String = fbuGetDBDcm()
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
        vnQuery += vnCrDate

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
        vnQuery += vbCrLf & "		 Where pcl.SchDTypeOID = " & enuSchDType.DO_Titip
        vnQuery += vnCrDate

        vnQuery += vbCrLf & "		 ) tb "
        vnQuery += vbCrLf & " Where 1 = 1 "

        If Val(DstListWarehouse.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "      And WarehouseOID =" & DstListWarehouse.SelectedValue
        End If
        If Val(DstCompany.SelectedValue) > 0 Then
            vnQuery += vbCrLf & "      And CompanyCode =" & DstCompany.SelectedValue
        End If

        If Trim(TxtPCLNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and PCLNo like '%" & fbuFormatString(Trim(TxtPCLNo.Text)) & "%'"
        End If
        If Trim(TxtPCLRefNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and PCLRefHNo like '%" & fbuFormatString(Trim(TxtPCLRefNo.Text)) & "%'"
        End If
        If Trim(TxtPickNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and isnull(PCKNo,'') like '%" & fbuFormatString(Trim(TxtPickNo.Text)) & "%'"
        End If
        If Trim(TxtDispatchNo.Text) <> "" Then
            vnQuery += vbCrLf & "            and isnull(DSPNo,'') like '%" & fbuFormatString(Trim(TxtDispatchNo.Text)) & "%'"
        End If

        vnQuery += vbCrLf & " Order by CreationDatetime DESC"

        pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
        GrvList.DataSource = vnDtb
        GrvList.DataBind()




        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub


    Protected Sub BtnListFind_Click(sender As Object, e As EventArgs) Handles BtnListFind.Click
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx")
        psFillGrvList()
    End Sub

    Protected Sub GrvList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrvList.SelectedIndexChanged

    End Sub

    Private Sub GrvList_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GrvList.RowCommand
        If Session("UserName") = "" Then Response.Redirect("~/Default.aspx?vpSessionEnd=1")
        Dim vnIdx As Integer = Convert.ToInt32(e.CommandArgument)
        If vnIdx >= GrvList.Rows.Count Then Exit Sub

        Dim vnGRow As GridViewRow = GrvList.Rows(vnIdx)
        If e.CommandName = "vTotalHarga" Then
            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMsgError.Text = pbMsgError
                LblMsgError.Visible = True
                Exit Sub
            End If


            vnSQLConn.Close()
            vnSQLConn.Dispose()
            vnSQLConn = Nothing
        End If
    End Sub

    Protected Sub BtnXLS_Click(sender As Object, e As EventArgs) Handles BtnXLS.Click
        Dim vnSQLConn As New SqlConnection
        If Not fbuConnectSQL(vnSQLConn) Then
            LblMsgError.Text = pbMsgError
            LblMsgError.Visible = True
            Exit Sub
        End If
        Dim vnUserLocationOID As String = Session("UserLocationOID")
        Dim vnUserCompanyCode As String = Session("UserCompanyCode")

        Dim vnFileName As String = ""
        pbuCreateXlsx_MonitoringPickList1(vnFileName, Session("UserOID"), DstListWarehouse, DstCompany,
                                          TxtPCLNo, TxtPCLRefNo, TxtPickNo, TxtDispatchNo,
                                          TxtListStart, TxtListEnd, vnSQLConn)

        vnSQLConn.Close()
        vnSQLConn.Dispose()
        vnSQLConn = Nothing
    End Sub

End Class