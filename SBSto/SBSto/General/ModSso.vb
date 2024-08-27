Imports System.Data.SqlClient
Module ModSso
    Public Const cbuPopwinStatus = "'popwin', 'width=1000, height=650, left=250, top=250'"
    Public Const cbuQR_NoTerima = "No.Terima:"
    Public Const cbuQR_TglTerima = "Tgl Terima:"
    Public Const cbuQR_IDTerima = "ID Terima:"
    Public Enum enuUserGroup
        Admin = 1
    End Enum

    Public Structure stuFolderName
        Const WebLog = "WebLog"
        Const SAPApiFolder = "SAPApiFolder"
    End Structure
    Public Enum enuUserOID_Special
        Customer = 1111
        Tukar_Faktur = 1112
    End Enum

    Public Enum enuUploadSource
        Xls = 1
        SAP_Api = 2
    End Enum
    Public Enum enuWarehouseOID
        Kepu = 1
        Prancis = 2
        Prancis2 = 13
        Surabaya = 5
    End Enum

    Public Structure stuWarehouse
        Const Kepu_Baru = "KEA01"
        Const Prancis_Baru = "PRA01"
    End Structure

    Public Structure stuCompanyCode
        Const BAD = "BAD"
        Const CAS = "CAS"
    End Structure
    Public Enum enuSOType
        WinAcc = 1
        CcLocation = 2
        CcBarang = 3
    End Enum
    Public Enum enuStorageType
        Rack = 1002
        Floor = 1003
        Staging = 1004
        CrossDock = 1005
        Karantina = 1006
        DO_Titip = 1007
        Damage = 1008
    End Enum

    Public Enum enuStaggingType
        Stag_IN = 1
        Stag_OUT = 2
    End Enum

    Public Enum enuRcvPOType
        Import = 1
        Local = 2
    End Enum

    Public Enum enuRcvKRType
        Release_Minus = -2
        Receive_Minus = -1
        Receive_Plus = 1
    End Enum
    Public Enum enuDestType
        None = 0
        Dalam_Kota = 1
        Luar_Kota = 2
        Lain_lain = 3
    End Enum
    Public Enum enuSmTRBDTypeOID
        Sales_Order = 1
        Picklist_Gantung = 2
        Stock_Gudang = 3
    End Enum
    Public Enum enuRcvType
        Opening_Balance = 0
        Pembelian = 1
        Retur = 2
        Lain_lain = 3
        Karantina = 4
    End Enum

    Public Enum enuAJMN 'SsoAdjustmentMinus
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Approved = 19
    End Enum
    Public Enum enuCSKU 'SsoChangeSKU
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Approved = 4
    End Enum

    Public Enum enuTCVOSO 'Void SO
        Cancelled = -2
        Baru = 0
        Prepared = 2
    End Enum
    Public Enum enuTCSSOH
        Cancelled = -2
        Baru = 0
        Scan_Open = 4
        Scan_Closed = 10
        Closed = 20
    End Enum

    Public Enum enuTCSSOC 'SO Compare
        Cancelled = -2
        Baru = 0
        Recompare = 15
        Closed = 20
    End Enum
    Public Enum enuTCSRCV 'Receiving
        Cancelled = -2
        Baru = 0
        Scan_Open = 4
        Scan_Closed = 10
        Closed = 20
    End Enum

    Public Enum enuTCSPCK_ga_pake_lagi 'Picking
        Cancelled = -2
        Baru = 0
        Scan_Open = 4
        Scan_Closed = 10
        Closed = 20
    End Enum

    Public Enum enuTCPICK 'Picklist
        Void = -6
        Cancelled = -2
        Baru = 0
        Prepared_Failed = 1
        Prepared = 2
        On_Picking = 16
        Picking_Done = 18
    End Enum
    Public Enum enuTCPCKG 'Picking
        Void = -6
        Cancelled = -2
        None = 0
        On_Picking = 16
        Picking_Done = 18
        On_Dispatch_Putaway = 19
        On_Putaway_Void = 21
        Move_Antar_StagingOut_Done = 22
        Putaway_Dispatch_Done = 24
    End Enum
    Public Enum enuTCDISP 'Dispatch
        Cancelled = -2
        On_Dispatch = 16
        Dispatch_Done = 18
        Driver_Confirm = 19
        Closed = 20
    End Enum
    Public Enum enuTCDISR 'Dispatch Receive
        Cancelled = -2
        On_Dispatch_Receive = 16
        Dispatch_Receive_Done = 18
        Driver_Receive_Confirm = 19
        Putaway_Process = 22
        All_Putaway_Clomplete = 24
    End Enum

    Public Enum enuTCDISG 'Dispatch Receive (Picking Status)
        Baru = 0
        Ready_To_Dispatch = 2
        On_Dispatch = 16
        Dispatch_Done = 18
        Driver_Confirm = 19
        Closed = 20
    End Enum
    Public Enum enuTCSPPO 'Upload PO
        Baru = 0
        In_PL = 12
        Sudah_Ada_Penerimaan_Selesai = 17
        Closed = 20
    End Enum

    Public Enum enuTCDSGO 'Moving Antar Staging Out
        Cancelled = 2
        Staging_Out_1_Preparation = 14
        On_Delivery_To_Staging_Out_2 = 16
        Staging_Out_2_On_Receiving = 18
        Staging_Out_2_Receive_Done = 19
        Closed = 20
    End Enum

    Public Enum enuTCCSSO 'Upload Customer SO
        Baru = 0
        In_TRB_Calculation = 12
    End Enum

    Public Enum enuTCPLSP 'Packing List dari Supplier
        Cancelled = -2
        Baru = 0
        Prepared = 2
        On_Receive = 16
        Receive_Done = 18
    End Enum

    Public Enum enuTCRCPO 'Penerimaan Pembelian
        Cancelled = -2
        On_Receive = 16
        Receive_Done = 18
        Receive_Approved = 19
        Putaway_Process = 22
        All_Putaway_Clomplete = 24
    End Enum

    Public Enum enuTCRCMS 'Penerimaan Lain-lain
        Cancelled = -2
        Baru = 0
        Prepared = 2
        On_Receive = 16
        Receive_Done = 18
    End Enum
    Public Enum enuTCRCKR 'Penerimaan Karantina
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Approved = 4
        On_Receive = 16
        Receive_Done = 18
    End Enum
    Public Enum enuTCPWAY 'Putaway gudang sama
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum

    Public Enum enuTCPYAY 'Putaway beda gudang
        Cancelled = -2
        Baru = 0
        On_Delivery_Putaway = 14
        On_Putaway = 16
        Putaway_Done = 18
    End Enum

    Public Enum enuTCPDLW 'Pindah Lokasi antar gudang
        Cancelled = -2
        Baru = 0
        On_Delivery_Movement = 14
        On_Movement = 16
        Movement_Done = 18
    End Enum
    Public Enum enuTCPDLK 'Pindah Lokasi gudang sama
        Cancelled = -2
        Baru = 0
        On_Movement = 16
        Movement_Done = 18
    End Enum
    Public Enum enuTCSMTB 'Summary Barang untuk TRB
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Closed_Sudah_TRB = 20
    End Enum
    Public Enum enuTCPDTW 'Putaway DO Titip gudang sama
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum
    Public Enum enuTCPDTY 'Putaway DO Titip beda gudang
        Cancelled = -2
        Baru = 0
        On_Delivery_Putaway = 14
        On_Putaway = 16
        Putaway_Done = 18
    End Enum

    Public Enum enuTCPDSW 'Putaway Penerimaan Dispatch gudang sama
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum
    Public Enum enuTCPDSY 'Putaway Penerimaan Dispatch beda gudang
        Cancelled = -2
        Baru = 0
        On_Delivery_Putaway = 14
        On_Putaway = 16
        Putaway_Done = 18
    End Enum
    Public Enum enuTCPDTV 'Putaway Void
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum

    Public Enum enuTCPTKW 'Putaway Karantina
        Cancelled = -2
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
    End Enum
    Public Enum enuTCSTKR 'Stock Karantina
        Baru = 0
        On_Putaway = 16
        Putaway_Done = 18
        Approved = 19
    End Enum
    Public Enum enuUserLocation
        All = 0
    End Enum
    Public Structure stuSsoCrp
        Const CrpBnsrphBarcodeSelectionQR = "CrpBnsrphBarcodeSelectionQR.rpt"
        Const CrpBnsrph = "CrpBnsrph.rpt"
        Const CrpSsoSOTally = "CrpSsoSOTally.rpt"
        Const CrpSsoSOTally_Storage = "CrpSsoSOTally_Storage.rpt"
        Const CrpSsoSOCycleCount = "CrpSsoSOCycleCount.rpt"
        Const CrpSsoSOStatus = "CrpSsoSOStatus.rpt"

        Const CrpSsoSOTallyDetail = "CrpSsoSOTallyDetail.rpt"
        Const CrpSsoSOTallyRcv = "CrpSsoSOTallyRcv.rpt"
        Const CrpSsoSOTallyPick = "CrpSsoSOTallyPick.rpt"
        Const CrpSsoSOTallyCompare = "CrpSsoSOTallyCompare.rpt"
        Const CrpSsoSOTallyCompareDetail = "CrpSsoSOTallyCompareDetail.rpt"
        Const CrpSsoSOTallyCompareDetail2 = "CrpSsoSOTallyCompareDetail2.rpt"
        Const CrpSsoStorage = "CrpSsoStorage.rpt"
        Const CrpSsoSmTRBSummary = "CrpSsoSmTRBSummary.rpt"
        Const CrpSsoPickList = "CrpSsoPickList.rpt"
        Const CrpSsoPickListReserve_SR = "CrpSsoPickListReserve_SR.rpt"
        Const CrpSsoSummPutw = "CrpSsoSummPutw.rpt"
        Const CrpSsoDailyCheckPutw = "CrpSsoDailyCheckPutw.rpt"
    End Structure

    Public Structure stuSsoReportType
        Const RptSOTally = "SOTally"
        Const RptSODetail = "SODetail"
        Const RptSOTallyRcv = "SOTallyRcv"
        Const RptSOTallyPick = "SOTallyPick"
        Const RptSOTallyCompare = "SOTallyCompare"
        Const RptSOCompareDetail = "SOCompareDetail"
        Const RptCycleCount = "SOTally_Storage"
        Const RptReportingSummPutw = "RptReportingSummPutw"
        Const RptReportingDailyCheckPutw = "RptReportingDailyCheckPutw"
        Const RptSOTallyCompare2 = "SOTallyCompare2"
        Const RptSOCompareDetail2 = "SOCompareDetail2"
        Const RptSOStatus = "SOStatus"
    End Structure

    Public Function fbuOpenTransStatus(vriRootFolder As String, vriParam As String) As String
        fbuOpenTransStatus = "window.open(""" & vriRootFolder & "Sso/WbfSsoTransStatus.aspx?" & vriParam & """, " & cbuPopwinStatus & ");"
    End Function

    Public Sub pbuFillDstCompany(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select '' CompanyCode,'ALL' CompanyName UNION"
        Else
            vnQuery += vbCrLf & "Select '' CompanyCode,'' CompanyName UNION"
        End If
        vnQuery += vbCrLf & "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany with(nolock))tb "
        vnQuery += vbCrLf & "order by case when CompanyCode='' then '' else CompanyName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CompanyCode"
            vriDst.DataTextField = "CompanyName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstUserGroup(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' SsoUserGroupName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' SsoUserGroupName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,SsoUserGroupName From Sys_SsoUserGroup_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else SsoUserGroupName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "SsoUserGroupName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstCity(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 CityOID,'ALL' CityName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 CityOID,'' CityName UNION"
        End If
        vnQuery += vbCrLf & "Select CityOID,CityName From " & fbuGetDBMaster() & "DimCity with(nolock))tb "
        vnQuery += vbCrLf & "order by case when CityOID=0 then '' else CityName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CityOID"
            vriDst.DataTextField = "CityName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstStorageType(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vStorageTypeName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vStorageTypeName UNION"
        End If
        'vnQuery += vbCrLf & "Select OID,StorageTypeName + ' - IsRack = '+ case when abs(IsRack)=1 then 'Y' else 'N' end vStorageTypeName From " & fbuGetDBMaster() & "Sys_StorageType_MA)tb "
        vnQuery += vbCrLf & "Select OID,StorageTypeName vStorageTypeName From " & fbuGetDBMaster() & "Sys_StorageType_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else vStorageTypeName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vStorageTypeName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuSetStorageTypeIs(vriStoTypeOID As String, vriChkIsMultiLevel As CheckBox, vriChkIsRack As CheckBox, vriChkIsStagging As CheckBox, vriChkIsCrossDock As CheckBox, vriChkIsKarantina As CheckBox, vriChkIsDOTitip As CheckBox, vriChkIsDamage As CheckBox, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select abs(IsMultiLevel)vIsMultiLevel,abs(IsRack)vIsRack,abs(IsStagging)vIsStagging,abs(IsCrossDock)vIsCrossDock,abs(IsKarantina)vIsKarantina,abs(IsDOTitip)vIsDOTitip,abs(IsDamage)vIsDamage From " & fbuGetDBMaster() & "Sys_StorageType_MA with(nolock) Where OID=" & vriStoTypeOID
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriChkIsCrossDock.Checked = IIf(vnDtb.Rows(0).Item("vIsCrossDock") = "1", True, False)
            vriChkIsMultiLevel.Checked = IIf(vnDtb.Rows(0).Item("vIsMultiLevel") = "1", True, False)
            vriChkIsRack.Checked = IIf(vnDtb.Rows(0).Item("vIsRack") = "1", True, False)
            vriChkIsStagging.Checked = IIf(vnDtb.Rows(0).Item("vIsStagging") = "1", True, False)
            vriChkIsKarantina.Checked = IIf(vnDtb.Rows(0).Item("vIsKarantina") = "1", True, False)
            vriChkIsDOTitip.Checked = IIf(vnDtb.Rows(0).Item("vIsDOTitip") = "1", True, False)
            vriChkIsDamage.Checked = IIf(vnDtb.Rows(0).Item("vIsDamage") = "1", True, False)
        End If
    End Sub

    Public Sub pbuFillDstWarehouse(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' WarehouseName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' WarehouseName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,WarehouseName From " & fbuGetDBMaster() & "Sys_Warehouse_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else WarehouseName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "WarehouseName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstWarehouse_Pr(vriWhsOID As Integer, vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' WarehouseName UNION"
        Else
            vnQuery += vbCrLf & ""
        End If
        vnQuery += vbCrLf & "Select OID,WarehouseName From " & fbuGetDBMaster() & "Sys_Warehouse_MA with(nolock) "

        If vriWhsOID = enuWarehouseOID.Prancis Or vriWhsOID = enuWarehouseOID.Prancis2 Then
            vnQuery += vbCrLf & "Where OID in(" & enuWarehouseOID.Prancis & "," & enuWarehouseOID.Prancis2 & ")"
        Else
            vnQuery += vbCrLf & "Where OID=" & vriWhsOID
        End If

        vnQuery += vbCrLf & ")tb"


        vnQuery += vbCrLf & "order by case when OID=0 then '' else WarehouseName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "WarehouseName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstWarehouse_ByUserOID(vriUserOID As String, vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnUserWarehouseCode As String = HttpContext.Current.Session("UserWarehouseCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' WarehouseName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' WarehouseName UNION"
        End If
        vnQuery += vbCrLf & "Select wh.OID,wh.WarehouseName "
        vnQuery += vbCrLf & "	   From " & fbuGetDBMaster() & "Sys_Warehouse_MA wh with(nolock)"

        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "	         inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=wh.OID"
            vnQuery += vbCrLf & "	   Where uw.UserOID=" & vriUserOID
        End If
        vnQuery += vbCrLf & ")tb order by case when OID=0 then '' else WarehouseName end"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "WarehouseName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstSubWarehouse(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vSubWhsName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vSubWhsName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,SubWhsCode + ' - ' + SubWhsName vSubWhsName From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else vSubWhsName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vSubWhsName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstSubWarehouse_ByCompanyCode(vriDst As DropDownList, vriAll As Boolean, vriCompanyCode As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vSubWhsName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vSubWhsName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,SubWhsCode + ' - ' + SubWhsName vSubWhsName From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA with(nolock) Where CompanyCode='" & vriCompanyCode & "')tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else vSubWhsName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vSubWhsName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstSubWarehouse_ByCompanyCode_ByUserOID(vriDst As DropDownList, vriAll As Boolean, vriCompanyCode As String, vriUserOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vSubWhsName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vSubWhsName UNION"
        End If
        vnQuery += vbCrLf & "Select sw.OID,sw.SubWhsCode + ' - ' + sw.SubWhsName vSubWhsName"
        vnQuery += vbCrLf & "       From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA sw with(nolock)"
        vnQuery += vbCrLf & "	         inner join Sys_SsoUserWarehouse_MA wh with(nolock) on wh.WarehouseOID=sw.WarehouseOID"
        vnQuery += vbCrLf & "	   Where sw.CompanyCode='" & vriCompanyCode & "' and wh.UserOID=" & vriUserOID & ")tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else vSubWhsName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vSubWhsName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstSubWarehouse_ByUserOID(vriUserOID As String, vriCompCode As String, vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vSubWhsName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vSubWhsName UNION"
        End If
        vnQuery += vbCrLf & "Select sw.OID,sw.SubWhsCode + ' - ' + sw.SubWhsName vSubWhsName"
        vnQuery += vbCrLf & "       From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA sw with(nolock)"
        vnQuery += vbCrLf & "	         inner join Sys_SsoUserWarehouse_MA wh with(nolock) on wh.WarehouseOID=sw.WarehouseOID"
        vnQuery += vbCrLf & "	   Where wh.UserOID=" & vriUserOID & " and sw.CompanyCode='" & vriCompCode & "')tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else vSubWhsName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vSubWhsName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstBuilding(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' BuildingName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' BuildingName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,BuildingName From " & fbuGetDBMaster() & "Sys_Building_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else BuildingName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "BuildingName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstBuilding_ByWarehouse(vriDst As DropDownList, vriAll As Boolean, vriWhsOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' BuildingName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' BuildingName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,BuildingName From " & fbuGetDBMaster() & "Sys_Building_MA with(nolock) Where WarehouseOID=" & vriWhsOID & ")tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else BuildingName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "BuildingName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Function fbuGetWarehouseOID_ByBuildingOID(vriBuildingOID As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select WarehouseOID From " & fbuGetDBMaster() & "Sys_Building_MA with(nolock) Where OID=" & vriBuildingOID
        Return (fbuGetDataNumSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetWarehouseName(vriWhsOID As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select WarehouseName From " & fbuGetDBMaster() & "Sys_Warehouse_MA with(nolock) Where OID=" & vriWhsOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetBuildingLantaiRelOID(vriBuildingOID As String, vriLantaiOID As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select OID From " & fbuGetDBMaster() & "Sys_BuildingLantaiRel_MA with(nolock) Where BuildingOID=" & vriBuildingOID & " and LantaiOID=" & vriLantaiOID
        Return (fbuGetDataNumSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetBuildingLantaiZonaRelOID(vriBuildingOID As String, vriLantaiOID As String, vriZonaOID As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnBuildingLantaiRelOID As String = fbuGetBuildingLantaiRelOID(vriBuildingOID, vriLantaiOID, vriSQLConn)
        Dim vnQuery As String
        vnQuery = "Select OID From " & fbuGetDBMaster() & "Sys_BuildingLantaiZonaRel_MA with(nolock) Where BuildingLantaiRelOID=" & vnBuildingLantaiRelOID & " and ZonaOID=" & vriZonaOID
        Return (fbuGetDataNumSQL(vnQuery, vriSQLConn))
    End Function
    Public Sub pbuFillDstLantai(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' LantaiDescription UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' LantaiDescription UNION"
        End If
        vnQuery += vbCrLf & "Select OID,LantaiDescription From " & fbuGetDBMaster() & "Sys_Lantai_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else LantaiDescription end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "LantaiDescription"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstLantai_ByWarehouse_ByBuilding(vriDst As DropDownList, vriAll As Boolean, vriWhsOID As String, vriBuildingOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' LantaiDescription UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' LantaiDescription UNION"
        End If
        vnQuery += vbCrLf & "Select OID,LantaiDescription From " & fbuGetDBMaster() & "Sys_Lantai_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else LantaiDescription end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "LantaiDescription"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstZona(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' ZonaName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' ZonaName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,ZonaName From " & fbuGetDBMaster() & "Sys_Zona_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else ZonaName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "ZonaName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstCompanyByUser(vriUserOID As String, vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnUserCompanyCode As String
        vnQuery = "Select UserCompanyCode From Sys_SsoUser_MA Where OID=" & vriUserOID
        vnUserCompanyCode = fbuGetDataStrSQL(vnQuery, vriSQLConn)

        If vnUserCompanyCode = "" Then
            vnQuery = "Select * From ("
            If vriAll Then
                vnQuery += vbCrLf & "Select '' CompanyCode,'ALL' CompanyName UNION"
            Else
                vnQuery += vbCrLf & "Select '' CompanyCode,'' CompanyName UNION"
            End If
            vnQuery += vbCrLf & "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany with(nolock))tb "
            vnQuery += vbCrLf & "order by case when CompanyCode='' then '' else CompanyName end"

        Else
            vnQuery = "Select CompanyCode,CompanyName From " & fbuGetDBMaster() & "DimCompany Where CompanyCode in(Select b.CompanyCode From Sys_SsoUserCompany_MA b with(nolock) Where UserOID=" & vriUserOID & ")"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CompanyCode"
            vriDst.DataTextField = "CompanyName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstLocation(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' vLocationName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' vLocationName UNION"
        End If
        vnQuery += vbCrLf & "Select OID, LocationCode + ' ' + LocationName vLocationName From Sys_SsoLocation_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when vLocationName='' then '' else vLocationName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "vLocationName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstGudang(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select '' GdgCode,'ALL' GdgName UNION"
        Else
            vnQuery += vbCrLf & "Select '' GdgCode,'' GdgName UNION"
        End If
        vnQuery += vbCrLf & "Select GdgCode,GdgName From " & fbuGetDBMaster() & "Sys_MstGudang_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when GdgCode='' then '' else GdgName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "GdgCode"
            vriDst.DataTextField = "GdgName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillDstRcvType(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' RcvTypeName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' RcvTypeName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,RcvTypeName From Sys_SsoRcvType_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else RcvTypeName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "RcvTypeName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstGudang_ByUserLocation(vriDst As DropDownList, vriAll As Boolean, vriUserLocationOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select '' GdgCode,'ALL' GdgName UNION"
        Else
            vnQuery += vbCrLf & "Select '' GdgCode,'' GdgName UNION"
        End If

        vnQuery += vbCrLf & "Select GdgCode,GdgName From " & fbuGetDBMaster() & "Sys_MstGudang_MA with(nolock)"
        If vriUserLocationOID <> "" And vriUserLocationOID <> "0" Then
            vnQuery += " Where LocationOID=" & vriUserLocationOID
        End If
        vnQuery += ")tb "

        vnQuery += vbCrLf & "order by case when GdgCode='' then '' else GdgName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "GdgCode"
            vriDst.DataTextField = "GdgName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuInsertStatusSSOH(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoSOStatus_TR(SOHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoSSOH & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusSOCompareH(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoSOCompareStatus_TR(SOCHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoSSOC & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusReceiving(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoRcvStatus_TR(RcvHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoReceiving & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPicking(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPickStatus_TR(PickHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPicking & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPL(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPLStatus_TR(PLHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPOPackingList & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusSTKR(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoStockKarantinaStatus_TR(STKRHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoStockKarantina & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPCL(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPCLStatus_TR(PCLHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPickList & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusPCK(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPCKStatus_TR(PCKHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPicking & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusSmTRB(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoSmTRBStatus_TR(SmTRBHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoSummaryTRB & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Function fbuGetLocationName(vriLocationOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select LocationName From Sys_SsoLocation_MA where OID=" & vriLocationOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetLocationOID_ByGudangCode(vriGdgCode As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select LocationOID From " & fbuGetDBMaster() & "Sys_MstGudang_MA where GdgCode='" & vriGdgCode & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetWarehouseOID_BySubWhsOID(vriSubWhsOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select WarehouseOID From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA where OID='" & vriSubWhsOID & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetSubWhsCode(vriSubWhsOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select SubWhsCode From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA where OID='" & vriSubWhsOID & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetSubWhsCode_ByOID(vriSubWhsOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select SubWhsCode From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA where OID='" & vriSubWhsOID & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetSubWhsCode_ByOID_Trans(vriSubWhsOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As String
        Dim vnQuery As String
        vnQuery = "Select SubWhsCode From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA where OID='" & vriSubWhsOID & "'"
        Return (fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans))
    End Function

    Public Function fbuGetSubWhOID_BySubWhsName(vriSubWhsName As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select OID From " & fbuGetDBMaster() & "Sys_SubWarehouse_MA where SubWhsName='" & vriSubWhsName & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetWhsCode_ByOID(vriWhsOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select WarehouseCode From " & fbuGetDBMaster() & "Sys_Warehouse_MA where OID=" & vriWhsOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetGudangCode_ByGudangName(vriGdgName As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select GdgCode From " & fbuGetDBMaster() & "Sys_MstGudang_MA where GdgName='" & vriGdgName & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetBarangUnit(vriCompanyCode As String, vriBrgCode As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select BRGUNIT From " & fbuGetDBMaster() & "Sys_MstBarang_MA where CompanyCode='" & vriCompanyCode & "' and BRGCODE='" & vriBrgCode & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Function fbuGetBarangUnitTrans(vriCompanyCode As String, vriBrgCode As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction) As String
        Dim vnQuery As String
        vnQuery = "Select BRGUNIT From " & fbuGetDBMaster() & "Sys_MstBarang_MA where CompanyCode='" & vriCompanyCode & "' and BRGCODE='" & vriBrgCode & "'"
        Return (fbuGetDataStrSQLTrans(vnQuery, vriSQLConn, vriSQLTrans))
    End Function
    Public Function fbuGetPL_ReceiveNo(vriRcvPORefOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select RcvPONo From Sys_SsoRcvPOHeader_TR Where RcvPORefTypeOID=" & enuRcvPOType.Import & " and RcvPORefOID=" & vriRcvPORefOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetNamaBarang(vriCompanyCode As String, vriBrgCode As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select BRGNAME From " & fbuGetDBMaster() & "Sys_MstBarang_MA where CompanyCode='" & vriCompanyCode & "' and BRGCODE='" & vriBrgCode & "'"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Sub pbuFillRdlUserGroup(vriRdl As RadioButtonList, vriSQLConn As SqlClient.SqlConnection)
        vriRdl.Items.Clear()

        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select OID,SsoUserGroupName,SsoUserGroupDescr From Sys_SsoUserGroup_MA with(nolock) order by SsoUserGroupName"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            Dim vnLItmX As ListItem
            Dim vnDRow As DataRow
            For vn = 0 To vnDtb.Rows.Count - 1
                vnDRow = vnDtb.Rows(vn)
                vnLItmX = New ListItem
                vnLItmX.Text = vnDRow.Item("SsoUserGroupName") & ", " & vnDRow.Item("SsoUserGroupDescr")
                vnLItmX.Value = vnDRow.Item("OID")
                vriRdl.Items.Add(vnLItmX)
            Next
        End If
    End Sub

    Public Function fbuGetPOHOID_ByPODOID(vriPODOID As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Select POHOID From Sys_SsoPODetail_TR Where OID=" & vriPODOID
        Return fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
    End Function

    Public Sub pbuPODetail_UpdatePLQty(vriPODOID As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Update Sys_SsoPODetail_TR set Qty_PL=isnull((Select sum(PLDQty) From Sys_SsoPLDetail_TR Where PODOID=" & vriPODOID & " and PLHOID in(Select b.OID From Sys_SsoPLHeader_TR b Where b.TransStatus<>" & enuTCPLSP.Cancelled & ")),0) Where OID=" & vriPODOID
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Function fbuGetRcvPODate(vriRcvPOHOID As String, vriFormat As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery As String
        If vriFormat = "ddMMMyy" Then
            vnQuery = "Select Convert(varchar(6),RcvPODate,106)+' '+Substring(Convert(varchar(4),RcvPODate,112),3,2) "
        Else
            vnQuery = "Select Convert(varchar(11),RcvPODate,106) "
        End If
        vnQuery += "From Sys_SsoRcvPOHeader_TR Where OID=" & vriRcvPOHOID
        Return fbuGetDataNumSQL(vnQuery, vriSQLConn)
    End Function

    Public Function fbuGetCutOfDate(vriCompanyCode As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery As String
        vnQuery = "Select Convert(varchar(11),CutOfDate,106) From Sys_SsoCutOfDate_CNF Where CompanyCode='" & vriCompanyCode & "'"
        Return fbuGetDataStrSQL(vnQuery, vriSQLConn)
    End Function

    Public Function fbuGetRcvPODate_ByPLHOID(vriPLHOID As String, vriFormat As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery As String
        If vriFormat = "ddMMMyy" Then
            vnQuery = "Select Convert(varchar(6),RcvPODate,106)+' '+Substring(Convert(varchar(4),RcvPODate,112),3,2) "
        Else
            vnQuery = "Select Convert(varchar(11),RcvPODate,106) "
        End If
        vnQuery += "From Sys_SsoRcvPOHeader_TR Where RcvPORefTypeOID=" & enuRcvPOType.Import & " and RcvPORefOID=" & vriPLHOID
        Return fbuGetDataStrSQL(vnQuery, vriSQLConn)
    End Function

    Public Function fbuGetMstStorageInfo_ByStorageOID(vriStorageOID As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select vStorageInfo From " & fbuGetDBMaster() & "fnTbl_SsoStorage_ByStorageOID('" & vriStorageOID & "')"
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function


    Public Function fbuGetPCKHOID_By_PickListHOID(vriPCLOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select OID From Sys_SsoPCKHeader_TR with(nolock) Where PCLHOID=" & vriPCLOID & " and TransStatus>" & enuTCPCKG.None
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function
    Public Function fbuGetPCKTransStatus(vriPCKHOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select TransStatus From Sys_SsoPCKHeader_TR with(nolock) Where OID=" & vriPCKHOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Sub pbuGetDtbPCKHOID_By_PickListHOID(vriDtb As DataTable, vriPCLHOID As String, vriSQLConn As SqlConnection)
        Dim vnQuery As String
        vnQuery = "Select *,ST.TransStatusDescr"
        vnQuery += vbCrLf & " From Sys_SsoPCKHeader_TR PM with(nolock)"
        vnQuery += vbCrLf & "      inner join Sys_SsoTransStatus_MA  ST with(nolock) on ST.TransStatus=PM.TransStatus and ST.TransCode=PM.TransCode"
        vnQuery += vbCrLf & "Where PM.PCLHOID=" & vriPCLHOID
        pbuFillDtbSQL(vriDtb, vnQuery, vriSQLConn)
    End Sub

    Public Function fbuGetPCLTransStatus(vriPCLHOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select TransStatus From Sys_SsoPCLHeader_TR with(nolock) Where OID=" & vriPCLHOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Sub pbuInsertStatusRcvMsc(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoRcvMscStatus_TR(RcvMscHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPenerimaanLain2 & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusRcvKR(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoRcvKRStatus_TR(RcvKRHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPenerimaanKarantina & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuSsoProcessDataKey(vriProcessDataKey As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoProcessDataKey(ProcessDataKey)"
        vnQuery += vbCrLf & "values('" & vriProcessDataKey & "')"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusRcvPO(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoRcvPOStatus_TR(RcvPOHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPenerimaanPembelian & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusPW(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPWStatus_TR(PWHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDTW(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDTWStatus_TR(DTWHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_DO_Titip & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPTV(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPTVStatus_TR(PTVHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_DO_Titip & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDSW(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDSWStatus_TR(DSWHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_Penerimaan_Dispatch & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPTK(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPTKStatus_TR(PTKHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_Karantina & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDSP(vriOID As Integer, vriDriverOID As Integer, vriVehicleOID As Integer, vriStatusNote As String, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDSPStatus_TR(DSPHOID,DcmSchDriverOID,DcmVehicleOID,StatusNote,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & "," & vriDriverOID & "," & vriVehicleOID & ",'" & vriStatusNote & "','" & stuTransCode.SsoDispatch & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusSGO(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoSGOStatus_TR(SGOHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoMoving_Antar_StagingOut & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDSR(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDSRStatus_TR(DSRHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoDispatchReceive & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusStockPick(vriOID As Integer, vriDSRHOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoStockPickStatus_TR(StockPickHOID,DSRHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & "," & vriDSRHOID & ",'" & stuTransCode.SsoDispatchReceive_Picking_Status & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPDL(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPDLStatus_TR(PDLHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPindahLokasi & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPDW(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPDWStatus_TR(PDWHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPindahLokasi_Antar_Wh & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusPY(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoPYStatus_TR(PYHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_Antar_Wh & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDTY(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDTYStatus_TR(DTYHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_DO_Titip_Antar_Wh & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusDSY(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoDSYStatus_TR(DSYHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoPutaway_DO_Titip_Antar_Wh & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub

    Public Sub pbuInsertStatusVoidSO(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoSOrderVoidStatus_TR(SOrderVoidHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoVoidSO & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Sub pbuInsertStatusCSKU(vriOID As Integer, vriStatus As Integer, vriUserOID As String, vriSQLConn As SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnQuery As String
        vnQuery = "Insert into Sys_SsoCSKUStatus_TR(CSKUHOID,TransCode,TransStatus,TransStatusUserOID,TransStatusDatetime)"
        vnQuery += vbCrLf & "values(" & vriOID & ",'" & stuTransCode.SsoChangeSKU & "'," & vriStatus & ",'" & vriUserOID & "',getdate())"
        pbuExecuteSQLTrans(vnQuery, cbuActionEdit, vriSQLConn, vriSQLTrans)
    End Sub
    Public Function fbuGetDSPTransStatus(vriDSPHOID As String, vriSQLConn As SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select TransStatus From Sys_SsoDSPHeader_TR with(nolock) Where OID=" & vriDSPHOID
        Return (fbuGetDataStrSQL(vnQuery, vriSQLConn))
    End Function

    Public Sub pbuFillDtbDSPHeader(vriDtb As DataTable, vriDSPHOID As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnQuery As String
        vnQuery = "Select DcmSchDriverOID,DcmVehicleOID From Sys_SsoDSPHeader_TR with(nolock) Where OID=" & vriDSPHOID
        pbuFillDtbSQL(vriDtb, vnQuery, vriSQLConn)
    End Sub

    Public Function fbuValNotaDetail_BrgCode(vriCompanyCode As String, vriNotaHOID As String, vriSQLConn As SqlClient.SqlConnection) As Boolean
        pbMsgError = ""

        Dim vnDBMaster As String = fbuGetDBMaster()
        Dim vnQuery As String
        vnQuery = "Select KodeBarang From " & fbuGetDBDcm() & "Sys_DcmNotaDetail_TR with(nolock) Where NotaHOID=" & vriNotaHOID & " and NOT KodeBarang in"
        vnQuery += vbCrLf & "(Select BRGCODE FROM " & vnDBMaster & "Sys_MstBarang_MA with(nolock) Where CompanyCode='" & vriCompanyCode & "'"
        vnQuery += vbCrLf & "UNION"
        vnQuery += vbCrLf & "Select PAKETCODE FROM " & vnDBMaster & "Sys_MstPaketH_MA with(nolock) Where CompanyCode='" & vriCompanyCode & "')"

        Dim vnBrgCek As String = fbuGetDataStrSQL(vnQuery, vriSQLConn)
        If vnBrgCek = "" Then
            Return True
        Else
            pbMsgError = vnBrgCek & " TIDAK TERDAFTAR DI MASTER BARANG"
            Return False
        End If
    End Function

    Public Sub pbuFillCstCompany(vriDst As CheckBoxList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select '' CompanyCode,'ALL' vCompanyCode UNION"
        Else
            vnQuery += vbCrLf & "Select '' CompanyCode,'' vCompanyCode UNION"
        End If
        vnQuery += vbCrLf & "Select mc.CompanyCode,mc.CompanyCode vCompanyCode"
        vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "DimCompany mc with(nolock) "
        vnQuery += vbCrLf & "       inner join Sys_SsoWmsCompany_MA cmp with(nolock) on cmp.CompanyCode=mc.CompanyCode"
        vnQuery += vbCrLf & ")tb "
        vnQuery += vbCrLf & "order by case when CompanyCode='' then '' else CompanyCode end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CompanyCode"
            vriDst.DataTextField = "vCompanyCode"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillCstCompanyByUser(vriUserOID As String, vriDst As CheckBoxList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnUserCompanyCode As String
        vnQuery = "Select UserCompanyCode From Sys_SsoUser_MA Where OID=" & vriUserOID
        vnUserCompanyCode = fbuGetDataStrSQL(vnQuery, vriSQLConn)

        If vnUserCompanyCode = "" Then
            vnQuery = "Select * From ("
            If vriAll Then
                vnQuery += vbCrLf & "Select '' CompanyCode,'ALL' vCompanyCode UNION"
            Else
                vnQuery += vbCrLf & "Select '' CompanyCode,'' vCompanyCode UNION"
            End If
            vnQuery += vbCrLf & "Select mc.CompanyCode,mc.CompanyCode vCompanyCode"
            vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "DimCompany mc with(nolock) "
            vnQuery += vbCrLf & "       inner join Sys_SsoWmsCompany_MA cmp with(nolock) on cmp.CompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & " Where mc.CompanyCode in('" & stuCompanyCode.BAD & "','" & stuCompanyCode.CAS & "'))tb "
            vnQuery += vbCrLf & "order by case when CompanyCode='' then '' else CompanyCode end"

        Else
            vnQuery = "Select mc.CompanyCode,mc.CompanyCode vCompanyCode"
            vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "DimCompany mc with(nolock)"
            vnQuery += vbCrLf & "       inner join Sys_SsoWmsCompany_MA cmp with(nolock) on cmp.CompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & " Where UserOID=" & vriUserOID
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CompanyCode"
            vriDst.DataTextField = "vCompanyCode"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillCstCompanyByUser_20231013_Orig(vriUserOID As String, vriDst As CheckBoxList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String

        Dim vnUserCompanyCode As String
        vnQuery = "Select UserCompanyCode From Sys_SsoUser_MA Where OID=" & vriUserOID
        vnUserCompanyCode = fbuGetDataStrSQL(vnQuery, vriSQLConn)

        If vnUserCompanyCode = "" Then
            vnQuery = "Select * From ("
            If vriAll Then
                vnQuery += vbCrLf & "Select '' CompanyCode,'ALL' vCompanyCode UNION"
            Else
                vnQuery += vbCrLf & "Select '' CompanyCode,'' vCompanyCode UNION"
            End If
            vnQuery += vbCrLf & "Select mc.CompanyCode,mc.CompanyCode vCompanyCode"
            vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "DimCompany mc with(nolock) "
            vnQuery += vbCrLf & "       inner jon Sys_SsoSOHeader_TR soh with(nolock) on soh.SOCompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & " Where mc.CompanyCode in('" & stuCompanyCode.BAD & "','" & stuCompanyCode.CAS & "'))tb "
            vnQuery += vbCrLf & "order by case when CompanyCode='' then '' else CompanyCode end"

        Else
            vnQuery = "Select mc.CompanyCode,mc.CompanyCode vCompanyCode"
            vnQuery += vbCrLf & "  From " & fbuGetDBMaster() & "DimCompany mc with(nolock)"
            vnQuery += vbCrLf & "       inner join Sys_SsoSOHeader_TR soh with(nolock) on soh.SOCompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & "       inner join Sys_SsoUserCompany_MA uc with(nolock) on uc.CompanyCode=mc.CompanyCode"
            vnQuery += vbCrLf & " Where UserOID=" & vriUserOID & " and mc.CompanyCode in('" & stuCompanyCode.BAD & "','" & stuCompanyCode.CAS & "')"
        End If
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "CompanyCode"
            vriDst.DataTextField = "vCompanyCode"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuFillCstWarehouse_ByUserOID(vriUserOID As String, vriCst As CheckBoxList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnUserWarehouseCode As String = HttpContext.Current.Session("UserWarehouseCode")

        Dim vnDtb As New DataTable
        Dim vnQuery As String

        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' WarehouseName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' WarehouseName UNION"
        End If
        vnQuery += vbCrLf & "Select wh.OID,wh.WarehouseName "
        vnQuery += vbCrLf & "	   From " & fbuGetDBMaster() & "Sys_Warehouse_MA wh with(nolock)"
        vnQuery += vbCrLf & "	         inner join Sys_SsoWmsWarehouse_MA wwh with(nolock) on wwh.WarehouseOID=wh.OID"

        If vnUserWarehouseCode <> "" Then
            vnQuery += vbCrLf & "	         inner join Sys_SsoUserWarehouse_MA uw with(nolock) on uw.WarehouseOID=wh.OID and uw.UserOID=" & vriUserOID
        End If
        vnQuery += vbCrLf & ")tb order by case when OID=0 then '' else WarehouseName end"

        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriCst.DataSource = vnDtb
            vriCst.DataValueField = "OID"
            vriCst.DataTextField = "WarehouseName"
            vriCst.DataBind()
            vriCst.SelectedIndex = -1
        End If
    End Sub


End Module
