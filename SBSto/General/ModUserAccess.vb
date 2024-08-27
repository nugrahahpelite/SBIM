Imports System.Data.SqlClient
Module ModUserAccess
    Public Structure stuTransCode
        Const SsoStockOB = "STOB"
        Const SsoSSOH = "SSOH"
        Const SsoSSOC = "SSOC"
        Const SsoReceiving = "SRCV"
        Const SsoPicking_Old = "SPCK"
        Const SsoPOPembelian = "SPPO"
        Const SsoCustomerSO = "CSSO"
        Const SsoPOPackingList = "PLSP"
        Const SsoSummaryTRB = "SMTB"

        Const SsoPenerimaanPembelian = "RCPO"
        Const SsoPenerimaanLain2 = "RCMS"
        Const SsoPenerimaanKarantina = "RCKR"
        Const SsoPenerimaanRetur = "RCRT"

        Const SsoAdjustmentMinus = "AJMN"

        Const SsoPutaway = "PWAY"
        Const SsoPutaway_Antar_Wh = "PYAY"
        Const SsoPutaway_Karantina = "PTKW"
        Const SsoPutaway_DO_Titip = "PDTW"
        Const SsoPutaway_DO_Titip_Antar_Wh = "PDTY"
        Const SsoPutaway_Penerimaan_Dispatch = "PDSW"
        Const SsoPutaway_Penerimaan_Dispatch_Antar_Wh = "PDSY"
        Const SsoPindahLokasi = "PDLK"
        Const SsoPindahLokasi_Antar_Wh = "PDLW"
        Const SsoStorageStock = "SGST"
        Const SsoStockKarantina = "STKR"
        Const SsoPickList = "PICK"
        Const SsoPicking = "PCKG"
        Const SsoDispatch = "DISP"
        Const SsoDispatchReceive = "DISR"
        Const SsoDispatchReceive_Picking_Status = "DISG"
        Const SsoMoving_Antar_StagingOut = "DSGO"
        Const DcmPDOT = "PDOT"

        Const SsoMsBarang = "MSBG"
        Const SsoMsCustomer = "MSCS"
        Const SsoMsGudang = "MSGD"
        Const SsoPrintQRBarang = "PRQR"
        Const SsoPrintSN = "PRSN"
        Const SsoRequestSN = "RQSN"
        Const SsoChangeSKU = "CSKU"
        Const SsoVoidSO = "VOSO"
    End Structure

    Public Structure stuTrAccessCode
        Const Create_EditDel = "CED"
        Const Close_Trans = "CLO"
        Const Cancel_Trans = "CNC"
        Const Print = "PRN"
        Const Scan_Close = "SCC"
        Const Scan_QR = "SCN"
        Const Scan_Open = "SCO"
        Const View_Data = "VIW"
        Const Void_Trans = "VOI"

        Const Upload_Xls = "UPX"
        Const Prepare = "PRP"
        Const Approve = "APP"
        Const Stagging_In_Start = "SGS"
        Const Stagging_In_Finish = "SGF"
    End Structure
    Public Function fbuGetDtbUGAccess(vriTCode As String, vriSQLConn As SqlConnection) As DataTable
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From Sys_SsoUserGroupAccess_MA Where UserGroupOID=" & HttpContext.Current.Session("UserGroup") & " and TransCode='" & vriTCode & "'"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)
        Return vnDtb
    End Function

    Public Function fbuValAccess(vriDtb As DataTable, vriAcc As String) As Boolean
        If HttpContext.Current.Session("UserGroup") Is Nothing Then HttpContext.Current.Response.Redirect("~/Default.aspx")
        If CStr(HttpContext.Current.Session("UserGroup")) = "" Then HttpContext.Current.Response.Redirect("~/Default.aspx")

        Dim vn As Integer
        For vn = 0 To vriDtb.Rows.Count - 1
            If vriDtb.Rows(vn).Item("TrAccessCode") = vriAcc Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
