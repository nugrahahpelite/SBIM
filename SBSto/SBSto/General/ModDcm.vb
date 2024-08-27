Imports System.Data.SqlClient
Module ModDcm
    Public Enum enuSchDType
        Invoice = 1
        Retur = 2
        SJ = 3
        PL = 4
        TRB = 5
        Invoice_Only = 6
        Perintah_Kirim_DO_Titip = 7
        DO_Titip = 8
    End Enum

    Public Enum enuTCNotaDOT
        Cancelled = -2
        Baru = 0
        Prepared = 2
    End Enum
    Public Enum enuTCPerintahKirimDOT
        Cancelled = -2
        Baru = 0
        Prepared = 2
        Dalam_Picklist = 11
        Picklist_Done = 12
    End Enum
    Public Sub pbuFillDstDcmGudang(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' LocationName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' LocationName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,LocationName From " & vnDBDcm & "Sys_DcmLocation_MA with(nolock))tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else LocationName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "LocationName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

    Public Sub pbuInsertNotaByPrepare(vriTextStream As Scripting.TextStream, vriCompanyCode As String, vriNotaNo As String, ByRef vriNotaHOID As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<-------------------pbuInsertNotaByPrepare")

        Dim vnDBDcm As String = fbuGetDBDcm()

        Dim vnQuery As String
        Dim vnQueryFrom As String
        Dim vnCriteria As String
        Dim vnNotaHOID As Integer

        vnQuery = "Select OID From " & vnDBDcm & "Sys_DcmNotaHeader_TR with(nolock) Where CompanyCode='" & vriCompanyCode & "' and NotaNo='" & vriNotaNo & "'"
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("1")
        vriTextStream.WriteLine("vnQuery")
        vriTextStream.WriteLine(vnQuery)
        vnNotaHOID = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)
        If vnNotaHOID = 0 Then
            vnCriteria = " Where CompanyCode='" & vriCompanyCode & "' and NO_NOTA='" & vriNotaNo & "'"

            vnQueryFrom = vbCrLf & "  From " & vnDBDcm & "Sys_DcmJUAL"

            vnQuery = "Select isnull(max(OID),0)+1 From " & vnDBDcm & "Sys_DcmNotaHeader_TR with(nolock)"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("2")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            vnNotaHOID = fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into " & vnDBDcm & "Sys_DcmNotaHeader_TR"
            vnQuery += vbCrLf & "(OID,CompanyCode,NotaNo,NotaDate,CustCode,CustName,CustAddress,CustCity,GDG,WarehouseOID,DcmLocationOID,UploadDatetime)"
            vnQuery += vbCrLf & "Select Top 1 " & vnNotaHOID & " OID,CompanyCode,NO_NOTA,TANGGAL,KODE_CUST,CUSTOMER,ALAMAT,KOTA,GDG,WarehouseOID,0,UploadDatetime"
            vnQuery += vbCrLf & vnQueryFrom & vnCriteria
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("3")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)

            vnQuery = "Insert into " & vnDBDcm & "Sys_DcmNotaDetail_TR"
            vnQuery += vbCrLf & "(NotaHOID,KodeBarang,NamaBarang,Price,Qty,QtyBonus,Satuan,NoRef,Salesman)"
            vnQuery += vbCrLf & "Select " & vnNotaHOID & ",KODE_BARANG,NAMA_BARANG,isnull(HARGA,0),QTY,QTYBONUS,SATUAN,NO_REF,SALESMAN" & vnQueryFrom & vnCriteria
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("4")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End If

        vriNotaHOID = vnNotaHOID
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("vnNotaHOID = " & vnNotaHOID)
        vriTextStream.WriteLine("vriNotaHOID = " & vriNotaHOID)

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<<==================pbuInsertNotaByPrepare")
    End Sub

    Public Sub pbuInsertNotaDetail_ByBarang(vriTextStream As Scripting.TextStream, vriCompanyCode As String, vriNotaHOID As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<-------------------pbuInsertNotaDetail_ByBarang")

        Dim vnDBDcm As String = fbuGetDBDcm()
        Dim vnDBMaster As String = fbuGetDBMaster()

        Dim vnQuery As String
        vnQuery = "Select count(1) From " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR with(nolock) Where NotaHOID=" & vriNotaHOID
        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("1")
        vriTextStream.WriteLine("vnQuery")
        vriTextStream.WriteLine(vnQuery)
        If fbuGetDataNumSQLTrans(vnQuery, vriSQLConn, vriSQLTrans) > 0 Then
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("2")
            vriTextStream.WriteLine("NotaDetail_ByBarang_TR Already Exist")

        Else
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("3")
            vriTextStream.WriteLine("NotaDetail_ByBarang_TR NOT Already Exist")

            vnQuery = "Insert into " & vnDBDcm & "Sys_DcmNotaDetail_ByBarang_TR"
            vnQuery += vbCrLf & "(NotaHOID,KodeBarang,TotalQty,TotalQtyBonus)"
            vnQuery += vbCrLf & "Select " & vriNotaHOID & ",KodeBarang,sum(vTotalQty),sum(vTotalQtyBonus)"
            vnQuery += vbCrLf & "  From("
            vnQuery += vbCrLf & "       Select KodeBarang,Sum(Qty)vTotalQty,Sum(QtyBonus)vTotalQtyBonus"
            vnQuery += vbCrLf & "         From " & vnDBDcm & "Sys_DcmNotaDetail_TR with(nolock)"
            vnQuery += vbCrLf & "		 Where NotaHOID=" & vriNotaHOID & " and NOT KodeBarang in(Select b.PAKETCODE From " & vnDBMaster & "Sys_MstPaketH_MA b with(nolock) Where rtrim(b.CompanyCode)=rtrim('" & vriCompanyCode & "'))"
            vnQuery += vbCrLf & "        Group by KodeBarang"
            vnQuery += vbCrLf & "       UNION"
            vnQuery += vbCrLf & "       Select pd.BRGCODE,Sum(nh.Qty)*pd.PaketQty vTotalQty,sum(nh.QtyBonus)*pd.PaketQty vTotalQtyBonus"
            vnQuery += vbCrLf & "         From " & vnDBDcm & "Sys_DcmNotaDetail_TR nh with(nolock)"
            vnQuery += vbCrLf & "              inner join " & vnDBMaster & "Sys_MstPaketH_MA ph with(nolock) on ph.PAKETCODE=nh.KodeBarang"
            vnQuery += vbCrLf & "              inner join " & vnDBMaster & "Sys_MstPaketD_MA pd with(nolock) on pd.PAKETHOID=ph.OID"
            vnQuery += vbCrLf & "        Where nh.NotaHOID=" & vriNotaHOID
            vnQuery += vbCrLf & "        Group by pd.BRGCODE,pd.PaketQty"
            vnQuery += vbCrLf & "		)tb"
            vnQuery += vbCrLf & "  Group by KodeBarang"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("4")
            vriTextStream.WriteLine("vnQuery")
            vriTextStream.WriteLine(vnQuery)
            pbuExecuteSQLTrans(vnQuery, cbuActionNew, vriSQLConn, vriSQLTrans)
        End If

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("<<==================pbuInsertNotaDetail_ByBarang")
    End Sub

    Public Sub pbuFillDstDcmDriver(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' DcmDriverName UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' DcmDriverName UNION"
        End If
        vnQuery += vbCrLf & "Select OID,DcmDriverName From " & fbuGetDBDcm() & "Sys_DcmDriver_MA)tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else DcmDriverName end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "DcmDriverName"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub
    Public Sub pbuFillDstDcmVehicle(vriDst As DropDownList, vriAll As Boolean, vriSQLConn As SqlClient.SqlConnection)
        Dim vnDtb As New DataTable
        Dim vnQuery As String
        vnQuery = "Select * From ("
        If vriAll Then
            vnQuery += vbCrLf & "Select 0 OID,'ALL' VehicleNo UNION"
        Else
            vnQuery += vbCrLf & "Select 0 OID,'' VehicleNo UNION"
        End If
        vnQuery += vbCrLf & "Select OID,VehicleNo From " & fbuGetDBDcm() & "Sys_DcmVehicle_MA)tb "
        vnQuery += vbCrLf & "order by case when OID=0 then '' else VehicleNo end"
        pbuFillDtbSQL(vnDtb, vnQuery, vriSQLConn)

        If vnDtb.Rows.Count > 0 Then
            vriDst.DataSource = vnDtb
            vriDst.DataValueField = "OID"
            vriDst.DataTextField = "VehicleNo"
            vriDst.DataBind()
            vriDst.SelectedIndex = -1
        End If
    End Sub

End Module
