<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfDmPerintahDOTDisplay.aspx.vb" Inherits="SBSto.WbfDmPerintahDOTDisplay" Title="SB WMS - Perintah Kirim DO Titip" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <style type="text/css">=
        .gridViewPager td
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
        }
        .StickyHeader th {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            top: 0
        }
        .GrvDetail {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
        }
        .modalPopup
        {
            background-color: #f4f4f4;
            width: 250px;
            border: 1px solid rgba(0, 0, 0, 0.5);
            border-radius: 12px;
            padding: 0;
            box-shadow:0 5px 15px rgba(0, 0, 0, 0.5);
        }
        .Grid td
        {
            padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
            line-height:200%
        }
        .Grid th
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            line-height:200%
        }
        .ChildGrid td
        {
            background-color: azure !important;
            color: black;
            font-size: 10pt;
            line-height:200%
        }
        .ChildGrid th
        {
            background-color: cadetblue !important;
            color: White;
            font-size: 10pt;
            line-height:200%
        }
        .highlight 
        {
            text-decoration: none;
            color:black;
            background:yellow;
        }
        </style>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtPKDOTScheduleDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtPKDOTScheduleDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowConfirmCancel() {
            if ('<%=Session("UserOID")%>' == "") {
                window.location = "Default.aspx";
                return;
            }
            else {
                if (document.getElementById("<%= TxtTransID.ClientID%>").value == "") {
                    return;
                }
                document.getElementById("<%= DivConfirm.ClientID%>").style.visibility = "visible";
                document.getElementById("<%= LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Cancel Surat Jalan ?? WARNING : Cancel Tidak Dapat Dibatalkan";
                document.getElementById("<%= HdfProcess.ClientID%>").value = "Cancel";
            }
        }
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = document.getElementById("<%= HdfProcess.ClientID%>").value + " in Progress.........";
        }
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
        }

        function GetFullPathFileName(vriString) {
            var FName;
            var FNameFound = '';
            var step;
            FName = '';
            for (step = vriString.length; step > -1; step--) {
                if (vriString.substr(step, 1) == "\\") {
                    FName = vriString.substr(step + 1, vriString.length);
                    FNameFound = 'Y';
                    return FName;
                }
            }
            if (FNameFound == '') {
                FName = vriString;
            }
            return FName;
        }
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
        }
    </script>        
</head>
    <body>
        
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>PERINTAH KIRIM DO TITIP</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:140px">
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
                                </td>
                                <td style="width:45px">
                                    <asp:Button ID="BtnSummary" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Summary" Width="100px" />
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="z-index:110;display:block;width:1650px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListFind" >
                                        <table style="width:85%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td style="width:150px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="100px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressFind();" />
                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:70px;text-align:right">Trans. ID :</td>
                                                <td style="width:80px"><asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="10" Width="80px"></asp:TextBox></td>                                                
                                                <td  style="width:100px;text-align:right">No. Perintah Kirim : </td>
                                                <td style="width:120px">
                                                    <asp:TextBox ID="TxtListNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                </td>
                                                <td style="width:75px;text-align:right">No. Ref</td>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td style="width:100px;text-align:right">
                                                                <asp:TextBox ID="TxtListRefNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="165px"></asp:TextBox>
                                                            </td>
                                                            <td>Periode :</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                            </td>
                                                            <td>s/d</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td></td>
                                                <td style="text-align:right">Warehouse :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">Status PL :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Prepared" runat="server" ForeColor="#336600" Text="Prepared" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp; <asp:CheckBox ID="ChkSt_InPicklist" runat="server" ForeColor="#336600" Text="Dalam Picklist" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_PicklistDone" runat="server" ForeColor="#336600" Text="Picklist Done" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
                                                            </td>
                                                            <td style="width:25px"></td>
                                                            <td>
                                                                &nbsp;</td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <table style="width: 100%;margin:auto;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" PageSize="20">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="Select" DataTextField="PKDOTNo" Text="Button" HeaderText="Nomor Perintah Kirim">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vPKDOTDate" HeaderText="Tanggal<br />Perintah Kirim" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vPKDOTScheduleDate" HeaderText="Schedule<br />Perintah Kirim" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PKDOTCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang Asal"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vCustomer" HeaderText="Customer"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vIsExpedition" HeaderText="Ekspedisi"> 
                                                                <HeaderStyle Width="45px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vIsPickListClosed" HeaderText="Sudah<br />Picklist" HtmlEncode="false">
                                                                <HeaderStyle Width="55px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PKDOTShipToName" HeaderText="Ship To"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PKDOTNote" HeaderText="Note" >
                                                                <HeaderStyle Width="120px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vPrepared" HeaderText="Prepared" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" ForeColor="White" Height="35px" Font-Overline="False" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                        </table>                                        
                                    </asp:Panel>
                                </div>
                                <div id="DivConfirm" runat="server" style="text-align:center;width:100%">
                                    <asp:Panel ID="PanConfirm" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%">
                                        
                                        <table style="width:75%;margin:auto">
                                            <tr style="text-align:left">
                                                <td colspan="2">
                                                    <br />
                                                    <asp:Label ID="LblConfirmMessage" runat="server" Text="Anda Yakin ?" Font-Size="17px"></asp:Label>
                                                    <br />
                                                    <asp:Label ID="LblConfirmProgress" runat="server" Font-Size="17px"></asp:Label>
                                                    <br />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table runat="server" id="tbConfirmNote" visible="false">
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="LblConfirmNote" runat="server" Text="Note : "></asp:Label>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtConfirmNote" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="270px" Height="35px" TextMode="MultiLine"></asp:TextBox>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblConfirmWarning" runat="server" ForeColor="#FF0066"></asp:Label>                                                    
                                                </td>
                                            </tr>
                                            <tr style="height:65px">
                                                <td style="text-align:center">
                                                    <table>
                                                        <tr>
                                                            <td>
                                                                <asp:Button ID="BtnConfirmYes" runat="server" OnClientClick="fsDisableYesConfirmSto();" Text="Yes" Width="115px" Font-Bold="True" Height="35px" />
                                                            </td>
                                                            <td>
                                                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                <asp:Button ID="BtnConfirmNo" runat="server" CssClass="no" Text="No" Width="145px" Font-Bold="True" Height="35px" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <br />                                        
                                        <br />
                                        <br />
                                    </asp:Panel>
                                </div>
                                <div id="DivListItem" runat="server" >
                                    <asp:Panel ID="PanListItem" class="myPanelGreyNS" runat="server" style="display:block;width:650px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListItemFind" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListItem" runat="server" Font-Size="17px" Font-Bold="True">DAFTAR BARANG</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListItemClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>Kode/Nama Barang</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListItem" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnListItemFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Cari" Width="112px" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListItem" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="450px" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:ButtonField CommandName="ItemCode" DataTextField="ItemCode" Text="Button" HeaderText="Item Code">
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="ItemName" HeaderText="Item Name" HtmlEncode="false">
                                                                <HeaderStyle Width="350px" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                                        <EmptyDataTemplate>
                                                            Tidak Ada Data
                                                        </EmptyDataTemplate>
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                                <div id="DivListCustomer" runat="server" >
                                    <asp:Panel ID="PanListCustomer" class="myPanelGreyNS" runat="server" style="display:block;width:850px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListCustomerFind" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListCustomer" runat="server" Font-Size="17px" Font-Bold="True">DAFTAR CUSTOMER</asp:Label>
                                                </td>
                                                <td>
                                                    <asp:Button ID="BtnListCustomerClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>Kode/Nama Customer</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListCustomer" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnListCustomerFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Cari" Width="112px" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListCustomer" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="750px" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:ButtonField CommandName="kode_cust" DataTextField="kode_cust" Text="Button" HeaderText="Kode Customer">
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="customer" HeaderText="Nama Customer" HtmlEncode="false">
                                                                <HeaderStyle Width="350px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="ALAMAT" HeaderText="Alamat" HtmlEncode="false">
                                                                <HeaderStyle Width="285px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="kota" HeaderText="Kota">
                                                                <HeaderStyle Width="185px" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                                        <EmptyDataTemplate>
                                                            Tidak Ada Data
                                                        </EmptyDataTemplate>
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                                <div id="DivPreview" runat="server" >
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="z-index:104;display:block;width:75%;margin-top:-45px">
                                        <table>
                                            <tr style="vertical-align:top">
                                                <td style="width:98%">
                                                    <iframe runat="server" id="ifrPreview" style="width:100%;height:750px" ></iframe>
                                                </td>
                                                <td>
                                                    <asp:Button ID="BtnPreviewClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                        </table>                            
                                    </asp:Panel>
                                </div>
                                <div id="DivAttach" runat="server" >
                                    <asp:Panel ID="PanAttach" class="myPanelGreyNS" runat="server" style="display:block;width:875px;height:620px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblAttachTitle" runat="server" Font-Size="17px" Font-Bold="True">ATTACH PHOTO</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnAttachClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>Note</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtAttachNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="450" Width="450px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="3">
                                                                <asp:Panel ID="PanAttachImg" runat="server" Visible="false" style="text-align:center;width: 800px; height: 450px;" BackColor="White" BorderColor="#666666" BorderStyle="Solid">
                                                                    <br />
                                                                    <asp:Image ID="ImgAttachImg" runat="server" style="width:650px;height:350px" />
                                                                    <br />
                                                                </asp:Panel>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td style="width:80px">No PK</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPKDOTNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#DDDDDD"></asp:TextBox>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:85px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="300px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgCompany" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td style="width:25px"></td>
                            <td style="width:85px">ID Transaksi</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td style="width:195px">
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="73px" ReadOnly="True"></asp:TextBox>
                            </td>
                            <td style="width:125px"></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Tanggal PK</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPKDOTDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px" BackColor="#DDDDDD"></asp:TextBox>
                                <asp:Label ID="LblMsgPKDOTDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td></td>
                            <td>Gudang</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstWhs" runat="server" style="height: 20px" Width="300px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgWhs" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td></td>
                            <td>Status</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="125px" ReadOnly="True"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Schedule Kirim</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPKDOTScheduleDate" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgPKDOTScheduleDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td></td>
                            <td>Gudang Tujuan</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstWhsDest" runat="server" style="height: 20px" Width="300px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgWhsDest" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td></td>
                            <td>Attachment</td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Customer</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPKDOTCust" runat="server" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>&nbsp;</td>
                            <td style="text-align:center">&nbsp;</td>
                            <td>
                                <asp:CheckBox ID="ChkExpedition" runat="server" Font-Size="12pt" ForeColor="#0066FF" Text="Ekspedisi" />
                            </td>
                            <td></td>
                            <td colspan="5" rowspan="4" style="vertical-align:top">
                                <asp:GridView ID="GrvAttach" runat="server" AutoGenerateColumns="False" CellPadding="4" Font-Names="Tahoma" Font-Size="12px" ForeColor="#333333" ShowHeaderWhenEmpty="True">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                        <HeaderStyle CssClass="myDisplayNone" />
                                        <ItemStyle CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="PKDOTImgNote" DataTextField="PKDOTImgNote" HeaderText="Note">
                                        <HeaderStyle Width="125px" />
                                        </asp:ButtonField>
                                        <asp:BoundField DataField="vUploadDatetime" HeaderText="Upload">
                                        <HeaderStyle Width="200px" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="vUploadDel" DataTextField="vUploadDel" HeaderText="">
                                        <HeaderStyle Width="85px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                        </asp:ButtonField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                    <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                    <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                    <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                    <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                    <EmptyDataTemplate>
                                        -
                                    </EmptyDataTemplate>
                                </asp:GridView>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td style="vertical-align:top">&nbsp;</td>
                            <td style="text-align:center;vertical-align:top">&nbsp;</td>
                            <td></td>
                            <td>&nbsp;</td>
                            <td style="vertical-align:top">Ship To</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtPKDOTShipToName" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="220px"></asp:TextBox>
                                <asp:Label ID="LblMsgCust" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td style="vertical-align:top">Note</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtPKDOTNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="45px" MaxLength="245" TextMode="MultiLine" Width="245px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td style="vertical-align:top">Alamat</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtPKDOTShipToAddress" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="38px" MaxLength="245" TextMode="MultiLine" Width="300px"></asp:TextBox>
                            </td>
                            <td>
                                &nbsp;</td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="7">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td>                            
                        </tr>
                    </table>
                    <div style="width:177%; height:525px; overflow:auto; border:ridge">
                        <table>
                            <tr style="vertical-align:top">
                                <td style="width:10px"></td>
                                <td colspan="2">

                                    <asp:GridView ID="GrvDetail" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="OID" HeaderText="OID" HtmlEncode="false">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:ButtonField CommandName="vAddItem" DataTextField="vAddItem" Text="Button" HeaderText="">
                                                <HeaderStyle Width="45px" />
                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <%-- 2 --%>
                                            <asp:ButtonField CommandName="BRGCODE" DataTextField="BRGCODE" Text="Button" HeaderText="Kode Barang">
                                                <HeaderStyle Width="85px" />
                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <%-- 3 --%>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="245px" />
                                            </asp:BoundField>
                                            <%-- 4 --%>
                                            <asp:BoundField DataField="AvailableDOTQty" DataFormatString="{0:n0}" HeaderText="Qty Tersedia">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 5 --%>
                                            <asp:BoundField DataField="RequestQty" DataFormatString="{0:n0}" HeaderText="Qty Request">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 6 --%>
                                            <asp:TemplateField HeaderText="Qty Request">
                                                <ItemTemplate>
                                                    <asp:TextBox runat="server" ID="TxtRequestQty" Width="100px" MaxLength="10"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <%-- 7 --%>
                                            <asp:BoundField DataField="PKDOTDQty" DataFormatString="{0:n0}" HeaderText="Qty<br />Perintah Kirim" HtmlEncode="false">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 8 --%>
                                            <asp:BoundField DataField="QtyOnPickList" DataFormatString="{0:n0}" HeaderText="Qty<br />Picklist" HtmlEncode="false">
                                                <HeaderStyle Width="75px" CssClass="myDisplayNone" />
                                                <ItemStyle HorizontalAlign="Right" CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 9 --%>
                                            <asp:BoundField DataField="vMessageItem" HeaderText="Message">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle ForeColor="Red" />
                                            </asp:BoundField>
                                            <%-- 10 --%>
                                            <asp:ButtonField CommandName="vDelItem" DataTextField="vDelItem" Text="Button" HeaderText="">
                                                <HeaderStyle Width="65px" CssClass="myDisplayNone" />
                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" CssClass="myDisplayNone" />
                                            </asp:ButtonField>
                                        </Columns>
                                        <EditRowStyle BackColor="#999999" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                        <EmptyDataTemplate>
                                            Tidak Ada Data
                                        </EmptyDataTemplate>
                                    </asp:GridView>
                                </td>
                                <td>
                                    <asp:Panel ID="PanReserved" runat="server" style="height:525px">
                                        <table>
                                            <tr>
                                                <td class="auto-style2">
                                                    <asp:Label ID="LblMsgReserved" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:GridView ID="GrvReserved" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false">
                                                                <HeaderStyle Width="80px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NotaHOID" HeaderText="NotaHOID">
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NotaNo" HeaderText="No. Invoice">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaDate" HeaderText="Tanggal<br />Invoice" HtmlEncode="false">
                                                                <HeaderStyle Width="80px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PKDOTNQty" DataFormatString="{0:n0}" HeaderText="Qty<br />Kirim" HtmlEncode="false">
                                                                <HeaderStyle Width="55px" />
                                                                <ItemStyle HorizontalAlign="Right" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                                        <EmptyDataTemplate>
                                                            Tidak Ada Data
                                                        </EmptyDataTemplate>
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <table>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfTransStatus" runat="server" Value="0" /></td>                            
                            <td><asp:HiddenField ID="HdfProcess" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfActionStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfDetailOID" runat="server" /></td>
                            <td>
                                <asp:HiddenField ID="HdfAttachOID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:HiddenField ID="HdfCompanyCode" runat="server" Value="1" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfWhs" runat="server" Value="1" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfWhsDest" runat="server" Value="1" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfCustCode" runat="server" Value="1" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfFupAttach" runat="server" />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>
        
    </body>
</asp:Content>
<asp:Content ID="Content2" runat="server" contentplaceholderid="head">
    </asp:Content>