<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSummaryStock.aspx.vb" Inherits="SBSto.WbfSsoSummaryStock" MasterPageFile="~/SBSto.Master" Title="SB WMS : Summary Stock" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>
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

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
      <%--      $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
         <%--       $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
            }
        });
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            <%--document.getElementById("<%= LblProgress.ClientID%>").innerText = "Proses Tampil Data...";--%>
        }
    </script>
</head>
    <body>
        
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnListFind" />
                <asp:PostBackTrigger ControlID="BtnXLS" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>SUMMARY STOCK</strong></td>                                
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px">&nbsp;</td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="display:block;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                   <table>
                                                            <tr>
                                                                <td style="width:75px">Warehouse</td>
                                                                <td>:</td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            <td>
                                                                                <asp:DropDownList ID="DstListWarehouse" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                                                </asp:DropDownList>
                                                                            </td>
                                                                            <td>
                                                                                <asp:Label ID="LblMsgListWarehouse" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                                <td rowspan="2">
                                                                    <asp:Button ID="BtnListFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                                    &nbsp;&nbsp;&nbsp;
                                                                    <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                                    <asp:Label ID="LblMsgError" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>Company</td>
                                                                <td>:</td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            <td>
                                                                                <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="350px">
                                                                                </asp:DropDownList>
                                                                            </td>
                                                                            <td>
                                                                                <asp:Label ID="LblMsgListCompany" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>Tanggal</td>
                                                                <td>:</td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            <td>
                                                                                <asp:TextBox ID="TxtListStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                            </td>
                                                                            <td>
                                                                                <asp:Label ID="LblMsgListStart" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                            </td>
                                                                            <td>&nbsp;</td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                        
                                                    
                                                                <td>
                                                                    <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />

                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>                                        
                                            </tr>
                                        </table>
                                        <table>
                                            <tr>
                                                <td>
                                                    <div id="DivStCard" runat="server" >
                                                        <asp:Panel ID="PanStCard" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;width:585px;height:580px;margin-left:100px;margin-top:100px" Visible="True" BorderStyle="Solid" BackColor="White">
                                                            <table style="width:100%;font-family: tahoma;font-size:11px">
                                                                <tr>
                                                                    <td>
                                                                        <asp:Label ID="LblStCardTitle" runat="server" Font-Size="17px" Font-Bold="True">STOCK CARD</asp:Label>
                                                                    </td>
                                                                    <td style="text-align:right">
                                                                        <asp:Button ID="BtnStCardClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2">
                                                                        <asp:Panel ID="PanStCard_D" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;height:500px;" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                                            <asp:GridView ID="GrvStCard" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader" >
                                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                <Columns>
                                                                                    <asp:BoundField DataField="OID" HeaderText="OID" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" CssClass="myDisplayNone" />
                                                                                        <ItemStyle Width="55px" CssClass="myDisplayNone" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Location" HtmlEncode="false">
                                                                                        <HeaderStyle Width="100px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="TransCode" HeaderText="TransCode" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="TransName" HeaderText="Transaksi" HtmlEncode="false">
                                                                                        <HeaderStyle Width="100px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="TransOID" HeaderText="Transaksi<br />OID" HtmlEncode="false">
                                                                                        <HeaderStyle Width="50px" />
                                                                                        <ItemStyle HorizontalAlign="Center"/>
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="vCreationDatetime" HeaderText="Creation" HtmlEncode="false">
                                                                                        <HeaderStyle Width="145px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="TransQty" HeaderText="Qty Transaksi" DataFormatString="{0:n0}" >
                                                                                        <HeaderStyle Width="90px" />
                                                                                        <ItemStyle HorizontalAlign="Center" />
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
                                                                        </asp:Panel>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </asp:Panel>
                                                    </div>
                                                         <div id="DivTRB" runat="server" >
                                                        <asp:Panel ID="PanTRB" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;width:585px;height:580px;margin-left:100px;margin-top:100px" Visible="True" BorderStyle="Solid" BackColor="White">
                                                            <table style="width:100%;font-family: tahoma;font-size:11px">
                                                                <tr>
                                                                    <td>
                                                                        <asp:Label ID="LblTRBTitle" runat="server" Font-Size="17px" Font-Bold="True">TRB BELUM PICKING DONE</asp:Label>
                                                                    </td>
                                                                    <td style="text-align:right">
                                                                        <asp:Button ID="BtnTRBClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2">
                                                                        <asp:Panel ID="Panel2" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;height:500px;" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                                            <asp:GridView ID="GrvTRB" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader" >
                                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                <Columns>
                                                                                    <asp:BoundField DataField="CompanyCode" HeaderText="Company" HtmlEncode="false">
                                                                                        <HeaderStyle Width="100px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="WarehouseAsalOID" HeaderText="Warehouse" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" />
                                                                                    </asp:BoundField>
                                                                                            <asp:BoundField DataField="KodeBrg" HeaderText="Kode Barang" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" />
                                                                                    </asp:BoundField>
                                                                                            <asp:BoundField DataField="vQtyTRB_Belum_PickingDone" HeaderText="Qty" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" />
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
                                                                        </asp:Panel>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </asp:Panel>
                                                    </div>
                                                    <div id="DivInv" runat="server" >
                                                        <asp:Panel ID="PanInv" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;width:585px;height:580px;margin-left:100px;margin-top:100px" Visible="True" BorderStyle="Solid" BackColor="White">
                                                            <table style="width:100%;font-family: tahoma;font-size:11px">
                                                                <tr>
                                                                    <td>
                                                                        <asp:Label ID="LblInvTitle" runat="server" Font-Size="17px" Font-Bold="True">INVOICE BELUM PICKING DONE</asp:Label>
                                                                    </td>
                                                                    <td style="text-align:right">
                                                                        <asp:Button ID="BtnInvClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2">
                                                                        <asp:Panel ID="PanInv_D" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;height:500px;" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                                            <asp:GridView ID="GrvInv" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader" >
                                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                <Columns>
                                                                                    <asp:BoundField DataField="NO_NOTA" HeaderText="Invoice" HtmlEncode="false">
                                                                                        <HeaderStyle Width="100px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="vTANGGAL" HeaderText="Tanggal" HtmlEncode="false">
                                                                                        <HeaderStyle Width="55px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="vCustomer" HeaderText="Customer" HtmlEncode="false">
                                                                                        <HeaderStyle Width="100px" />
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="QTY" HeaderText="Qty" HtmlEncode="false">
                                                                                        <HeaderStyle Width="65px" />
                                                                                        <ItemStyle HorizontalAlign="Center"/>
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="QTYBONUS" HeaderText="Qty Bonus" HtmlEncode="false">
                                                                                        <HeaderStyle Width="65px" />
                                                                                        <ItemStyle HorizontalAlign="Center"/>
                                                                                    </asp:BoundField>
                                                                                    <asp:BoundField DataField="vUploadDatetime" HeaderText="Upload" HtmlEncode="false">
                                                                                        <HeaderStyle Width="145px" />
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
                                                                        </asp:Panel>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </asp:Panel>
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr style="vertical-align:top">
                                                <td>
                                                    <div style="height:525px; overflow:auto; border:ridge">
                                                        <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader" Width="1122px">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <Columns>
                                                                <asp:BoundField DataField="CompanyCode" HeaderText="Company" HtmlEncode="false" >
                                                                    <HeaderStyle Width="65px" CssClass="myDisplayNone" />
                                                                    <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="WarehouseOID" HeaderText="WarehouseOID" HtmlEncode="false" >
                                                                    <HeaderStyle Width="85px" CssClass="myDisplayNone" />
                                                                    <ItemStyle CssClass="myDisplayNone" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" HtmlEncode="false" >
                                                                    <HeaderStyle Width="85px" />
                                                                </asp:BoundField>
                                                                 <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false" >
                                                                    <HeaderStyle HorizontalAlign="Center" Width="85px" />
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                </asp:BoundField>                                                          
                                                                <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false" >
                                                                    <HeaderStyle HorizontalAlign="Center" Width="85px" />
                                                                </asp:BoundField>
                                                                <asp:ButtonField CommandName="vQty_StockCard" DataTextField="vQty_StockCard" Text="Button" HeaderText="Quantity<br />Stock Card" datatextformatstring="{0:n0}" >
                                                                    <HeaderStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="White" />
                                                                    <ItemStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                                </asp:ButtonField>
                                                                <asp:ButtonField CommandName="vQtyInv_Belum_PickingDone" DataTextField="vQtyInv_Belum_PickingDone" Text="Button" HeaderText="Quantity<br />Invoice Belum<br />Picking Done" datatextformatstring="{0:n0}" >
                                                                    <HeaderStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="White" />
                                                                    <ItemStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                                </asp:ButtonField>
                                                                <asp:ButtonField CommandName="vQtyTRB_Belum_PickingDone" DataTextField="vQtyTRB_Belum_PickingDone" Text="Button" HeaderText="Quantity<br />TRB Belum<br />Picking Done" datatextformatstring="{0:n0}" >
                                                                    <HeaderStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="White" />
                                                                    <ItemStyle Width="100px" HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                                </asp:ButtonField>
                                                                <asp:BoundField DataField="vQty_StockCard_Winacc" HeaderText="Quantity<br />Akhir" HtmlEncode="false" >
                                                                    <HeaderStyle HorizontalAlign="Center" Width="85px" />
                                                                    <ItemStyle HorizontalAlign="Center" />
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
                                                    </div>
                                                </td>                                               
                                            </tr>
                                            <tr>
                                                <td></td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>        
    </body>
</asp:Content>
