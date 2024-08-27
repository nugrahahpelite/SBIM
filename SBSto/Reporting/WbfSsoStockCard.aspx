<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoStockCard.aspx.vb" Inherits="SBSto.WbfSsoStockCard" MasterPageFile="~/SBSto.Master" Title="SB WMS : Stock Card" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Stock Card</title>
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
</head>
    <body>
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnXLS" />
            </Triggers>     
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td style="font-size:15px;height:28px" colspan="3"><strong>STOCK CARD</strong></td>
                        </tr>
                        <tr>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="width:85px">Warehouse</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListWarehouse" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListWarehouse" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td>Building</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListBuilding" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListBuilding" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkStorageOID" runat="server" AutoPostBack="True" ForeColor="Blue" Text="By Storage OID" />
                                                    </td>
                                                    <td rowspan="2">
                                                        <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        &nbsp;&nbsp;
                                                        <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>Lantai</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListLantai" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListLantai" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td>Zona</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListZona" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListZona" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblStorageOID" runat="server" ForeColor="Blue" Visible="False">Storage OID</asp:Label>
                                                        <asp:TextBox ID="TxtStorageOID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Visible="False" Width="45px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>Storage Type</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListStorageType" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListStorageType" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td colspan="4">
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:Panel ID="PanListRackN" runat="server" Visible="false">
                                                                        <table>
                                                                            <tr>
                                                                                <td>Storage Number</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackN_Start" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                </td>
                                                                                <td class="auto-style2">
                                                                                    <asp:Label ID="LblMsgListRackN_Start" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                                </td>
                                                                                <td>
                                                                                    &nbsp;</td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                    <asp:Panel ID="PanListRackY" runat="server" Visible="false">
                                                                        <table>
                                                                            <tr>
                                                                                <td>SequenceNo</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_SeqNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="45px"></asp:TextBox>
                                                                                </td>
                                                                                <td>Column</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_Column" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="45px"></asp:TextBox>
                                                                                </td>
                                                                                <td>Level</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_Level" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="45px"></asp:TextBox>
                                                                                </td>
                                                                                <td>
                                                                                    <asp:Label ID="LblMsgListRackY" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                    <asp:Panel ID="PanListStagging" runat="server" Visible="false">
                                                                        <table>
                                                                            <tr>
                                                                                <td>Stagging</td>
                                                                                <td>:</td>
                                                                                <td>
                                                                                    <table>
                                                                                        <tr>
                                                                                            <td>
                                                                                                <asp:RadioButtonList ID="RdbListStagging" runat="server" RepeatDirection="Horizontal">
                                                                                                    <asp:ListItem Value="1" Selected="True">IN</asp:ListItem>
                                                                                                    <asp:ListItem Value="2">OUT</asp:ListItem>
                                                                                                </asp:RadioButtonList>
                                                                                            </td>
                                                                                            <td>
                                                                                                &nbsp;</td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td style="vertical-align:top">
                                                        <asp:CheckBox ID="ChkSummByWarehouse" runat="server" AutoPostBack="True" ForeColor="Blue" Text="By Warehouse" />
                                                    </td>
                                                    <td>

                                                        <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />
                                                        <asp:Label ID="LblProgressXLS" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>

                                                    </td>
                                                </tr>
                                            </table>
                                        </td>                                        
                                    </tr>
                                    <tr>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="width:85px">Company</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="350px">
                                                        </asp:DropDownList>
                                                        <asp:Label ID="LblMsgListCompany" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td></td>
                                                    <td>Barang</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtListBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="20" Width="95px" BackColor="#CCCCCC" ReadOnly="True"></asp:TextBox>
                                                        <asp:TextBox ID="TxtListBrgName" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="300px" BackColor="#CCCCCC" ReadOnly="True"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnListBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgListBrg" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="width:85px">No. Penerimaan</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtListRcvNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="190px" BackColor="#CCCCCC"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnListRcvNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                    </td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkSummAllRcv" runat="server" AutoPostBack="True" ForeColor="Blue" Text="All Penerimaan" />
                                                        <asp:Label ID="LblMsgListRcvNo" runat="server" ForeColor="#FF0066"></asp:Label>
                                                    </td>
                                                    <td style="width:200px"></td>
                                                    <td>&nbsp;</td>
                                                    <td>
                                                        &nbsp;
                                                        </td>
                                                    <td></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <div id="DivLsBrg" runat="server" >
                                                <asp:Panel ID="PanLsBrg" class="myPanelGreyNS" runat="server" style="display:block;width:525px;height:580px;margin-left:10px;margin-top:0px" Visible="True" BorderStyle="Solid" BackColor="White">
                                                    <table style="width:100%;font-family: tahoma;font-size:11px">
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="LblLsBrg" runat="server" Font-Size="17px" Font-Bold="True">DAFTAR BARANG</asp:Label>
                                                            </td>
                                                            <td style="text-align:right">
                                                                <asp:Button ID="BtnLsBrgClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <table>
                                                                    <tr>
                                                                        <td>Barang</td>
                                                                        <td>:</td>
                                                                        <td>                                                                
                                                                            <asp:TextBox ID="TxtLsBrg" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                                        </td>
                                                                        <td>
                                                                            <asp:Button ID="BtnLsBrg" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  TAMPILKAN DATA" Width="125px" Font-Bold="True" />
                                                                        </td>                                                            
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <asp:GridView ID="GrvLsBrg" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:ButtonField CommandName="Select" DataTextField="BRGCODE" HeaderText="Kode Barang">
                                                                        <HeaderStyle Width="100px" />
                                                                        </asp:ButtonField>
                                                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                                        <HeaderStyle Width="245px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" HtmlEncode="false">
                                                                        <HeaderStyle Width="100px" />
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
                                            <div id="DivLsRcvPO" runat="server" >
                                                <asp:Panel ID="PanLsRcvPO" class="myPanelGreyNS" runat="server" style="display:block;width:525px;height:580px;margin-left:10px;margin-top:0px" Visible="True" BorderStyle="Solid" BackColor="White">
                                                    <table style="width:100%;font-family: tahoma;font-size:11px">
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="Label1" runat="server" Font-Size="17px" Font-Bold="True">DAFTAR PENERIMAAN PEMBELIAN</asp:Label>
                                                            </td>
                                                            <td style="text-align:right">
                                                                <asp:Button ID="BtnLsRcvPOClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <table>
                                                                    <tr>
                                                                        <td>Nomor</td>
                                                                        <td>:</td>
                                                                        <td>                                                                
                                                                            <asp:TextBox ID="TxtLsRcvPONo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                                        </td>
                                                                        <td>
                                                                            <asp:Button ID="BtnLsRcvPOFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  TAMPILKAN DATA" Width="125px" Font-Bold="True" />
                                                                        </td>                                                            
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <asp:GridView ID="GrvLsRcvPO" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:ButtonField CommandName="Select" DataTextField="RcvPONo" HeaderText="No.Penerimaan">
                                                                            <HeaderStyle Width="120px" />
                                                                        </asp:ButtonField>
                                                                        <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal<br />Penerimaan" HtmlEncode="false">
                                                                            <HeaderStyle Width="90px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="RcvPOSupplierName" HeaderText="Supplier" HtmlEncode="false">
                                                                            <HeaderStyle Width="145px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="RcvTypeName" HeaderText="Jenis" HtmlEncode="false">
                                                                            <HeaderStyle Width="100px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="OID" HeaderText="OID" HtmlEncode="false">
                                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                                            <ItemStyle CssClass="myDisplayNone" />
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
                                        </td>
                                    </tr>
                                </table>
                            </td>                
                        </tr>                                        
                        <tr>
                            <td>
                                <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:700px;" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" HeaderStyle-CssClass="StickyHeader" >
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                <HeaderStyle Width="80px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BuildingName" HeaderText="Building">
                                                <HeaderStyle Width="70px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="LantaiDescription" HeaderText="Lantai">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ZonaName" HeaderText="Zona">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageTypeName" HeaderText="Storage Type">
                                                <HeaderStyle Width="55px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageSequenceNumber" HeaderText="Sequence<br />Number" HtmlEncode="false">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageColumn" HeaderText="Column" HtmlEncode="false">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageLevel" HeaderText="Level" HtmlEncode="false">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageNumber" HeaderText="Storage<br />Number" HtmlEncode="false">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStorageStagIO" HeaderText="Stagging" HtmlEncode="false">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStorageOID" HeaderText="Storage<br />OID" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStockCardOID" HeaderText="Stock Card<br />OID" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Location" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company" HtmlEncode="false">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="RcvPONo" HeaderText="No.<br />Terima" HtmlEncode="false">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal<br />Terima" HtmlEncode="false">
                                                <HeaderStyle Width="70px" CssClass="myDisplayNone"/>
                                                <ItemStyle CssClass="myDisplayNone"/>
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode<br />Barang" HtmlEncode="false">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="300px" />
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
                                            <asp:BoundField DataField="vQtyOnHand" HeaderText="Qty<br />On Hand" DataFormatString="{0:#,##}" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
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
            
                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red"/>
                                        <EmptyDataTemplate>Tidak Ada Data</EmptyDataTemplate>
                                    </asp:GridView>
                                </asp:Panel>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HdfListRcvPOHOID" runat="server" Value="0" />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>

            </ContentTemplate>
        </asp:UpdatePanel>
        </div>
    </body>
</html>
</asp:Content>
<asp:Content ID="Content2" runat="server" contentplaceholderid="head">
    <style type="text/css">
        .auto-style2 {
            width: 3px;
        }
    </style>
    </asp:Content>
