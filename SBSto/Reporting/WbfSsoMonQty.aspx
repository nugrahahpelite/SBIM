<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoMonQty.aspx.vb" Inherits="SBSto.WbfSsoMonQty" MasterPageFile="~/SBSto.Master" Title="SB WMS : Monitor Quantity Stock" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Monitor Quantity Stock</title>
</head>
    <body>
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td style="font-size:15px;height:28px" colspan="3"><strong>MONITOR QUANTITY STOCK</strong></td>
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
                                                    <td></td>
                                                    <td>Building</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListBuilding" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td></td>
                                                    <td rowspan="2">
                                                        <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
                                                    <td></td>
                                                    <td>Zona</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListZona" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                </tr>
                                                <tr style="vertical-align:top">
                                                    <td>Storage Type</td>
                                                    <td>:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListStorageType" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td></td>
                                                    <td colspan="3">
                                                        <table>
                                                            <tr>
                                                                <td style="vertical-align:top">
                                                                    <asp:Panel ID="PanListRackN" runat="server" Visible="false">
                                                                        <table>
                                                                            <tr>
                                                                                <td>Storage Number</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackN_Start" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                </td>
                                                                                <td>s/d</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackN_End" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                    <asp:Panel ID="PanListRackY" runat="server" Visible="false">
                                                                        <table>
                                                                            <tr>
                                                                                <td>SequenceNo.Column.Level</td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_SeqNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="40px"></asp:TextBox>
                                                                                </td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_Column" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="40px"></asp:TextBox>
                                                                                </td>
                                                                                <td>
                                                                                    <asp:TextBox ID="TxtListRackY_Level" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="40px"></asp:TextBox>
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
                                                        <asp:TextBox ID="TxtListBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="20" Width="95px"></asp:TextBox>
                                                        <asp:TextBox ID="TxtListBrgName" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="190px"></asp:TextBox>
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
                                                        <asp:TextBox ID="TxtListRcvNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="190px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnListRcvNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                    </td>
                                                    <td></td>
                                                    <td>
                                                        &nbsp;</td>
                                                    <td>
                                                        &nbsp;</td>
                                                    <td>
                                                        &nbsp;</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:CheckBox ID="ChkVarianOnly" runat="server" Checked="True" Font-Size="10pt" ForeColor="Blue" Text="Tampilkan HANYA Data Selisih" />
                                                    </td>
                                                    <td></td>
                                                    <td>
                                                        <asp:DropDownList ID="DstMonQty" runat="server" style="height: 20px" Width="350px">
                                                        </asp:DropDownList>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red"></asp:Label>
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
                                                                        <asp:BoundField DataField="RcvPOTypeName" HeaderText="PO Type" HtmlEncode="false">
                                                                            <HeaderStyle Width="100px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="OID" HeaderText="OID" HtmlEncode="false">
                                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                                            <ItemStyle CssClass="myDisplayNone" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="RcvPORefTypeOID" HeaderText="RcvPORefTypeOID" HtmlEncode="false">
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
                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <%-- 0 --%>
                                        <asp:BoundField DataField="vStorageOID" HeaderText="Storage<br />OID" HtmlEncode="false">
                                            <HeaderStyle Width="55px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 1 --%>
                                        <asp:BoundField DataField="vStorageStockOID" HeaderText="Storage Stock<br />OID" HtmlEncode="false">
                                            <HeaderStyle Width="55px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 2 --%>
                                        <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Location" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <%-- 3 --%>
                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company" HtmlEncode="false">
                                            <HeaderStyle Width="50px" />
                                        </asp:BoundField>
                                        <%-- 4 --%>
                                        <asp:BoundField DataField="RcvPONo" HeaderText="No.<br />Terima" HtmlEncode="false">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <%-- 5 --%>
                                        <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal<br />Terima" HtmlEncode="false">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <%-- 6 --%>
                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode<br />Barang" HtmlEncode="false">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <%-- 7 --%>
                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                            <HeaderStyle Width="220px" />
                                        </asp:BoundField>
                                        <%-- 8 --%>
                                        <asp:BoundField DataField="QtyOnHand" HeaderText="Qty<br />On Hand" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 9 --%>
                                        <asp:BoundField DataField="vQtyStockCard" HeaderText="Qty<br />Stock Card" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 10 --%>
                                        <asp:BoundField DataField="QtyOnPutaway" HeaderText="Qty<br />On Putaway" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 11 --%>
                                        <asp:BoundField DataField="vQtyOnPutaway_Trans" HeaderText="Qty<br />On Putaway<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 12 --%>
                                        <asp:BoundField DataField="QtyOnPutawayWh" HeaderText="Qty<br />On Putaway<br />Antar Wh" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 13 --%>
                                        <asp:BoundField DataField="vQtyOnPutawayWh_Trans" HeaderText="Qty<br />On Putaway<br />Antar Wh<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 14 --%>
                                        <asp:BoundField DataField="QtyOnMovement" HeaderText="Qty<br />On Movement" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 15 --%>
                                        <asp:BoundField DataField="vQtyOnMovement_Trans" HeaderText="Qty<br />On Movement<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 16 --%>
                                        <asp:BoundField DataField="QtyOnMovementWh" HeaderText="Qty<br />On Movement<br />Antar Wh" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 17 --%>
                                        <asp:BoundField DataField="vQtyOnMovementWh_Trans" HeaderText="Qty<br />On Movement<br />Antar Wh<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 18 --%>
                                        <asp:BoundField DataField="QtyOnPickList" HeaderText="Qty<br />On Picklist" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 19 --%>
                                        <asp:BoundField DataField="vQtyOnPickList_Trans" HeaderText="Qty<br />On Picklist<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 20 --%>
                                        <asp:BoundField DataField="QtyOnPicking" HeaderText="Qty<br />Picking" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 21 --%>
                                        <asp:BoundField DataField="vQtyOnPicking_Trans" HeaderText="Qty<br />Picking<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 22 --%>
                                        <asp:BoundField DataField="QtyOnDispatch" HeaderText="Qty<br />On Dispatch" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 23 --%>
                                        <asp:BoundField DataField="vQtyOnDispatch_Trans" HeaderText="Qty<br />On Dispatch<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 24 --%>
                                        <asp:BoundField DataField="QtyOnKarantina" HeaderText="Qty<br />On Karantina" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 25 --%>
                                        <asp:BoundField DataField="vQtyOnKarantina_Trans" HeaderText="Qty<br />On Karantina<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 26 --%>
                                        <asp:BoundField DataField="QtyOnPutawayKr" HeaderText="Qty Karantina<br />On Putaway" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 27 --%>
                                        <asp:BoundField DataField="vQtyOnPutawayKr_Trans" HeaderText="Qty Karantina<br />On Putaway<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 28 --%>
                                        <asp:BoundField DataField="QtyOnPutawayDtw" HeaderText="Qty DO Titip<br />On Putaway" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 29 --%>
                                        <asp:BoundField DataField="vQtyOnPutawayDtw_Trans" HeaderText="Qty DO Titip<br />On Putaway<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 30 --%>
                                        <asp:BoundField DataField="QtyOnPutawayDty" HeaderText="Qty DO Titip<br />On Putaway<br >Antar Wh" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 31 --%>
                                        <asp:BoundField DataField="vQtyOnPutawayDty_Trans" HeaderText="Qty DO Titip<br />On Putaway<br >Antar Wh<br />Trans" DataFormatString="{0:#,##}" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
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
    </asp:Content>
