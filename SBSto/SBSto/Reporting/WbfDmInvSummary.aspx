<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfDmInvSummary.aspx.vb" Inherits="SBSto.WbfDmInvSummary" Title="SB WMS : Summary Invoice" MasterPageFile="~/SBSto.Master" %>
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
</head>
    <body>
        
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>SUMMARY INVOICE VS PICKLIST YANG SUDAH DIPREPARE</strong></td>                                
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
                                            <tr style="vertical-align:top">
                                                <td style="width:450px">
                                                    <table>
                                                        <tr>
                                                            <td>Company</td>
                                                            <td>:</td>
                                                            <td>
                                                                <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="350px">
                                                                </asp:DropDownList>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td>
                                                    <asp:Button ID="BtnNotaFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:CheckBox ID="ChkVarianOnly" runat="server" Checked="True" Font-Size="10pt" ForeColor="Blue" Text="Tampilkan HANYA Data Error" />
                                                    <br />
                                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                    <asp:Label ID="LblMsgReturn0" runat="server" Font-Size="10pt" ForeColor="#0066FF">1 Total Qty Invoice &lt; Total Qty Picklist</asp:Label>
                                                    <br />
                                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                    <asp:Label ID="LblMsgReturn1" runat="server" Font-Size="10pt" ForeColor="#0066FF">2 Total Qty Picklist di Invoice &lt;&gt; Total Qty Picklist</asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td colspan="2">
                                                    <asp:Panel ID="PanInv" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:450px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                        <asp:GridView ID="GrvInv" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <Columns>
                                                                <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                                    <HeaderStyle Width="50px" />
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                </asp:BoundField>
                                                                <asp:ButtonField CommandName="NotaNo" DataTextField="NotaNo" Text="Button" HeaderText="No. Invoice">
                                                                    <HeaderStyle Width="145px" />
                                                                    <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                                </asp:ButtonField>
                                                                <asp:BoundField DataField="KodeBarang" HeaderText="Kode Barang" HtmlEncode="false">
                                                                    <HeaderStyle Width="95px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vTotalQty_Invoice" HeaderText="A<br />Total<br />Qty Invoice" DataFormatString="{0:n0}" HtmlEncode="false">
                                                                    <HeaderStyle Width="75px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vTotalQty_Invoice_OnPickList" HeaderText="B<br />Total<br />Qty Picklist<br />di Invoice" DataFormatString="{0:n0}" HtmlEncode="false">
                                                                    <HeaderStyle Width="75px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vTotalQty_Picklist" HeaderText="C<br />Total<br />Qty Picklist" DataFormatString="{0:n0}" HtmlEncode="false">
                                                                    <HeaderStyle Width="75px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vSelisihAC" HeaderText="A - C<br />Selisih<br />" DataFormatString="{0:n0}" HtmlEncode="false">
                                                                    <HeaderStyle Width="75px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vSelisihBC" HeaderText="B - C<br />Selisih<br />" DataFormatString="{0:n0}" HtmlEncode="false">
                                                                    <HeaderStyle Width="75px" />
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
                                                    </asp:Panel>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td></td>
                            <td>
                                <asp:HiddenField ID="HdfCompCode" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfCustCode" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfBrgCode" runat="server" />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>        
    </body>
</asp:Content>
