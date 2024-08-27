<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoTRBDisplay.aspx.vb" Inherits="SBSto.WbfSsoTRBDisplay" Title="SB WMS - TRB" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Upload TRB</title>

        <script src="../JScript/jquery-1.12.4.js"></script>
        <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

        <script type="text/javascript">
            $(function () {
                $("#<%= TxtTRBStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtTRBEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            });
            $(document).ready(function () {
                Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

                function EndRequestHandler(sender, args) {
                    $("#<%= TxtTRBStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                    $("#<%= TxtTRBEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
            });            
        </script>
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
        .auto-style1 {
            width: 300px;
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
                <asp:Panel ID="PanTRBData" runat="server" style="height:525px">
                    <table style="width:100%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td colspan="2">
                                <table>
                                    <tr>
                                        <td>Company</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstTRBCompany" runat="server" style="height: 20px" Width="350px"></asp:DropDownList>
                                                    </td>
                                                </tr>
                                            </table>                                                                
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgTRBCompany" runat="server" ForeColor="#FF0066"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>No. TRB</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtTRBNo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>                                                                                                                            
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgTRBError" runat="server" ForeColor="#FF0066"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="DstTRBWhAsal" runat="server" style="height: 20px" Visible="False" Width="350px">
                                            </asp:DropDownList>
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="DstTRBWhTujuan" runat="server" style="height: 20px" Visible="False" Width="350px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>Periode</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtTRBStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                    <td>s/d</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtTRBEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnTRBFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                            <asp:Label ID="LblMsgTRBFindError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                        <td>
                                            &nbsp;</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table>
                                    <tr style="vertical-align:top">
                                        <td>
                                            <asp:GridView ID="GrvTRBH" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="TRBHOID" HeaderText="TRB OID" HtmlEncode="false">
                                                        <HeaderStyle Width="75px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:ButtonField CommandName="NoBukti" DataTextField="NoBukti" Text="Button" HeaderText="No. TRB" >
                                                        <HeaderStyle Width="115px" />
                                                        <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="vTanggal" HeaderText="Tanggal TRB">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="GudangAsal" HeaderText="Gudang Asal" HtmlEncode="false">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="GudangTujuan" HeaderText="Gudang Tujuan" HtmlEncode="false">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="GudangAsalOID" HeaderText="GudangAsalOID" HtmlEncode="false">
                                                        <HeaderStyle CssClass="myDisplayNone" />
                                                        <ItemStyle CssClass="myDisplayNone" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="GudangTujuanOID" HeaderText="GudangTujuanOID" HtmlEncode="false">
                                                        <HeaderStyle CssClass="myDisplayNone" />
                                                        <ItemStyle CssClass="myDisplayNone" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vIsPickListClosed" HeaderText="Sudah<br />Picklist" HtmlEncode="false">
                                                        <HeaderStyle Width="55px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vTRBCancel" HeaderText="Cancel" HtmlEncode="false">
                                                        <HeaderStyle Width="55px" />
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
                                        </td>
                                        <td>
                                            <asp:Panel ID="PanTRBD" runat="server" style="height:525px">
                                                <table>
                                                    <tr>
                                                        <td class="auto-style2">
                                                            <asp:Label ID="LblMsgTRBDNo" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                            &nbsp; -&nbsp;
                                                            <asp:Label ID="LblMsgTRBHOID" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:Panel ID="PanTRBD_D" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:450px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                                <asp:GridView ID="GrvTRBD" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:BoundField DataField="vSeqNo" HeaderText="No" HtmlEncode="false">
                                                                            <HeaderStyle Width="80px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="KodeBrg" HeaderText="Kode Barang" HtmlEncode="false">
                                                                            <HeaderStyle Width="80px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="NamaBrg" HeaderText="Nama Barang">
                                                                            <HeaderStyle Width="245px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Qty" DataFormatString="{0:n0}" HeaderText="Qty">
                                                                            <HeaderStyle Width="75px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="QtyOnPickList" DataFormatString="{0:n0}" HeaderText="Qty<br />Picklist" HtmlEncode="false">
                                                                            <HeaderStyle Width="75px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Satuan" HeaderText="Satuan">
                                                                            <HeaderStyle Width="55px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Keterangan" HeaderText="Keterangan">
                                                                            <HeaderStyle Width="245px" />
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
                                        </td>
                                    </tr>
                                </table>
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
            height: 18px;
        }
    </style>
</asp:Content>
