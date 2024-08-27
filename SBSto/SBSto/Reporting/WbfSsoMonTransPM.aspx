<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoMonTransPM.aspx.vb" Inherits="SBSto.WbfSsoMonTransPM" MasterPageFile="~/SBSto.Master" Title="SB WMS : Monitor Transaksi" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Monitor Transaksi</title>
    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
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
        </style>
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
                            <td style="font-size:15px;height:28px" colspan="3"><strong>MONITOR TRANSAKSI</strong></td>
                        </tr>
                        <tr>
                            <td>
                                <table>
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
                                                                    <asp:DropDownList ID="DstListWarehouse" runat="server" AutoPostBack="True" style="height: 20px" Width="225px"></asp:DropDownList>
                                                                </td>
                                                                <td>
                                                                    <asp:Label ID="LblMsgListWarehouse" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                </td>
                                                            </tr>
                                                        </table>                                                        
                                                    </td>
                                                    <td>Tanggal</td>
                                                    <td>:</td>
                                                    <td>                                                        
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:TextBox ID="TxtListStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                    s/d
                                                                    <asp:TextBox ID="TxtListEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                </td>
                                                                <td>
                                                                    <asp:Label ID="LblMsgListStart" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                </td>
                                                                <td>
                                                                    &nbsp;</td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnRefreshAll" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="REFRESH ALL" Width="112px" />
                                                    </td>
                                                    <td>

                                                        <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>

                                                    </td>
                                                </tr>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>                
                        </tr>
                    </table>
                    <table style="width:95%">
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError0" runat="server" Font-Bold="True" ForeColor="#0066FF">TRANSAKSI</asp:Label>
                                        </td>
                                        <td></td>
                                        <td>STATUS</td>
                                        <td>
                                            <asp:CheckBox ID="ChkPM_InProgress" runat="server" Checked="True" ForeColor="#336600" Text="IN Progress" />
                                            &nbsp;<asp:CheckBox ID="ChkPM_Done" runat="server" Checked="True" ForeColor="#336600" Text="Done" />
                                            &nbsp;<asp:CheckBox ID="ChkPM_Batal" runat="server" Checked="True" ForeColor="#FF0066" Text="Batal" />
                                            &nbsp; &nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:Panel ID="PanPM" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:700px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvPM" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" HeaderStyle-CssClass="StickyHeader" ShowHeaderWhenEmpty="True">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="vTransOID" HeaderText="ID Transaksi">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vRefOID" HeaderText="Ref ID">
                                                <HeaderStyle Width="65px" CssClass="myDisplayNone" />
                                                <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransCode" HeaderText="TransCode">
                                                <HeaderStyle CssClass="myDisplayNone" Width="65px" />
                                                <ItemStyle CssClass="myDisplayNone" HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransName" HeaderText="Transaksi">
                                                <HeaderStyle Width="115px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTransCompanyCode" HeaderText="Company" HtmlEncode="false">
                                                <HeaderStyle Width="55px" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="vTransNo" DataTextField="vTransNo" HeaderText="No.Transaksi" Text="Button">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="vTransDate" HeaderText="Tanggal&lt;br /&gt;Transaksi" HtmlEncode="false">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="vRefNo" DataTextField="vRefNo" HeaderText="Ref No" Text="Button">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="vTransType" HeaderText="Type" HtmlEncode="false">
                                                <HeaderStyle Width="55px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTransNote" HeaderText="Note" HtmlEncode="false">
                                                <HeaderStyle Width="215px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vCreation" HeaderText="Creation">
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                        </Columns>
                                        <EditRowStyle BackColor="#999999" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" Font-Overline="False" ForeColor="White" Height="35px" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                        <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                        <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                    </asp:GridView>
                                </asp:Panel>
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