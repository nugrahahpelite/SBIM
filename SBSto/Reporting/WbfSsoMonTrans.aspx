<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoMonTrans.aspx.vb" Inherits="SBSto.WbfSsoMonTrans" MasterPageFile="~/SBSto.Master" Title="SB WMS : Monitor Transaksi" %>
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
                                            <asp:Label ID="LblMsgError0" runat="server" Font-Bold="True" ForeColor="#0066FF">PENERIMAAN PEMBELIAN</asp:Label>
                                        </td>
                                        <td></td>
                                        <td>STATUS</td>
                                        <td>
                                            <asp:CheckBox ID="ChkGR_OnReceive" runat="server" Checked="True" ForeColor="#336600" Text="On Receive" />
                                            &nbsp;<asp:CheckBox ID="ChkGR_ReceiveDone" runat="server" Checked="True" ForeColor="#336600" Text="Receive Done" />
                                            &nbsp;<asp:CheckBox ID="ChkGR_ReceiveApp" runat="server" Checked="True" ForeColor="#336600" Text="Receive Approved" />
                                            &nbsp;
                                            <asp:CheckBox ID="ChkGR_PutawayProcess" runat="server" ForeColor="#336600" Text="Putaway Process" Checked="True" />
                                            &nbsp;
                                            <asp:CheckBox ID="ChkGR_AllPutawayComplete" runat="server" ForeColor="#336600" Text="All Putaway Complete" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:Panel ID="PanGR" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:250px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvGR" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="RcvTypeName" HeaderText="Tipe">
                                                <HeaderStyle Width="65px" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="Select" DataTextField="RcvPONo" HeaderText="Nomor Penerimaan" Text="Button">
                                                <HeaderStyle Width="125px" />
                                                <ItemStyle ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal Penerimaan">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="RcvPORefNo" HeaderText="PL/DO">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="RcvPOTypeName" HeaderText="Import/&lt;br /&gt;Local" HtmlEncode="false">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="RcvPOCompanyCode" HeaderText="Company">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vCreation" HeaderText="Creation">
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vReceivedDone" HeaderText="Receive Done">
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vReceivedApp" HeaderText="Receive Approved">
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
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
                    <table style="width:95%">
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="#0066FF">PUTAWAY</asp:Label>
                                        </td>
                                        <td></td>
                                        <td>STATUS</td>
                                        <td>
                                            <asp:CheckBox ID="ChkPtw_Process" runat="server" Checked="True" ForeColor="#336600" Text="Putaway Process" />
                                            &nbsp;<asp:CheckBox ID="ChkPtw_Done" runat="server" Checked="True" ForeColor="#336600" Text="Putaway Done" />
                                            &nbsp; &nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:Panel ID="Panel1" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:250px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvPtw" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="TransCode" HeaderText="TransCode" >
                                                <HeaderStyle Width="65px" CssClass="myDisplayNone" />
                                                <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransName" HeaderText="Transaksi" >
                                                <HeaderStyle Width="115px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPtwOID" HeaderText="ID Transaksi" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPtwCompanyCode" HeaderText="Company" HtmlEncode="false">
                                                <HeaderStyle Width="55px" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="vPtwNo" DataTextField="vPtwNo" HeaderText="No.Putaway" Text="Button">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="RcvPONo" HeaderText="No.Penerimaan" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPtwDate" HeaderText="Tanggal<br />Putaway" HtmlEncode="false">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Center"/>
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vOnDelivery" HeaderText="On Delivery" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vOnPutaway" HeaderText="On Putaway" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPutawayDone" HeaderText="Putaway Done" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
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
                    <table style="width:95%">
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:Label ID="Label3" runat="server" Font-Bold="True" ForeColor="#0066FF">PICKLIST</asp:Label>
                                        </td>
                                        <td></td>
                                        <td>STATUS</td>
                                        <td>
                                            <asp:CheckBox ID="ChkPck_Baru" runat="server" ForeColor="#336600" Text="Baru" />
                                            &nbsp;&nbsp;<asp:CheckBox ID="ChkPck_Prepared" runat="server" Checked="True" ForeColor="#336600" Text="Prepared" />
                                            &nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkPck_OnPicking" runat="server" Checked="True" ForeColor="#336600" Text="On Picking" />
                                            &nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkPck_PickingDone" runat="server" ForeColor="#336600" Text="Picking Done" />
                                            &nbsp;&nbsp; &nbsp; &nbsp;</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:Panel ID="PanPck" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:250px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvPck" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="Select" DataTextField="PCLNo" Text="Button" HeaderText="Nomor Pick List">
                                                <HeaderStyle Width="125px" />
                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="vPCLDate" HeaderText="Tanggal<br />Pick List" HtmlEncode="false" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPCLScheduleDate" HeaderText="Schedule<br />Pick List" HtmlEncode="false" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLCompanyCode" HeaderText="Company" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SchDTypeName" HeaderText="Jenis" >
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang Asal"> 
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLRefHOID" HeaderText="PCLRefHOID"> 
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLRefHNo" HeaderText="Ref No."> 
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPCLRefHInfo" HeaderText="Ref Info" HtmlEncode="false"> 
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCKNo" HeaderText="No. Picking"> 
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vDspPtwNo" HeaderText="No. Dispatch/Putaway"> 
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLNote" HeaderText="Note" >
                                                <HeaderStyle Width="120px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStatusPickList" HeaderText="Status Picklist" >
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStatusPicking" HeaderText="Status Picking" >
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
                    <table style="width:95%">
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:Label ID="Label2" runat="server" Font-Bold="True" ForeColor="#0066FF">PINDAH LOKASI</asp:Label>
                                        </td>
                                        <td></td>
                                        <td>STATUS</td>
                                        <td>
                                            <asp:CheckBox ID="ChkMove_Process" runat="server" Checked="True" ForeColor="#336600" Text="Movement Process" />
                                            &nbsp;<asp:CheckBox ID="ChkMove_Done" runat="server" Checked="True" ForeColor="#336600" Text="Movement Done" />
                                            &nbsp; &nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:Panel ID="PanMove" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:250px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvMove" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="TransCode" HeaderText="TransCode" >
                                                <HeaderStyle Width="65px" CssClass="myDisplayNone" />
                                                <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransName" HeaderText="Transaksi" >
                                                <HeaderStyle Width="115px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vMoveOID" HeaderText="ID Transaksi" >
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vMoveCompanyCode" HeaderText="Company" HtmlEncode="false">
                                                <HeaderStyle Width="55px" />
                                            </asp:BoundField>
                                            <asp:ButtonField CommandName="vMoveNo" DataTextField="vMoveNo" HeaderText="No.Pindah Lokasi" Text="Button">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <asp:BoundField DataField="vMoveDate" HeaderText="Tanggal<br />Pindah Lokasi" HtmlEncode="false">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Center"/>
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vOnDelivery" HeaderText="On Delivery" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vOnMovement" HeaderText="On Movement" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vMovementDone" HeaderText="Movement Done" >
                                                <HeaderStyle Width="145px" />
                                                <ItemStyle HorizontalAlign="Left" />
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