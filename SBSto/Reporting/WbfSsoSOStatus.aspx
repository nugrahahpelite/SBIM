<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSOStatus.aspx.vb" Inherits="SBSto.WbfSsoSOStatus" MasterPageFile="~/SBSto.Master" Title="SB WMS : Summary Status Stock Opname" %>
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
                <asp:PostBackTrigger ControlID="BtnPdf" />
                <asp:PostBackTrigger ControlID="BtnXLS" />
            </Triggers>     
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td style="font-size:15px;height:28px" colspan="3"><strong>SUMMARY STATUS STOCK OPNAME</strong></td>
                        </tr>
                        <tr>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:CheckBox ID="ChkSt_NotClosed" runat="server" ForeColor="Blue" Text="NOT CLOSED" Checked="True" />
                                            &nbsp;&nbsp;&nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_Closed" runat="server" ForeColor="Red" Text="CLOSED" />
                                        </td>
                                        <td style="width:45px"></td>
                                        <td style="width:345px">
                                            <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                            &nbsp;&nbsp;
                                            <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnPdf" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  PDF   " Width="112px" />
                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                            <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />
                                            &nbsp;&nbsp;
                                            <asp:Label ID="LblProgressXLS" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>                
                        </tr>
                        <tr>
                            <td>
                                <div id="DivPreview" runat="server" >
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="z-index:300;display:block;width:75%;margin-top:-45px">
                                        <table>
                                            <tr style="vertical-align:top">
                                                <td style="width:98%">
                                                    <iframe runat="server" id="ifrPreview" style="width:100%;height:750px" ></iframe>
                                                </td>
                                                <td>
                                                    <asp:Button ID="BtnPreviewClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClick="BtnPreviewClose_Click" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                        </table>                            
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:700px;" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" HeaderStyle-CssClass="StickyHeader" >
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="SOHOID" HeaderText="SO OID">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SOCompanyCode" HeaderText="Company">
                                                <HeaderStyle Width="70px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SONo" HeaderText="SONo">
                                                <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SONote" HeaderText="SO Note">
                                                <HeaderStyle Width="245px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTotalItem_System" HeaderText="Total Item<br />In System" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTotalQty_System" HeaderText="Total Qty<br />In System" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTotalItem_Scanned" HeaderText="Total Item<br />Scanned" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPercentageItem_Scanned" HeaderText="Percentage Item<br />Scanned" DataFormatString="{0:0.00}%" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTotalQty_Scanned" HeaderText="Total Qty<br />Scanned" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPercentageQty_Scanned" HeaderText="Percentage Qty<br />Scanned" DataFormatString="{0:0.00}%" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vTotalItem_Selisih" HeaderText="Total Item<br />Selisih" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="90px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPercentageItem_Selisih" HeaderText="Percentage Item<br />Selisih" DataFormatString="{0:0.00}%" HtmlEncode="false">
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
                                &nbsp;</td>
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
