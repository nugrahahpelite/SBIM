<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoTransStatus.aspx.vb" Inherits="SBSto.WbfSsoTransStatus" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../Content/CssBnsrp.css" rel="stylesheet"  type="text/css" />

    <style type="text/css">

.ui-priority-primary,
.ui-widget-content .ui-priority-primary,
.ui-widget-header .ui-priority-primary {
	font-weight: bold;
}
*,::after,::before{text-shadow:none!important;box-shadow:none!important}*,::after,::before{box-sizing:border-box}</style>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table class="myPanelGreyLight" >
                <tr>
                    <td colspan="3">
                        <asp:Label ID="LblPageTitle" runat="server" ForeColor="#333333" CssClass="myPanelGreyLight" Font-Bold="True" Font-Size="15px">HISTORY STATUS</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="width:115px">Nomor Transaksi</td>
                    <td style="width:20px;text-align:center">:</td>
                    <td style="width:415px">
                        <asp:TextBox ID="TxtTransNo" runat="server" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="12px" Width="185px"></asp:TextBox>                        
                    </td>   
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Visible="False"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%" AllowPaging="True">
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <%-- 0 --%>
                                <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                    <HeaderStyle Width="185px" />
                                </asp:BoundField>
                                <%-- 1 --%>
                                <asp:BoundField DataField="vTransStatusBy" HeaderText="Status by">
                                    <HeaderStyle Width="300px" />
                                </asp:BoundField>
                                <%-- 2 --%>
                                <asp:BoundField DataField="vTransStatusInfo" HeaderText="Info" HtmlEncode="false">
                                    <HeaderStyle Width="450px" />
                                </asp:BoundField>
                                <%-- 3 --%>
                                <asp:BoundField DataField="vTransStatusDatetime" HeaderText="Status at">
                                    <HeaderStyle HorizontalAlign="Center" Width="145px" />
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
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
