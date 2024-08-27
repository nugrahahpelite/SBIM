<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoStorageCompanyMs.aspx.vb" Inherits="SBSto.WbfSsoStorageCompanyMs" MasterPageFile="~/SBSto.Master" Title="SB WMS : Master Storage - Company" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Master Storage</title>
    <script type="text/javascript">
        function fsShowSimpanProgress() {
            document.getElementById("<%= BtnSimpanCompany.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatalCompany.ClientID%>").style.display = "none";
            document.getElementById("<%= LblMsgSimpan.ClientID%>").innerText = "Sedang Proses...";
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
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER STORAGE - COMPANY</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <asp:DropDownList ID="DstWarehouse" runat="server" style="height: 20px" Width="225px" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </td>
                                        <td style="width:45px"></td>
                                        <td>Building</td>
                                        <td>:</td>
                                        <td>
                                            <asp:DropDownList ID="DstBuilding" runat="server" style="height: 20px" Width="225px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Lantai</td>
                                        <td>:</td>
                                        <td>
                                            <asp:DropDownList ID="DstLantai" runat="server" style="height: 20px" Width="225px">
                                            </asp:DropDownList>
                                        </td>
                                        <td style="width:45px"></td>
                                        <td>Zona</td>
                                        <td>:</td>
                                        <td>
                                            <asp:DropDownList ID="DstZona" runat="server" style="height: 20px" Width="225px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Storage Type</td>
                                        <td>:</td>
                                        <td>
                                            <asp:DropDownList ID="DstStorageType" runat="server" style="height: 20px" Width="225px" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </td>
                                        <td></td>
                                        <td colspan="3">
                                            <asp:Button ID="BtnDisplay" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Display Data" Width="125px" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td colspan="7">
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td colspan="7">

                                            <asp:GridView ID="GrvLsStorage" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True" Width="945px">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="OID" HeaderText="OID">
                                                        <HeaderStyle Width="65px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="BuildingName" HeaderText="Building">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="LantaiDescription" HeaderText="Lantai">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ZonaName" HeaderText="Zona">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:ButtonField CommandName="Select" DataTextField="StorageTypeName" HeaderText="Storage Type">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="StorageSequenceNumber" HeaderText="Sequence&lt;br /&gt;Number" HtmlEncode="false">
                                                        <HeaderStyle Width="75px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="StorageLevel" HeaderText="Level" HtmlEncode="false">
                                                        <HeaderStyle Width="70px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="StorageColumn" HeaderText="Column" HtmlEncode="false">
                                                        <HeaderStyle Width="70px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="StorageNumber" HeaderText="Storage&lt;br /&gt;Number" HtmlEncode="false">
                                                        <HeaderStyle Width="70px" />
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
                            </td>
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <table style="width:100%; font-family:Tahoma;font-size:12px;" >
                                    <tr>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>LIST COMPANY</strong></td>
                                    </tr>
                                    <tr>
                                        <td style="font-size:15px;height:28px"><asp:Button ID="BtnEditCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Edit Company" Width="125px" />
                                            <asp:Button ID="BtnSimpanCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Blue" Height="30px" Text="Simpan" Visible="False" Width="75px" OnClientClick="fsShowSimpanProgress();" />
                                            <asp:Button ID="BtnBatalCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Red" Height="30px" Text="Batal" Visible="False" Width="65px" />
                                            <asp:Label ID="LblMsgSimpan" runat="server" ForeColor="#0066FF"></asp:Label>
                                            <asp:Label ID="LblMsgErrorNE" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:HiddenField ID="HdfStorageOID" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align:top">
                                            <asp:GridView ID="GrvCompany" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" PageSize="20">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:TemplateField>
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="ChkCompany" Text="" runat="server" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="35px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:TemplateField>
                                                    <asp:ButtonField DataTextField="CompanyName" HeaderText="Company" Text="Company" CommandName="Select">
                                                        <HeaderStyle Width="245px" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="CompanyCode" HeaderText="CompanyCode" />
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
    </asp:Content>
