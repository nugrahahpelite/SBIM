<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoProductMs.aspx.vb" Inherits="SBSto.WbfSsoProductMs" MasterPageFile="~/SBSto.Master" Title="SB WMS : Master Barang" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Master Barang</title>

        <script src="../JScript/jquery-1.12.4.js"></script>
        <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

        <script type="text/javascript">
            function fsShowInProgress(vriProcess) {
                document.getElementById("<%= BtnXlsUpload.ClientID%>").style.display = "none";
                document.getElementById("<%= LblXlsProses.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
                document.getElementById("<%= LblMsgXlsProsesError.ClientID%>").innerText = "";
            }
            function setLblFupUpload() {
                document.getElementById("<%= LblFupXls.ClientID%>").textContent = GetFullPathFileName(document.getElementById("<%= FupXls.ClientID%>").value);
            }
            function GetFullPathFileName(vriString) {
                var FName;
                var FNameFound = '';
                var step;
                FName = '';
                for (step = vriString.length; step > -1; step--) {
                    if (vriString.substr(step, 1) == "\\") {
                        FName = vriString.substr(step + 1, vriString.length);
                        FNameFound = 'Y';
                        return FName;
                    }
                }
                if (FNameFound == '') {
                    FName = vriString;
                }
                return FName;
            }
        </script>
    </head>
    <body>
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnXlsUpload" />
            </Triggers>
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td style="width:10px"></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER BARANG</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Company</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="300px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgCompany" runat="server" ForeColor="Red"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>                                        
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>File xls</td>
                                        <td>&nbsp;</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>

                                                        <asp:FileUpload ID="FupXls" runat="server" accept=".xls" onchange="setLblFupUpload()" style="color:transparent" Width="425px" />

                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 24px"></td>
                                        <td style="height: 24px">&nbsp;</td>
                                        <td style="height: 24px">&nbsp;</td>
                                        <td style="height: 24px">
                                            <table>
                                                <tr>
                                                    <td style="height: 20px">
                                                        <asp:Label ID="LblMsgFupXls" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                        <asp:Label ID="LblFupXls" runat="server" ForeColor="#3333FF"></asp:Label>
                                                    </td>
                                                    <td style="height: 20px">
                                                        &nbsp;</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Nama Worksheet</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtXlsWorksheet" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px">Sheet 1</asp:TextBox>
                                                    </td>
                                                    <td>

                                                        <asp:Label ID="LblXlsWorksheet" runat="server" ForeColor="Red" Visible="False"></asp:Label>

                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <asp:Label ID="LblXlsProses" runat="server" ForeColor="#0066FF"></asp:Label>
                                            <asp:Label ID="LblMsgXlsProsesError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblMsgError" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>
                                            &nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnXlsUpload" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" OnClientClick="fsShowInProgress('Upload')" Text="Upload Xls" Width="145px" />
                                                    </td>
                                                    <td style="width:25px"></td>
                                                    <td>
                                                        <asp:Button ID="BtnGenQRCode" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Generate QR Code" Width="145px" Enabled="False" Visible="False" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>                        
                            </td>
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <asp:Panel ID="PanList" class="myPanelGreyNSa" runat="server" Width="100%" style="height:500px" Visible="True">
                                    <table style="width: 90%;margin:auto;font-family: tahoma;font-size:11px">
                                        <tr>
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>LIST BARANG</strong></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table>
                                                    <tr>
                                                        <td>Company</td>
                                                        <td>:</td>
                                                        <td>
                                                            <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="300px"></asp:DropDownList>
                                                        </td>                                        
                                                    </tr>
                                                    <tr>
                                                        <td class="auto-style1" >Barang</td>
                                                        <td class="auto-style1">:<br />
                                                        </td>
                                                        <td class="auto-style1"><asp:TextBox ID="TxtListKriteria" runat="server" Width="245px" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox></td>
                                                        <td style="width:35px">&nbsp;</td>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        </td>
                                                        <td style="width:10px"></td>                                                        
                                                    </tr>
                                                    <tr>
                                                        <td></td>
                                                        <td></td>
                                                        <td colspan="3">
                                                            <asp:CheckBox ID="ChkListSN_Yes" runat="server" Checked="True" Text="Is SN" />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <asp:CheckBox ID="ChkListSN_No" runat="server" Checked="True" Text="Not SN" />
                                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                            <asp:CheckBox ID="ChkListIsSNNew" runat="server" Text="Is SN Baru (Belum Diupdate)" ForeColor="#FF0066" AutoPostBack="True" />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td></td>
                                                        <td></td>
                                                        <td colspan="3">
                                                            <asp:CheckBox ID="ChkListActive_Yes" runat="server" Checked="True" Text="Is Active" />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <asp:CheckBox ID="ChkListActive_No" runat="server" Checked="True" Text="Not Active" />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td></td>
                                                        <td colspan="4">
                                                            <asp:Label ID="LblMsgListError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr style="vertical-align:top">
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang">
                                                            <HeaderStyle Width="245px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:ButtonField CommandName="vConfirmNotSN" DataTextField="vConfirmNotSN" HeaderText="Not SN">
                                                            <HeaderStyle Width="85px" CssClass="myDisplayNone" />
                                                            <ItemStyle HorizontalAlign="Center" ForeColor="#0066FF" CssClass="myDisplayNone" />
                                                        </asp:ButtonField>
                                                        <asp:ButtonField CommandName="vIsSN" DataTextField="vIsSN" HeaderText="Is SN">
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" ForeColor="#0066FF" />
                                                        </asp:ButtonField>
                                                        <asp:ButtonField CommandName="vIsActive" DataTextField="vIsActive" HeaderText="Is Active">
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" ForeColor="#0066FF" />
                                                        </asp:ButtonField>
                                                        <asp:ButtonField CommandName="vHS" DataTextField="vHS" HeaderText="History">
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" ForeColor="#0066FF" />
                                                        </asp:ButtonField>
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
                                            <td>
                                                <asp:Label ID="LblHS" runat="server" Font-Size="12px" ForeColor="#3333FF"></asp:Label>
                                                <br />
                                                <asp:GridView ID="GrvHS" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="25" ShowHeaderWhenEmpty="True">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                                            <HeaderStyle Width="45px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vIsSN" HeaderText="Is SN">
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vIsActive" HeaderText="Is Active">
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vEditBy" HeaderText="Edit isSN By">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vEditAt" HeaderText="Edit isSN at">
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