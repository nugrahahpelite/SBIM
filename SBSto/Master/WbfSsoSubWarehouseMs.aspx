<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSubWarehouseMs.aspx.vb" Inherits="SBSto.WbfSsoSubWarehouseMs" MasterPageFile="~/SBSto.Master" Title="SB WMS : Master Sub Warehouse" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Master Sub Warehouse</title>
    <script type="text/javascript">
        function fsShowSimpanProgress() {
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
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
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER SUB WAREHOUSE</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>ID Sub Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td><asp:TextBox ID="TxtOID" runat="server" Width="87px" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True"></asp:TextBox></td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkActive" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Active" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Company</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="height: 20px">
                                                        <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td style="height: 20px" class="auto-style5">
                                            <asp:Label ID="LblMsgCompany" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="height: 20px">
                                                        <asp:DropDownList ID="DstWhs" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td style="height: 20px" class="auto-style5">
                                            <asp:Label ID="LblMsgWhs" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Kode Sub Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="height: 20px"><asp:TextBox ID="TxtSubWhsCode" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="12px" MaxLength="10" CssClass="setuppercase"></asp:TextBox></td>
                                                    <td style="height: 20px">
                                            <asp:Label ID="LblMsgSubWhsCode" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Nama Sub Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="height: 20px"><asp:TextBox ID="TxtWhsSubName" runat="server" Width="214px" Font-Names="Tahoma" Font-Size="12px" MaxLength="50" CssClass="setuppercase"></asp:TextBox></td>
                                                    <td style="height: 20px">
                                            <asp:Label ID="LblMsgSubWhsName" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Description</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td style="height: 20px"><asp:TextBox ID="TxtWhsSubDescr" runat="server" Width="300px" Font-Names="Tahoma" Font-Size="12px" MaxLength="200" CssClass="setuppercase" Height="45px" TextMode="MultiLine"></asp:TextBox></td>
                                                    <td style="height: 20px">
                                            <asp:Label ID="LblMsgSubWhsDescr" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>
                                        <td>
                                        <td>
                                            &nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblMsgErrorNE" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowSimpanProgress();" />
                                                        <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                                        <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                                        <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                                        <asp:Label ID="LblMsgSimpan" runat="server" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:HiddenField ID="HdfActionStatus" runat="server" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>
                                        <td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Visible="False"></asp:Label>
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
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>LIST SUB WAREHOUSE</strong></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table>
                                                    <tr>
                                                        <td>Nama Sub Warehouse</td>
                                                        <td>:<br />
                                                        </td>
                                                        <td><asp:TextBox ID="TxtKriteria" runat="server" Width="245px" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox></td>
                                                        <td style="width:35px"></td>
                                                        <td style="width:35px">Company</td>
                                                        <td style="width:10px">
                                                            <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="225px">
                                                            </asp:DropDownList>
                                                        </td>
                                                        <td style="width:35px">
                                                            &nbsp;</td>
                                                        <td style="width:35px"></td>
                                                        <td><asp:Button ID="BtnFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" /></td>
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr>
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" Width="945px" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:BoundField DataField="SubWhsCode" HeaderText="Kode Sub Warehouse">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:ButtonField CommandName="Select" DataTextField="SubWhsName" HeaderText="Nama Sub Warehouse">
                                                            <HeaderStyle Width="185px" />
                                                        </asp:ButtonField>
                                                        <asp:BoundField DataField="SubWhsDescription" HeaderText="Deskripsi">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                            <HeaderStyle Width="60px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                                            <HeaderStyle Width="65px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="Status" HeaderText="Status">
                                                            <HeaderStyle Width="85px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CreationDatetime" HeaderText="Creation Date">
                                                            <HeaderStyle Width="145px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CreationUserName" HeaderText="Creation By">
                                                            <HeaderStyle Width="125px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="ModificationDatetime" HeaderText="Modification Date">
                                                            <HeaderStyle Width="145px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="ModificationUserName" HeaderText="Modification By">
                                                            <HeaderStyle Width="125px" />
                                                        </asp:BoundField>
                                                        <asp:CommandField ShowSelectButton="True" SelectText="Pilih" Visible="False" >
                                                            <HeaderStyle Width="54px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:CommandField>
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
        .auto-style5 {
            margin-left: 40px;
        }
    </style>
</asp:Content>
