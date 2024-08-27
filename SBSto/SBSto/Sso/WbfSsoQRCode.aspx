<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoQRCode.aspx.vb" Inherits="SBSto.WbfSsoQRCode" MasterPageFile="~/SBSto.Master" Title="SB WMS : Print QR" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Print QR Barang</title>

        <script src="../JScript/jquery-1.12.4.js"></script>
        <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

        <script type="text/javascript">
            function fsShowInProgress(vriProcess) {
                document.getElementById("<%= LblGenProses.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
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
                                        <td style="width:10px"></td>
                                        <td style="font-size:15px;height:28px" colspan="5"><strong>PRINT QR</strong></td>
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
                                        <td>Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstWarehouse" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgWarehouse" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width:125px">Print OID</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtOID" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="85px" BackColor="#DBDBDB" ReadOnly="True"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Kode Barang</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" BackColor="#DBDBDB" ReadOnly="True"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 24px"></td>
                                        <td style="height: 24px">Nama Barang</td>
                                        <td style="height: 24px">:</td>
                                        <td style="height: 24px">
                                            <table>
                                                <tr>
                                                    <td class="auto-style2">
                                                        <asp:TextBox ID="TxtBrgName" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="400px" BackColor="#DBDBDB" ReadOnly="True"></asp:TextBox>
                                                    </td>
                                                    <td style="height: 20px">
                                                        <asp:Label ID="LblMsgBrgName" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Satuan</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        
                                                        <asp:TextBox ID="TxtBrgUnit" runat="server" BackColor="#DBDBDB" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="85px"></asp:TextBox>
                                                        
                                                    </td>                                                    
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Jumlah Print</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPrintCount" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="95px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgPrintCount" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Note</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPrintNote" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="300px" Height="35px" MaxLength="450" TextMode="MultiLine"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgPrintNote" runat="server" ForeColor="Red" Visible="False"></asp:Label>
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
                                            <asp:Label ID="LblGenProses" runat="server" ForeColor="#0066FF"></asp:Label>
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
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnPrint" class="myButtonAct" runat="server" Text="Print" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>&nbsp;</td>
                                        <td></td>
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
                                            <div id="DivPreview" runat="server" >
                                                <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:75%;margin-top:-45px">
                                                    <table>
                                                        <tr style="vertical-align:top">
                                                            <td style="width:98%">
                                                                <iframe runat="server" id="ifrPreview" style="width:100%;height:750px" ></iframe>
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnPreviewClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                            </td>
                                                        </tr>
                                                    </table>                            
                                                </asp:Panel>
                                            </div>
                                        </td>
                                    </tr>
                                </table>                        
                            </td>
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <asp:Panel ID="PanList" class="myPanelGreyNSa" runat="server" Width="100%" style="height:500px" Visible="True">
                                    <table style="width: 90%;margin:auto;font-family: tahoma;font-size:11px">
                                        <tr>
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>HISTORY PRINT QR</strong></td>
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
                                                        <td class="auto-style1"><asp:TextBox ID="TxtListBarang" runat="server" Width="245px" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox></td>
                                                        <td style="width:35px"></td>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        </td>
                                                        <td style="width:10px"></td>                                                        
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr>
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:BoundField DataField="OID" HeaderText="OID" >
                                                            <HeaderStyle Width="50px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company" >
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" >
                                                            <HeaderStyle Width="80px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" >
                                                            <HeaderStyle Width="80px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                            <HeaderStyle Width="245px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" >
                                                            <HeaderStyle Width="60px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="PrintCount" HeaderText="Jumlah">
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="PrintNote" HeaderText="Note">
                                                            <HeaderStyle Width="145px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="PrintDatetime" HeaderText="Print Datetime">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vPrintUserName" HeaderText="Print User Name">
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
        .auto-style1 {
            height: 18px;
        }
        .auto-style2 {
            height: 20px;
            width: 408px;
        }
    </style>
</asp:Content>
