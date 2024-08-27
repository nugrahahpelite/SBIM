<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSalesOrder.aspx.vb" Inherits="SBSto.WbfSsoSalesOrder" Title="SB WMS - Sales Order" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Upload Sales Order</title>

        <script src="../JScript/jquery-1.12.4.js"></script>
        <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

        <script type="text/javascript">
            $(function () {
                $("#<%= TxtSOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtSOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            });
            $(document).ready(function () {
                Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

                function EndRequestHandler(sender, args) {
                    $("#<%= TxtSOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                    $("#<%= TxtSOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
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
                <asp:Panel ID="PanSOData" runat="server" style="height:525px">
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
                                                        <asp:DropDownList ID="DstPOCompany" runat="server" style="height: 20px" Width="350px"></asp:DropDownList>
                                                    </td>
                                                </tr>
                                            </table>                                                                
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgPOCompany" runat="server" ForeColor="#FF0066"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>No. Sales Order</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtSONo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>                                                                                                                            
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkSt_TRB_Not" runat="server" Checked="True" ForeColor="#336600" Text="Belum Summary TRB" />
                                            &nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_TRB_Sebagian" runat="server" Checked="True" ForeColor="#336600" Text="Summary TRB Sebagian" />
                                            &nbsp;&nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_TRB_Full" runat="server" Checked="True" ForeColor="#336600" Text="Full Summary TRB" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>Customer</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtSOCustomer" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        &nbsp;</td>
                                                </tr>
                                            </table>                                                                                                                                                                                                
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgPOError" runat="server" ForeColor="#FF0066"></asp:Label>
                                        </td>                                                            
                                    </tr>
                                    <tr>
                                        <td>Periode</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtSOStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                    <td>s/d</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtSOEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnSOFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgSOFindError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnSOUpload" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="UPLOAD SALES ORDER" Width="195px" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table>
                                    <tr style="vertical-align:top">
                                        <td>
                                            <asp:GridView ID="GrvSOH" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:ButtonField CommandName="SalesOrderNo" DataTextField="SalesOrderNo" Text="Button" HeaderText="No. SO" >
                                                        <HeaderStyle Width="115px" />
                                                        <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="vSalesOrderDate" HeaderText="Tanggal SO">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vSUB" HeaderText="Kode Customer">
                                                        <HeaderStyle Width="90px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="NAMA_CUSTOMER" HeaderText="Nama Customer" HtmlEncode="false">
                                                        <HeaderStyle Width="145px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SalesOrderHOID" HeaderText="Sales Order Header&lt;br /&gt;OID" HtmlEncode="false">
                                                        <HeaderStyle Width="75px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vSOVoid" HeaderText="Void" HtmlEncode="false">
                                                        <HeaderStyle Width="45px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vSOVoid_Info" HeaderText="Void Info" HtmlEncode="false">
                                                        <HeaderStyle Width="215px" />
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
                                            <asp:Panel ID="PanSOD" runat="server" style="height:525px">
                                                <table>
                                                    <tr>
                                                        <td style="width:1000px">
                                                            <asp:Label ID="LblMsgSODNo" runat="server" ForeColor="#0066FF" Font-Size="12px"></asp:Label>
                                                            &nbsp;&nbsp;&nbsp; -&nbsp;&nbsp;
                                                            <asp:Label ID="LblMsgSOHOID" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:GridView ID="GrvSOD" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" AllowPaging="True">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <Columns>
                                                                    <asp:BoundField DataField="GDGOJL" HeaderText="Gudang" HtmlEncode="false">
                                                                        <HeaderStyle Width="80px" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="BRG" HeaderText="Kode Barang" HtmlEncode="false">
                                                                        <HeaderStyle Width="80px" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang">
                                                                        <HeaderStyle Width="245px" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="QTY" HeaderText="Qty" DataFormatString="{0:n0}" >
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:ButtonField CommandName="QTY_TRB" DataTextField="QTY_TRB" Text="Button" HeaderText="Qty TRB" datatextformatstring="{0:n0}" >
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                                    </asp:ButtonField>
                                                                    <asp:BoundField DataField="vSalesOrderDOID" HeaderText="Sales Order Detail<br />OID" HtmlEncode="false">
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="SalesOrderHOID" HeaderText="Sales Order Header<br />OID" HtmlEncode="false">
                                                                        <HeaderStyle Width="75px" />
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
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="PanSOUpload" runat="server" style="height:525px" Visible="false">
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td style="width:10px"></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>UPLOAD XLS SALES ORDER</strong></td>
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
                                        <td>:</td>
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
                                                        <asp:Label ID="LblMsgFupXls" runat="server" ForeColor="Red"></asp:Label>
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

                                                        <asp:Label ID="LblMsgXlsWorksheet" runat="server" ForeColor="Red"></asp:Label>

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
                                            <asp:Label ID="LblMsgXlsProsesError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblMsgError" runat="server" ForeColor="Red"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnXlsUpload" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" OnClientClick="fsShowInProgress('XML')" Text="Upload Xls" Width="145px" />
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
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>HISTORY UPLOAD XLS SALES ORDER</strong></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table>
                                                    <tr>
                                                        <td class="auto-style1" >
                                                            <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="300px">
                                                            </asp:DropDownList>
                                                        </td>
                                                        <td style="width:35px"></td>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        </td>
                                                        <td>
                                                            <asp:Label ID="LblMsgFindError" runat="server" ForeColor="Red"></asp:Label>
                                                        </td>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnData" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Data Sales Order  " Width="145px" Font-Bold="True" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr>
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:ButtonField CommandName="OID" DataTextField="OID" HeaderText="OID">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:ButtonField>
                                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="XlsFileName" HeaderText="Xls File Name">
                                                            <HeaderStyle Width="200px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="UploadStartDatetime" HeaderText="Upload Start">
                                                            <HeaderStyle Width="145px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="UploadEndDatetime" HeaderText="Upload End">
                                                            <HeaderStyle Width="145px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vUploadBy" HeaderText="Upload By">
                                                            <HeaderStyle Width="145px" />
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
        width: 3px;
    }
    </style>
</asp:Content>
