<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoPO.aspx.vb" Inherits="SBSto.WbfSsoPO" Title="SB WMS - PO" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Upload PO</title>

        <script src="../JScript/jquery-1.12.4.js"></script>
        <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

        <script type="text/javascript">
            $(function () {
                $("#<%= TxtPOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtPOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtPOHEta.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            });
            $(document).ready(function () {
                Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

                function EndRequestHandler(sender, args) {
                    $("#<%= TxtPOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                    $("#<%= TxtPOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                    $("#<%= TxtPOHEta.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
            function fsShowInProgress(vriProcess) {
                document.getElementById("<%= BtnXlsUpload.ClientID%>").style.display = "none";
                document.getElementById("<%= LblXlsProses.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
                document.getElementById("<%= LblMsgXlsProsesError.ClientID%>").innerText = "";
            }
            function fsShowInProgressSAP(vriProcess) {
                document.getElementById("<%= BtnPOSap.ClientID%>").style.display = "none";
                document.getElementById("<%= LblPOSapProses.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
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
            function fsDisableYesPOHEta() {
                document.getElementById("<%= BtnPOHEtaYes.ClientID%>").style.display = "none";
                document.getElementById("<%= BtnPOHEtaNo.ClientID%>").style.display = "none";
                document.getElementById("<%= LblPOHEtaProgress.ClientID%>").innerText = "In Progress.........";
            }
            function fsDisableYesPOHClo() {
                document.getElementById("<%= BtnPOHCloYes.ClientID%>").style.display = "none";
                document.getElementById("<%= BtnPOHCloNo.ClientID%>").style.display = "none";
                document.getElementById("<%= LblPOHCloProgress.ClientID%>").innerText = "In Progress.........";
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
                <asp:Panel ID="PanPOData" runat="server" style="height:2125px">
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
                                        <td>No. PO</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPONo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>                                                                                                                            
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkSt_PL_Not" runat="server" Checked="True" ForeColor="#336600" Text="Belum PL" />
                                            &nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_PL_Sebagian" runat="server" Checked="True" ForeColor="#336600" Text="PL Sebagian" />
                                            &nbsp;&nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_PL_Full" runat="server" Checked="True" ForeColor="#336600" Text="Full PL" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>Supplier</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPOSupplier" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        &nbsp;</td>
                                                </tr>
                                            </table>                                                                                                                                                                                                
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkSt_GR_Not" runat="server" Checked="True" ForeColor="#0066FF" Text="Belum Penerimaan" Visible="False" />
                                            &nbsp;
                                            <asp:CheckBox ID="ChkSt_GR_Sebagian" runat="server" Checked="True" ForeColor="#0066FF" Text="Penerimaan Sebagian" Visible="False" />
                                            &nbsp;&nbsp;
                                            <asp:CheckBox ID="ChkSt_GR_Full" runat="server" Checked="True" ForeColor="#0066FF" Text="Full Penerimaan" Visible="False" />
                                        </td>
                                        <td>
                                            <asp:HiddenField ID="HdfPOHOID" runat="server" />
                                        </td>
                                        <td>
                                            <asp:HiddenField ID="HdfPOHStatus" runat="server" />
                                        </td>
                                        <td>
                                            <asp:HiddenField ID="HdfPOHRowIdx" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>Periode</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPOStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                    <td>s/d</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtPOEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkPO_NotClosed" runat="server" Checked="True" ForeColor="#990033" Text="PO NOT Closed" />
                                            &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ChkPO_Closed" runat="server" ForeColor="#990033" Text="PO Closed" />
                                        </td>
                                        <td colspan="3">
                                            <asp:Label ID="LblMsgPOError" runat="server" ForeColor="#FF0066"></asp:Label>
                                            <asp:Label ID="LblMsgPOFindError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>Barang</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtPOBrg" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        &nbsp;</td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnPOFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                        </td>
                                        <td>&nbsp;</td>
                                        <td>
                                            <asp:Button ID="BtnPOUpload" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="UPLOAD PO" Width="125px" />
                                        </td>
                                        <td>
                                            <asp:Button ID="BtnPOGetSAP" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="GET PO SAP" Width="125px" Visible="False" />
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
                                            <asp:GridView ID="GrvPOH" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                        <HeaderStyle Width="45px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:ButtonField CommandName="PO_NO" DataTextField="PO_NO" Text="Button" HeaderText="No. PO" >
                                                        <HeaderStyle Width="115px" />
                                                        <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="vPO_DATE" HeaderText="Tanggal PO">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vETA_DATE" HeaderText="ETA">
                                                        <HeaderStyle Width="70px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vSupplier" HeaderText="Supplier">
                                                        <HeaderStyle Width="145px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vPLExist" HeaderText="PL">
                                                        <HeaderStyle Width="75px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vGRExist" HeaderText="GR">
                                                        <HeaderStyle Width="75px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" HtmlEncode="false">
                                                        <HeaderStyle Width="75px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="TransStatus" HeaderText="TransStatus" >
                                                        <HeaderStyle Width="75px" CssClass="myDisplayNone" />
                                                        <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="POHOID" HeaderText="PO Header&lt;br /&gt;OID" HtmlEncode="false">
                                                        <HeaderStyle Width="75px" CssClass="myDisplayNone" />
                                                        <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
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
                                            <asp:Panel ID="PanPOD" runat="server" style="height:525px" BorderStyle="Solid" Visible="False">
                                                <table>
                                                    <tr>
                                                        <td class="auto-style2">
                                                            <table>
                                                                <tr style="vertical-align:top">
                                                                    <td>
                                                                        <asp:Label ID="LblMsgPOHOID" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                        &nbsp; -&nbsp;
                                                                        <asp:Label ID="LblMsgPOHNo" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                        &nbsp; -&nbsp;
                                                                        <asp:Label ID="LblMsgPOHSupplier" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                        <asp:Label ID="LblMsgPOHStatus" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                    </td>
                                                                    <td rowspan="3">
                                                                        <asp:Button ID="BtnPOHEta" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Edit ETA" Width="125px" />
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                        <asp:Button ID="BtnPOHClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Close PO" Width="125px" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td>
                                                                        <asp:Label ID="LblMsgPOHDate" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                        <asp:Label ID="LblMsgPOHETADate" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td>
                                                                        <asp:Panel ID="PanPOHClose" runat="server" Visible="False" >
                                                                            <table>
                                                                                <tr>
                                                                                    <td colspan="2"><asp:Label ID="LblMsgPOHClose" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td style="width:65px">CLOSE NOTE</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="TxtPOHCloseNote" runat="server" BorderStyle="None" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="450" ReadOnly="True" TextMode="MultiLine" Width="300px" Height="35px"></asp:TextBox>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </asp:Panel>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:RadioButtonList ID="RdbPOD" runat="server" RepeatDirection="Horizontal" AutoPostBack="True">
                                                                <asp:ListItem Selected="True" Value="PO">Item PO</asp:ListItem>
                                                                <asp:ListItem>PL</asp:ListItem>
                                                                <asp:ListItem Value="GR">Penerimaan</asp:ListItem>
                                                            </asp:RadioButtonList>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <div id="DivPOHEta" runat="server" style="text-align:center;width:100%">
                                                                <asp:Panel ID="PanPOHEta" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:850px;left:20%">
                                        
                                                                    <table style="width:75%;margin:auto">
                                                                        <tr style="text-align:left">
                                                                            <td colspan="3">
                                                                                <br />
                                                                                <asp:Label ID="LblPOHEtaProgress" runat="server" Font-Size="17px"></asp:Label>
                                                                                <br />
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width:175px">
                                                                                <table>
                                                                                    <tr>
                                                                                        <td>
                                                                                            ETA&nbsp;
                                                                                        </td>
                                                                                        <td>
                                                                                            <asp:TextBox ID="TxtPOHEta" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="15" Width="100px"></asp:TextBox>
                                                                                        </td>
                                                                                        <td>
                                                                                            <asp:Label ID="LblMsgPOHEta" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                            <td style="text-align:right;width:275px">
                                                                                <asp:Button ID="BtnPOHEtaYes" runat="server" Font-Bold="True" Height="35px" OnClientClick="fsDisableYesPOHEta();" Text="Yes" Width="115px" />
                                                                            </td>
                                                                            <td style="text-align:left">
                                                                                <asp:Button ID="BtnPOHEtaNo" runat="server" CssClass="no" Font-Bold="True" Height="35px" Text="No" Width="145px" />
                                                                            </td>
                                                                        </tr>
                                                                        <tr style="height:65px">
                                                                            <td></td>
                                                                        </tr>
                                                                    </table>
                                                                    <br />                                        
                                                                    <br />
                                                                    <br />
                                                                </asp:Panel>
                                                            </div>
                                                            <div id="DivPOHClo" runat="server" style="text-align:center;width:100%">
                                                                <asp:Panel ID="PanPOHClo" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%">
                                        
                                                                    <table style="width:75%;margin:auto">
                                                                        <tr style="text-align:left">
                                                                            <td colspan="2">
                                                                                <br />
                                                                                <asp:Label ID="LblPOHCloProgress" runat="server" Font-Size="17px"></asp:Label>
                                                                                <br />
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td>
                                                                                <table>
                                                                                    <tr>
                                                                                        <td>
                                                                                            Close Note
                                                                                        </td>
                                                                                        <td>
                                                                                            <asp:TextBox ID="TxtPOHCloNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="450" Width="300px" Height="35px" TextMode="MultiLine"></asp:TextBox>
                                                                                        </td>
                                                                                        <td>
                                                                                            <asp:Label ID="LblMsgPOHClo" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                        <tr style="height:65px">
                                                                            <td style="text-align:center">
                                                                                <table>
                                                                                    <tr>
                                                                                        <td>
                                                                                            <asp:Button ID="BtnPOHCloYes" runat="server" OnClientClick="fsDisableYesPOHClo();" Text="Yes" Width="115px" Font-Bold="True" Height="35px" />
                                                                                        </td>
                                                                                        <td>
                                                                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                                            <asp:Button ID="BtnPOHCloNo" runat="server" CssClass="no" Text="No" Width="145px" Font-Bold="True" Height="35px" />
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <br />                                        
                                                                    <br />
                                                                    <br />
                                                                </asp:Panel>
                                                            </div>
                                                            <asp:GridView ID="GrvPOD" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <Columns>
                                                                    <asp:BoundField DataField="BRG" HeaderText="Kode Barang" HtmlEncode="false">
                                                                        <HeaderStyle Width="80px" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang">
                                                                        <HeaderStyle Width="245px" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="QTY" DataFormatString="{0:n0}" HeaderText="Qty">
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="QTY_PL" DataFormatString="{0:n0}" HeaderText="Qty PL">
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="QTY_RCV" DataFormatString="{0:n0}" HeaderText="Total Qty Receive">
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="vPODOID" HeaderText="PO Detail&lt;br /&gt;OID" HtmlEncode="false">
                                                                        <HeaderStyle Width="75px" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="POHOID" HeaderText="PO Header&lt;br /&gt;OID" HtmlEncode="false">
                                                                        <HeaderStyle Width="75px" CssClass="myDisplayNone" />
                                                                        <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
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
                                                            <asp:Panel ID="PanPL" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%" Visible="false">
                                                                <asp:GridView ID="GrvPLH" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:BoundField DataField="OID" HeaderText="ID Transaksi">
                                                                            <HeaderStyle Width="65px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <asp:ButtonField CommandName="PLNo" DataTextField="PLNo" HeaderText="Nomor PL" Text="Button">
                                                                            <HeaderStyle Width="125px" />
                                                                            <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                                        </asp:ButtonField>
                                                                        <asp:BoundField DataField="vPLDate" HeaderText="Tanggal PL">
                                                                            <HeaderStyle Width="65px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="RcvPONo" HeaderText="No Terima">
                                                                            <HeaderStyle Width="100px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal&lt;br /&gt;Terima" HtmlEncode="false">
                                                                            <HeaderStyle Width="65px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                                            <HeaderStyle Width="100px" />
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
                                                                <br />
                                                                <asp:GridView ID="GrvPLD" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" HeaderStyle-CssClass="StickyHeader" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                                            <ItemStyle CssClass="myDisplayNone" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                                                            <HeaderStyle Width="65px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                                            <HeaderStyle Width="345px" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="PLDQty" DataFormatString="{0:n0}" HeaderText="Qty">
                                                                            <HeaderStyle Width="60px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="PLDCtn" DataFormatString="{0:n0}" HeaderText="Ctn">
                                                                            <HeaderStyle Width="60px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
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
                                                            <asp:Panel ID="PanGR" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%" Visible="false">
                                                                <asp:GridView ID="GrvRcvPOH" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <%-- 0 --%>
                                                                        <asp:BoundField DataField="OID" HeaderText="ID Transaksi">
                                                                            <HeaderStyle Width="65px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <%-- 1 --%>
                                                                        <asp:ButtonField CommandName="RcvPONo" DataTextField="RcvPONo" HeaderText="Nomor Penerimaan" Text="Button">
                                                                            <HeaderStyle Width="125px" />
                                                                            <ItemStyle ForeColor="#0033CC" />
                                                                        </asp:ButtonField>
                                                                        <%-- 2 --%>
                                                                        <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal Penerimaan">
                                                                            <HeaderStyle Width="70px" />
                                                                        </asp:BoundField>
                                                                        <%-- 3 --%>
                                                                        <asp:BoundField DataField="RcvPORefNo" HeaderText="PL/DO">
                                                                            <HeaderStyle Width="70px" />
                                                                        </asp:BoundField>
                                                                        <%-- 4 --%>
                                                                        <asp:BoundField DataField="RcvPOTypeName" HeaderText="Import/&lt;br /&gt;Local" HtmlEncode="false">
                                                                            <HeaderStyle Width="70px" />
                                                                        </asp:BoundField>
                                                                        <%-- 5 --%>
                                                                        <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                                            <HeaderStyle Width="100px" />
                                                                        </asp:BoundField>
                                                                        <%-- 6 --%>
                                                                        <asp:BoundField DataField="TransStatus" HeaderText="Status">
                                                                            <HeaderStyle Width="100px" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <%-- 7 --%>
                                                                        <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                                            <HeaderStyle Width="100px" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <%-- 8 --%>
                                                                        <asp:BoundField DataField="RcvPORefTypeOID" HeaderText="RcvPORefTypeOID">
                                                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                                                            <ItemStyle HorizontalAlign="Left" CssClass="myDisplayNone" />
                                                                        </asp:BoundField>
                                                                        <%-- 9 --%>
                                                                        <asp:BoundField DataField="RcvPORefOID" HeaderText="RcvPORefOID">
                                                                            <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                                                            <ItemStyle HorizontalAlign="Left" CssClass="myDisplayNone" />
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
                                                                <br />
                                                                <asp:GridView ID="GrvRcvPOSumm" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <%-- 0 --%>
                                                                        <asp:BoundField DataField="POHOID" HeaderText="POHOID">
                                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                                            <ItemStyle CssClass="myDisplayNone" />
                                                                        </asp:BoundField>
                                                                        <%-- 1 --%>
                                                                        <asp:BoundField DataField="PO_NO" HeaderText="No.PO">
                                                                            <HeaderStyle Width="100px" />
                                                                        </asp:BoundField>
                                                                        <%-- 2 --%>
                                                                        <asp:ButtonField CommandName="BRG" DataTextField="BRG" HeaderText="Kode Barang" Text="Button">
                                                                            <HeaderStyle Width="75px" />
                                                                            <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                                        </asp:ButtonField>
                                                                        <%-- 3 --%>
                                                                        <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang" HtmlEncode="false">
                                                                            <HeaderStyle Width="200px" />
                                                                        </asp:BoundField>
                                                                        <%-- 4 --%>
                                                                        <asp:BoundField DataField="vSumPLQty" DataFormatString="{0:n0}" HeaderText="Qty PL">
                                                                            <HeaderStyle Width="50px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <%-- 5 --%>
                                                                        <asp:BoundField DataField="vSumPOQty" DataFormatString="{0:n0}" HeaderText="Qty PO">
                                                                            <HeaderStyle CssClass="myDisplayNone" Width="50px" />
                                                                            <ItemStyle CssClass="myDisplayNone" HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <%-- 6 --%>
                                                                        <asp:BoundField DataField="vSumRetDRealQty" DataFormatString="{0:n0}" HeaderText="Qty Retur">
                                                                            <HeaderStyle CssClass="myDisplayNone" Width="50px" />
                                                                            <ItemStyle CssClass="myDisplayNone" HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <%-- 7 --%>
                                                                        <asp:BoundField DataField="vSumRcvPOScanQty" DataFormatString="{0:n0}" HeaderText="Qty Receive">
                                                                            <HeaderStyle Width="50px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <%-- 8 --%>
                                                                        <asp:BoundField DataField="vRcvPOQty_Total" DataFormatString="{0:n0}" HeaderText="Total&lt;br /&gt;Qty PO Receive" HtmlEncode="false">
                                                                            <HeaderStyle Width="50px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <%-- 9 --%>
                                                                        <asp:BoundField DataField="vQtyVarian" DataFormatString="{0:n0}" HeaderText="Selisih" HtmlEncode="false">
                                                                            <HeaderStyle Width="50px" />
                                                                            <ItemStyle HorizontalAlign="Right" />
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
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="PanPOUpload" runat="server" style="height:525px" Visible="false">
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td style="width:10px"></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>UPLOAD XLS PO</strong></td>
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
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>HISTORY UPLOAD XLS PO</strong></td>
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
                                                            <asp:Button ID="BtnData" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Data PO" Width="112px" Font-Bold="True" />
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

                <asp:Panel ID="PanPOSAP" runat="server" style="height:525px" Visible="false">
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td style="width:10px"></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>GET PO SAP</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Company</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstPOSapCompany" runat="server" style="height: 20px" Width="300px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblPOSapCompany" runat="server" ForeColor="Red"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>                                        
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblPOSapError" runat="server" ForeColor="Red"></asp:Label>
                                            <asp:Label ID="LblPOSapProses" runat="server" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnPOSap" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" OnClientClick="fsShowInProgressSAP('Calling SAP API')" Text="GET PO SAP" Width="145px" />
                                                    </td>                                                    
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>                        
                            </td>
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <asp:Panel ID="Panel2" class="myPanelGreyNSa" runat="server" Width="100%" style="height:500px" Visible="True">
                                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                                        <tr>
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>HISTORY UPLOAD XLS PO</strong></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table>
                                                    <tr>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnPOSapFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        </td>
                                                        <td>
                                                            <asp:Label ID="LblMsgPOSapFindError" runat="server" ForeColor="Red"></asp:Label>
                                                        </td>
                                                        <td style="width:35px">
                                                            <asp:Button ID="BtnPOSapPO" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Data PO" Width="112px" Font-Bold="True" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr>
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvPOSap" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:ButtonField CommandName="OID" DataTextField="OID" HeaderText="OID">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:ButtonField>
                                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="XlsFileName" HeaderText="Api Result File Name">
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
                                                        <asp:BoundField DataField="vStatusSuccess" HeaderText="Status">
                                                            <HeaderStyle Width="45px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="StatusMessage" HeaderText="Status Message">
                                                            <HeaderStyle Width="345px" />
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
        .auto-style2 {
            height: 18px;
        }
    </style>
</asp:Content>
