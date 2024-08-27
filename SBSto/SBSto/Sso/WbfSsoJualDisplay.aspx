<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoJualDisplay.aspx.vb" Inherits="SBSto.WbfSsoJualDisplay" Title="SB WMS : Invoice" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListUploadStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListUploadEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListUploadStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListUploadEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowProgress(vriProses) {
            document.getElementById("<%= BtnRefreshCustomer.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnNotaFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblMsgRefreshCustomer.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
        }
    </script>
</head>
    <body>
        
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>INVOICE</strong></td>                                
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px">&nbsp;</td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="display:block;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr style="vertical-align:top">
                                                <td style="width:450px">
                                                    <table>
                                                        <tr>
                                                            <td>Company</td>
                                                            <td>:</td>
                                                            <td>
                                                                <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="350px">
                                                                </asp:DropDownList>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>No. Invoice</td>
                                                            <td></td>
                                                            <td>
                                                                <asp:TextBox ID="TxtNotaNo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Customer</td>
                                                            <td></td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListCustomer" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>
                                                                <asp:Label ID="LblMsgRefreshCustomer" runat="server" ForeColor="#0066FF"></asp:Label>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td></td>
                                                            <td></td>
                                                            <td>
                                                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>Gudang</td>
                                                            <td></td>
                                                            <td>
                                                                <asp:DropDownList ID="DstWarehouse" runat="server" style="height: 20px" Width="250px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td rowspan="2">
                                                                <asp:Button ID="BtnRefreshCustomer" runat="server" class="myButtonFinda" Enabled="False" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="#0066FF" Height="30px" OnClientClick="fsShowProgress('Refresh');" Text="REFRESH CUSTOMER" Visible="False" Width="125px" />
                                                                <asp:Button ID="BtnNotaFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                                            </td>
                                                            <td rowspan="2">
                                                                <asp:Button ID="BtnSummary" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Summary" Width="100px" />
                                                            </td>
                                                            <td rowspan="2">
                                                                <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Periode</td>
                                                            <td></td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                &nbsp; s/d
                                                                <asp:TextBox ID="TxtListEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Upload</td>
                                                            <td></td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListUploadStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                &nbsp; s/d
                                                                <asp:TextBox ID="TxtListUploadEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                            </td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkListIsPicklist_Yes" runat="server" Checked="True" ForeColor="#0066FF" Text="Sudah  Picklist" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkListIsPicklist_No" runat="server" Checked="True" ForeColor="#009933" Text="Belum Picklist" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                                <HeaderStyle Width="50px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="no_nota" DataTextField="no_nota" Text="Button" HeaderText="No. Invoice">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vtanggal" HeaderText="Tanggal Invoice">
                                                                <HeaderStyle Width="80px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="kode_cust" HeaderText="Kode Customer">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CUSTOMER" HeaderText="Customer" HtmlEncode="false">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang" HtmlEncode="false">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="ALAMAT" HeaderText="Alamat" HtmlEncode="false">
                                                                <HeaderStyle Width="245px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="kota" HeaderText="Kota">
                                                                <HeaderStyle Width="105px" CssClass="myDisplayNone" />
                                                                <ItemStyle Width="105px" CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vDOTitip" HeaderText="DO Titip">
                                                                <HeaderStyle Width="55px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vIsPickListClosed" HeaderText="Sudah<br />Picklist" HtmlEncode="false">
                                                                <HeaderStyle Width="55px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vUploadDatetime" HeaderText="Upload Datetime">
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PriorityName" HeaderText="Prioritas">
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaPRIODatetime" HeaderText="Prioritas Datetime">
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaCancel" HeaderText="Cancel">
                                                                <HeaderStyle Width="45px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NotaCancelNote" HeaderText="Cancel Note">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NotaCancelReturNo" HeaderText="No. Retur">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaCancelDatetime" HeaderText="Cancel Datetime">
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaTKF" HeaderText="Tukar Faktur">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vNotaNo_Baru" HeaderText="Faktur Baru">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NotaRJSNo" HeaderText="No. RJS">
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
                                            <tr>
                                                <td>
                                                    <div id="DivNota" runat="server" >
                                                        <asp:Panel ID="PanNota" runat="server" style="display:block;width:1700px;height:580px;margin-top:25px" Visible="True" BorderStyle="Solid" BackColor="LightGray">
                                                            <table style="width:100%;font-family: tahoma;font-size:11px">
                                                                <tr>
                                                                    <td>
                                                                        <asp:Label ID="LblNota" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DATA INVOICE</asp:Label>
                                                                    </td>
                                                                    <td>
                                                                        <asp:Button ID="BtnNotaClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" Visible="False" />
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table style="width:100%;font-family: tahoma;font-size:11px">
                                                                <tr>
                                                                    <td style="width:10px"></td>
                                                                    <td style="width:115px">Nomor Invoice</td>
                                                                    <td style="width:20px;text-align:center">:</td>
                                                                    <td style="width:145px">
                                                                        <asp:TextBox ID="TxtNotaNo1" runat="server" BackColor="#E0E0E0" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="210px"></asp:TextBox>
                                                                    </td>                            
                                                                    <td style="width:25px"></td>
                                                                    <td style="width:100px">Company</td>
                                                                    <td style="width:20px;text-align:center">:</td>
                                                                    <td style="width:115px">
                                                                        <asp:TextBox ID="TxtNotaCompany" runat="server" BackColor="#E0E0E0" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="95px"></asp:TextBox>
                                                                    </td>
                                                                    <td style="width:25px"></td>
                                                                    <td>&nbsp;</td>
                                                                    <td style="text-align:center">&nbsp;</td>
                                                                    <td>
                                                                        &nbsp;</td>
                                                                </tr>
                                                                <tr>
                                                                    <td></td>
                                                                    <td>Tanggal Invoice</td>
                                                                    <td style="text-align:center">:</td>
                                                                    <td>
                                                                        <asp:TextBox ID="TxtNotaDate" runat="server" BackColor="#E0E0E0" Font-Names="Tahoma" Font-Size="12px" Width="100px"></asp:TextBox>
                                                                    </td>
                                                                    <td></td>
                                                                    <td>Customer</td>                         
                                                                    <td style="text-align:center">:</td>
                                                                    <td>
                                                                        <asp:TextBox ID="TxtNotaCustomer" runat="server" BackColor="#E0E0E0" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="210px"></asp:TextBox>
                                                                    </td>
                                                                    <td></td>
                                                                    <td>
                                                                        <asp:CheckBox ID="ChkSummary" runat="server" AutoPostBack="True" Font-Bold="True" Font-Size="12px" ForeColor="#0066FF" Text="SUMMARY" />
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table>
                                                                <tr style="vertical-align:top">
                                                                    <td></td>
                                                                    <td>
                                                                        <asp:GridView ID="GrvNota" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True">
                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                            <Columns>
                                                                                <asp:BoundField DataField="NotaHOID" HeaderText="Nota HOID">
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vNotaDOID" HeaderText="Detail OID">
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="KODE_BARANG" HeaderText="Kode Barang">
                                                                                    <HeaderStyle Width="80px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang">
                                                                                    <HeaderStyle Width="250px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="QTY" HeaderText="Qty" DataFormatString="{0:n0}" >
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="QTYBONUS" HeaderText="Qty Bonus" DataFormatString="{0:n0}" >
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="QtyOnPKDOT" HeaderText="Qty<br />Perintah Kirim<br />DO Titip" DataFormatString="{0:n0}" HtmlEncode="false" >
                                                                                    <HeaderStyle Width="80px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="QtyOnPickList" HeaderText="Qty<br />Dalam Picklist" DataFormatString="{0:n0}" HtmlEncode="false" >
                                                                                    <HeaderStyle Width="80px" CssClass="myDisplayNone" />
                                                                                    <ItemStyle HorizontalAlign="Right" CssClass="myDisplayNone" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="SATUAN" HeaderText="Satuan">
                                                                                    <HeaderStyle Width="45px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="NO_REF" HeaderText="No Ref">
                                                                                    <HeaderStyle Width="45px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="SALESMAN" HeaderText="Sales">
                                                                                    <HeaderStyle Width="65px" />
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
                                                                        <asp:Label ID="LblDetailNota_ByBarang" runat="server" Font-Bold="False" Font-Size="14px" ForeColor="#0066FF">TOTAL QTY INVOICE PER BARANG</asp:Label>
                                                                        <br />
                                                                        <asp:GridView ID="GrvNota_ByBarang" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True">
                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                            <Columns>
                                                                                <asp:BoundField DataField="NotaHOID" HeaderText="Nota HOID">
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="KodeBarang" HeaderText="Kode Barang">
                                                                                    <HeaderStyle Width="80px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang">
                                                                                    <HeaderStyle Width="250px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="TotalQty" HeaderText="Total Qty" DataFormatString="{0:n0}" >
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="TotalQtyBonus" HeaderText="Total Qty Bonus" DataFormatString="{0:n0}" >
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="TotalQtyOnPKDOT" HeaderText="Qty<br />Perintah Kirim<br />DO Titip" DataFormatString="{0:n0}" HtmlEncode="false" >
                                                                                    <HeaderStyle Width="80px" CssClass="myDisplayNone" />
                                                                                    <ItemStyle HorizontalAlign="Right" CssClass="myDisplayNone" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="TotalQtyOnPickList" HeaderText="Qty<br />Dalam Picklist" DataFormatString="{0:n0}" HtmlEncode="false" >
                                                                                    <HeaderStyle Width="80px" />
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
                                                                        <asp:GridView ID="GrvSumNota" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" Visible="False">
                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                            <Columns>
                                                                                <asp:BoundField DataField="vKodeBarang" HeaderText="Kode Barang">
                                                                                    <HeaderStyle Width="80px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vNamaBarang" HeaderText="Nama Barang">
                                                                                    <HeaderStyle Width="150px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vSatuan" HeaderText="Satuan">
                                                                                    <HeaderStyle Width="45px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtyNota" DataFormatString="{0:n0}" HeaderText="Qty">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtyNotaBonus" DataFormatString="{0:n0}" HeaderText="Qty<br />Bonus" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtySKK_Closed" DataFormatString="{0:n0}" HeaderText="Qty SKK<br />Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtySKK_NotClosed" DataFormatString="{0:n0}" HeaderText="Qty SKK<br />Not Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtySJ_Closed" DataFormatString="{0:n0}" HeaderText="Qty SJ<br />Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtySJ_NotClosed" DataFormatString="{0:n0}" HeaderText="Qty SJ<br />Not Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtyPL_Closed" DataFormatString="{0:n0}" HeaderText="Qty PL<br />Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtyPL_NotClosed" DataFormatString="{0:n0}" HeaderText="Qty PL<br />Not Closed" HtmlEncode="false">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vQtySisa" DataFormatString="{0:n0}" HeaderText="Qty Sisa">
                                                                                    <HeaderStyle Width="70px" />
                                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                                </asp:BoundField>
                                                                                <asp:ButtonField CommandName="vDetail" DataTextField="vDetail" Text="Button" HeaderText="">
                                                                                    <HeaderStyle Width="45px" />
                                                                                    <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                                                </asp:ButtonField>
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
                                                                    <td>

                                                                        <asp:Label ID="LblDetailNota" runat="server" Font-Bold="False" Font-Size="14px" ForeColor="#0066FF">DATA INVOICE</asp:Label>
                                                                        <br />
                                                                        <asp:GridView ID="GrvDetailNota" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True">
                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                            <Columns>
                                                                                <asp:BoundField DataField="vTransOID" HeaderText="vTransOID">
                                                                                    <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                                                                    <ItemStyle Width="100px" CssClass="myDisplayNone" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransCode" HeaderText="vTransCode">
                                                                                    <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                                                                    <ItemStyle Width="100px" CssClass="myDisplayNone" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransName" HeaderText="Transaksi">
                                                                                    <HeaderStyle Width="100px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransNo" HeaderText="No. Transaksi">
                                                                                    <HeaderStyle Width="100px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransDate" HeaderText="Tanggal<br />Transaksi" HtmlEncode="false">
                                                                                    <HeaderStyle Width="100px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransStatus" HeaderText="Status<br />Transaksi" HtmlEncode="false">
                                                                                    <HeaderStyle Width="100px" CssClass="myDisplayNone" />
                                                                                    <ItemStyle Width="100px" CssClass="myDisplayNone" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransStatusName" HeaderText="Status<br />Transaksi" HtmlEncode="false">
                                                                                    <HeaderStyle Width="100px" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="vTransQty" DataFormatString="{0:n0}" HeaderText="Qty<br />Transaksi" HtmlEncode="false">
                                                                                    <HeaderStyle Width="80px" />
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

                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </asp:Panel>
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>        
    </body>
</asp:Content>
