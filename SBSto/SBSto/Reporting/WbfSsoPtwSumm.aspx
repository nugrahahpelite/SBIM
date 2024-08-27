<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoPtwSumm.aspx.vb" Inherits="SBSto.WbfSsoPtwSumm" MasterPageFile="~/SBSto.Master" Title="SB WMS : Summary Putaway" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Summary Putaway</title>
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
</head>
    <body>
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnProOK" />
                <asp:PostBackTrigger ControlID="BtnXLS" />
            </Triggers>  
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td style="font-size:15px;height:28px" colspan="3"><strong>SUMMARY PUTAWAY</strong></td>
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
                                                            </tr>
                                                        </table>                                                        
                                                    </td>
                                                    <td rowspan="2">
                                                        <asp:Button ID="BtnFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                        &nbsp;&nbsp;&nbsp;&nbsp;
                                                        <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>Company</td>
                                                    <td>:</td>
                                                    <td>
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="350px"></asp:DropDownList>
                                                                </td>
                                                                <td>
                                                                    <asp:Label ID="LblMsgListCompany" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                </td>
                                                            </tr>
                                                        </table>                                                        
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>Periode</td>
                                                    <td>:</td>
                                                    <td>                                                        
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:TextBox ID="TxtListStart" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                </td>
                                                                <td>s/d</td>
                                                                <td>
                                                                    <asp:TextBox ID="TxtListEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td>
                                                       
                                                        <asp:Button ID="BtnPreview" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Preview   " Width="112px" />
                                                    </td>
                                                    <td></td>
                                                </tr>
                                                <tr>
                                                    <td>Barang</td>
                                                    <td>:</td>
                                                    <td>
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:TextBox ID="TxtListBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="20" Width="95px"></asp:TextBox>
                                                                    <asp:TextBox ID="TxtListBrgName" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="190px"></asp:TextBox>
                                                                </td>
                                                                <td>
                                                                    <asp:Button ID="BtnListBrgCode" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                                </td>
                                                                <td>
                                                                    <asp:Label ID="LblMsgListBrg" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />
                                                        <asp:Label ID="LblProgressXLS" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td></td>
                                                    <td></td>
                                                    <td>
                                                        &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
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
                                            <div id="DivPrOption" runat="server" style="width:100%">
                                                <asp:Panel ID="PanPrOption" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;width:450px;left:20%">
                                                    <br />
                                                    <br />
                                                    <table style="width:75%;margin:auto">
                                                        <tr>
                                                            <td>
                                                                <asp:RadioButtonList ID="RdbProXls" runat="server" RepeatDirection="Horizontal">
                                                                    <asp:ListItem Selected="True">Pdf</asp:ListItem>
                                                                </asp:RadioButtonList>
                                                            </td>
                                                        </tr>
                                                        <tr style="height:65px">
                                                            <td style="text-align:center">
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:Button ID="BtnProOK" runat="server" Text="OK" Width="115px" Font-Bold="True" Height="35px" />
                                                                        </td>
                                                                        <td>
                                                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                            <asp:Button ID="BtnProCancel" runat="server" CssClass="no" Text="Cancel" Width="145px" Font-Bold="True" Height="35px" />
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
                                            <div id="DivPreview" runat="server" >
                                                <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="z-index:300;display:block;width:75%;margin-top:45px">
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
                                </table>
                            </td>                
                        </tr>     
                        <tr>
                             <td>
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px"  ReadOnly="True" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TxtListFind" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Visible="false"></asp:TextBox>
                             </td>
                            <td>
                                
                            </td>
                            </tr>
                        <tr>
                            <td>
                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <%-- 0 --%>
                                        <asp:BoundField DataField="TransCode" HeaderText="TransCode" HtmlEncode="false">
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TransName" HeaderText="Transaksi Name" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vPtwCompanyCode" HeaderText="Company" HtmlEncode="false">
                                            <HeaderStyle Width="55px" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="vPtwNo" DataTextField="vPtwNo" HeaderText="No.Putaway" Text="Button">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle ForeColor="#0033CC" />
                                        </asp:ButtonField>
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
                                        <asp:BoundField DataField="PCKNo" HeaderText="No.Picking" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PCLRefHNo" HeaderText="No.Invoice" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                            <HeaderStyle Width="75px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="RcvPOHOID" HeaderText="RcvPOHOID" HtmlEncode="false">
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="RcvPONo" HeaderText="No.Penerimaan" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                            <HeaderStyle Width="145px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vSumPtwScan1Qty" DataFormatString="{0:n0}" HeaderText="Qty<br />Scan 1" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vPtwReceiveQty" DataFormatString="{0:n0}" HeaderText="Qty<br />Diterima" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vSumPtwScan2Qty" DataFormatString="{0:n0}" HeaderText="Qty<br />Scan 2" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CreationDatetime" HeaderText="CreationDatetime" >
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
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
                    <table>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HdfDetailRowIdx" runat="server" />
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

