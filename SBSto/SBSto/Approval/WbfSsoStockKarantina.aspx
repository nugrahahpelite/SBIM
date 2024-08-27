<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoStockKarantina.aspx.vb" Inherits="SBSto.WbfSsoStockKarantina" MasterPageFile="~/SBSto.Master" Title="SB WMS : Stock Karantina" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Stock Karantina</title>

    <script type="text/javascript">    
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = "Wait...Process in Progress.........";
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
                    <table style="width: 90%;font-family: tahoma;font-size:11px">
                        <tr>
                            <td style="font-size:15px;height:28px" colspan="3"><strong>STOCK KARANTINA</strong></td>
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
                                                </tr>
                                                <tr>
                                                    <td></td>
                                                    <td></td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkSt_Baru" runat="server" Checked="True" Text="Baru" />
                                                        &nbsp;&nbsp;&nbsp;
                                                        <asp:CheckBox ID="ChkSt_OnPutaway" runat="server" Text="On Putaway" />
                                                        &nbsp;
                                                        <asp:CheckBox ID="ChkSt_PutawayDone" runat="server" Checked="True" Text="Putaway Done" />
                                                        &nbsp;
                                                        <asp:CheckBox ID="ChkSt_Approved" runat="server" Text="Sudah Approved" />
                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                        <asp:CheckBox ID="ChkOSOnly" runat="server" ForeColor="#0066FF" Text="OUTSTANDING ONLY" />
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
                                            <div id="DivConfirm" runat="server" style="text-align:center;width:100%">
                                                <asp:Panel ID="PanConfirm" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%">
                                        
                                                    <table style="width:75%;margin:auto">
                                                        <tr style="text-align:left">
                                                            <td colspan="2">
                                                                <br />
                                                                <asp:Label ID="LblConfirmMessage" runat="server" Text="Anda Yakin Approve ?" Font-Size="17px"></asp:Label>
                                                                <br />
                                                                <asp:Label ID="LblConfirmProgress" runat="server" Font-Size="17px">Approve Tidak Dapat Dibatalkan</asp:Label>
                                                                <br />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <table runat="server" id="tbConfirmNote" visible="false">
                                                                    <tr>
                                                                        <td>
                                                                            <asp:Label ID="LblConfirmNote" runat="server" Text="Note : "></asp:Label>
                                                                        </td>
                                                                        <td>
                                                                            <asp:TextBox ID="TxtConfirmNote" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="270px" Height="35px" TextMode="MultiLine"></asp:TextBox>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="LblConfirmWarning" runat="server" ForeColor="#FF0066"></asp:Label>                                                    
                                                            </td>
                                                        </tr>
                                                        <tr style="height:65px">
                                                            <td style="text-align:center">
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:Button ID="BtnConfirmYes" runat="server" OnClientClick="fsDisableYesConfirmSto();" Text="Yes" Width="115px" Font-Bold="True" Height="35px" />
                                                                        </td>
                                                                        <td>
                                                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                                            <asp:Button ID="BtnConfirmNo" runat="server" CssClass="no" Text="No" Width="145px" Font-Bold="True" Height="35px" />
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
                                        <%-- 0 --%>
                                        <asp:BoundField DataField="vStorageOID" HeaderText="Storage<br />OID" HtmlEncode="false">
                                            <HeaderStyle Width="65px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 1 --%>
                                        <asp:BoundField DataField="vStorageStockOID" HeaderText="Storage Stock<br />OID" HtmlEncode="false">
                                            <HeaderStyle Width="65px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 2 --%>
                                        <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Location" HtmlEncode="false">
                                            <HeaderStyle Width="115px" />
                                        </asp:BoundField>
                                        <%-- 3 --%>
                                        <asp:BoundField DataField="TransCode" HeaderText="TransCode" HtmlEncode="false">
                                            <HeaderStyle Width="50px" CssClass="myDisplayNone" />
                                            <ItemStyle Width="50px" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 4 --%>
                                        <asp:BoundField DataField="TransName" HeaderText="Transaksi" HtmlEncode="false">
                                            <HeaderStyle Width="50px" />
                                        </asp:BoundField>
                                        <%-- 5 --%>
                                        <asp:BoundField DataField="TransOID" HeaderText="ID Transaksi" HtmlEncode="false">
                                            <HeaderStyle Width="50px" />
                                        </asp:BoundField>
                                        <%-- 6 --%>
                                        <asp:BoundField DataField="CompanyCode" HeaderText="Company" HtmlEncode="false">
                                            <HeaderStyle Width="50px" />
                                        </asp:BoundField>
                                        <%-- 7 --%>
                                        <asp:BoundField DataField="RcvPOHOID" HeaderText="RcvPOHOID" HtmlEncode="false">
                                            <HeaderStyle Width="50px" CssClass="myDisplayNone" />
                                            <ItemStyle Width="50px" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 8 --%>
                                        <asp:BoundField DataField="RcvPONo" HeaderText="No.<br />Terima" HtmlEncode="false">
                                            <HeaderStyle Width="120px" />
                                        </asp:BoundField>
                                        <%-- 9 --%>
                                        <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal<br />Terima" HtmlEncode="false">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <%-- 10 --%>
                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode<br />Barang" HtmlEncode="false">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <%-- 11 --%>
                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                            <HeaderStyle Width="220px" />
                                        </asp:BoundField>
                                        <%-- 12 --%>
                                        <asp:BoundField DataField="NoteKarantina" HeaderText="Note" HtmlEncode="false">
                                            <HeaderStyle Width="145px" />
                                        </asp:BoundField>
                                        <%-- 13 --%>
                                        <asp:BoundField DataField="QtyKarantina" HeaderText="Qty<br />Karantina" DataFormatString="{0:n0}" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 14 --%>
                                        <asp:BoundField DataField="QtyKrRelease" HeaderText="Qty<br />Release" DataFormatString="{0:n0}" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 15 --%>
                                        <asp:BoundField DataField="QtyKrReceive" HeaderText="Qty<br />Receive" DataFormatString="{0:n0}" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 16 --%>
                                        <asp:BoundField DataField="vQtyKrOutstanding" HeaderText="Qty<br />Outstanding" DataFormatString="{0:n0}" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <%-- 17 --%>
                                        <asp:BoundField DataField="TransStatus" HeaderText="TransStatus" HtmlEncode="false">
                                            <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                            <ItemStyle Width="45px" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 18 --%>
                                        <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" HtmlEncode="false">
                                            <HeaderStyle Width="85px" />
                                        </asp:BoundField>
                                        <%-- 19 --%>
                                        <asp:ButtonField CommandName="vApprove" DataTextField="vApprove" Text="Button" HeaderText="Approve">
                                            <HeaderStyle Width="125px" />
                                            <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                        </asp:ButtonField>
                                        <%-- 20 --%>
                                        <asp:BoundField DataField="vCreationDatetime" HeaderText="Creation" HtmlEncode="false">
                                            <HeaderStyle Width="145px" />
                                        </asp:BoundField>
                                        <%-- 21 --%>
                                        <asp:BoundField DataField="vPutawayDone" HeaderText="Putaway Done" HtmlEncode="false">
                                            <HeaderStyle Width="145px" />
                                        </asp:BoundField>
                                        <%-- 22 --%>
                                        <asp:BoundField DataField="vApproved" HeaderText="Approval" HtmlEncode="false">
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
