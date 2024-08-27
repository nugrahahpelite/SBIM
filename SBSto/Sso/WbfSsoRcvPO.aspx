<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoRcvPO.aspx.vb" Inherits="SBSto.WbfSsoRcvPO" Title="SB WMS - Penerimaan Pembelian" MasterPageFile="~/SBSto.Master" %>
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
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowInProgress(vriProcess) {
            document.getElementById("<%= LblXlsProses.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
            document.getElementById("<%= LblMsgXlsProsesError.ClientID%>").innerText = "";
        }
        function fsShowFindProgress() {
            document.getElementById("<%= BtnFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblFindProgress.ClientID%>").innerText = "Sedang Proses...";
        }
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
        }
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = document.getElementById("<%= HdfProcess.ClientID%>").value + " in Progress.........";
        }
    </script>
    <style type="text/css">=
        .gridViewPager td
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
        }
        .StickyHeader th {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            top: 0
        }
        .GrvDetail {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
        }
        .modalPopup
        {
            background-color: #f4f4f4;
            width: 250px;
            border: 1px solid rgba(0, 0, 0, 0.5);
            border-radius: 12px;
            padding: 0;
            box-shadow:0 5px 15px rgba(0, 0, 0, 0.5);
        }
        .Grid td
        {
            padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
            line-height:200%
        }
        .Grid th
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            line-height:200%
        }
        .ChildGrid td
        {
            background-color: azure !important;
            color: black;
            font-size: 10pt;
            line-height:200%
        }
        .ChildGrid th
        {
            background-color: cadetblue !important;
            color: White;
            font-size: 10pt;
            line-height:200%
        }
        .highlight 
        {
            text-decoration: none;
            color:black;
            background:yellow;
        }
    </style>
</head>
    <body>
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnStatus" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>PENERIMAAN</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td>
                                    <asp:Button ID="BtnApprove" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Approve Penerimaan" Width="145px" Enabled="False" Visible="False" />
                                </td>
                                <td style="width:420px">
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
                                </td>
                                <td style="width:25px">
                                    &nbsp;</td>
                                <td style="width:125px">
                                    &nbsp;</td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    &nbsp;</td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    &nbsp;</td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    &nbsp;</td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    &nbsp;</td>
                                <td style="width:45px"></td>
                                <td style="width:130px">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td></td>
                                <td colspan="6">

                                    <asp:Label ID="LblXlsProses" runat="server" ForeColor="#0066FF"></asp:Label>
                                    <asp:Label ID="LblMsgXlsProsesError" runat="server" ForeColor="Red" Visible="False"></asp:Label>

                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px">&nbsp;</td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="z-index:110;display:block;width:1650px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:85%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td style="width:150px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressFind();" />
                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:70px">Trans. ID :</td>
                                                <td style="width:200px"><asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox></td>                                                
                                                <td  style="width:115px;text-align:right">Nomor Penerimaan : </td>
                                                <td style="width:200px">
                                                    <asp:TextBox ID="TxtListNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="165px"></asp:TextBox>
                                                </td>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>Periode :</td>
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
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="DstListRcvType" runat="server" style="height: 20px" Width="125px">
                                                    </asp:DropDownList>
                                                </td>
                                                <td style="text-align:right">Warehouse :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">Status&nbsp; :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_OnReceive" runat="server" ForeColor="#336600" Text="On Receive" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ReceiveDone" runat="server" ForeColor="#336600" Text="Receive Done" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ReceiveApproved" runat="server" ForeColor="#336600" Text="Receive Appoved" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ReceivePtwProcess" runat="server" ForeColor="#336600" Text="Putaway Process" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ReceivePtwComplete" runat="server" ForeColor="#336600" Text="All Putaway Complete" />
                                                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <table style="width: 100%;margin:auto;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" PageSize="20">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvTypeName" HeaderText="Tipe" >
                                                                <HeaderStyle Width="65px" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="Select" DataTextField="RcvPONo" Text="Button" HeaderText="Nomor Penerimaan">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal Penerimaan" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvPORefNo" HeaderText="PL/DO" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvPOTypeName" HeaderText="Import/<br />Local" HtmlEncode="false">
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvPOCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvPODoneNote" HeaderText="Note Selesai" >
                                                                <HeaderStyle Width="190px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vReceivedDone" HeaderText="Receive Done" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vReceivedApp" HeaderText="Receive App" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" ForeColor="White" Height="35px" Font-Overline="False" />
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
                                <div id="DivConfirm" runat="server" style="text-align:center;width:100%">
                                    <asp:Panel ID="PanConfirm" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;text-align:center;width:750px;left:20%">
                                        
                                        <table style="width:75%;margin:auto">
                                            <tr style="text-align:left">
                                                <td colspan="2">
                                                    <br />
                                                    <asp:Label ID="LblConfirmMessage" runat="server" Text="Anda Yakin ?" Font-Size="17px"></asp:Label>
                                                    <br />
                                                    <asp:Label ID="LblConfirmProgress" runat="server" Font-Size="17px"></asp:Label>
                                                    <br />
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
                                <div id="DivPreview" runat="server" >
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="z-index:300;display:block;width:75%;margin-top:-45px">
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
                    <table runat="server" id="tbTrans" >
                        <tr>
                            <td style="width:10px"></td>
                            <td style="width:145px">No. Penerimaan</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#E2E2E2"></asp:TextBox>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:100px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtCompany" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td style="width:25px"></td>
                            <td>ID Transaksi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="73px" ReadOnly="True"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Tanggal Penerimaan</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px" Enabled="False"></asp:TextBox>
                                &nbsp;</td>                            
                            <td></td>
                            <td>Warehouse</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtWarehouse" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>Status</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="190px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="BtnStatus" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>PL/DO/RETUR</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransRefNo" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>&nbsp;</td>                         
                            <td style="text-align:center">&nbsp;</td>
                            <td colspan="5" style="vertical-align:top">
                                <asp:RadioButtonList ID="RdlRcvType" runat="server" RepeatDirection="Horizontal" Enabled="False" Font-Size="12pt">
                                    <asp:ListItem Value="1">PEMBELIAN</asp:ListItem>
                                    <asp:ListItem Value="2">RETUR</asp:ListItem>
                                    <asp:ListItem Value="3">LAIN-LAIN</asp:ListItem>
                                    <asp:ListItem Value="4">KARANTINA</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>                            
                        </tr>
                        <tr>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td></td>
                            <td style="vertical-align:top">&nbsp;</td>
                            <td style="text-align:center;vertical-align:top">&nbsp;</td>
                            <td style="vertical-align:top">
                                &nbsp;</td>
                            <td style="vertical-align:top">
                                <asp:RadioButtonList ID="RdlRcvPOType" runat="server" Enabled="False" Font-Size="12pt" ForeColor="Blue" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">IMPORT</asp:ListItem>
                                    <asp:ListItem Value="2">LOCAL</asp:ListItem>
                                    <asp:ListItem Value="3">-</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td></td>
                            <td style="vertical-align:top">&nbsp;</td>
                            <td style="text-align:center;vertical-align:top">&nbsp;</td>                            
                        </tr>
                        <tr>
                            <td></td>
                            <td colspan="10">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td>                            
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="10">
                                <table>
                                    <tr>
                                        <td>
                                            <asp:Button ID="BtnFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="Find" Width="55px" OnClientClick="fsShowFindProgress()" />
                                            <asp:Label ID="LblFindProgress" runat="server" ForeColor="#3333FF"></asp:Label>
                                        </td>
                                        <td>Kode/Nama Barang</td>
                                        <td>
                                            <asp:TextBox ID="TxtFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="190px"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                        </tr>
                    </table>
                    <table>
                        <tr style="vertical-align:top">
                            <td style="width:10px"></td>
                            <td>
                                <asp:GridView ID="GrvSumm" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
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
                                        <asp:ButtonField CommandName="BRG" DataTextField="BRG" Text="Button" HeaderText="Kode Barang">
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
                                            <HeaderStyle Width="50px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Right" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 6 --%>
                                        <asp:BoundField DataField="vSumRetDRealQty" DataFormatString="{0:n0}" HeaderText="Qty Retur">
                                            <HeaderStyle Width="50px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Right" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 7 --%>
                                        <asp:BoundField DataField="vSumRcvPOScanQty" DataFormatString="{0:n0}" HeaderText="Qty Receive">
                                            <HeaderStyle Width="50px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <%-- 8 --%>
                                        <asp:BoundField DataField="vRcvPOQty_Total" DataFormatString="{0:n0}" HeaderText="Total<br />Qty PO Receive" HtmlEncode="false">
                                            <HeaderStyle Width="50px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <%-- 9 --%>
                                        <asp:BoundField DataField="vQtyVarian" DataFormatString="{0:n0}" HeaderText="Selisih Qty<br />PL - Receive" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <%-- 10 --%>
                                        <asp:BoundField DataField="vQtyVarian" DataFormatString="{0:n0}" HeaderText="Selisih Qty<br />PO - Total Receive" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
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
                            <td>
                                <asp:Panel runat="server" ID="PanData" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:650px;height:850px" BackColor="White" Visible="false">
                                    <table style="margin-top:15px;margin-left:15px;margin-bottom:15px;width:95%">
                                        <tr>
                                            <td colspan="6" style="text-align:left"><asp:Label ID="LblDataTitle" runat="server" Text="Data Hasil Scan" Font-Bold="True"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td colspan="6" style="background-color:#808080;height:4px">
                                                <asp:Label ID="Label3" runat="server" Height="3px" Visible="False" Width="100%"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="5">
                                                <asp:CheckBox ID="ChkSt_DelNo" runat="server" Checked="True" ForeColor="#336600" Text="Data Masuk" AutoPostBack="True" />
                                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                <asp:CheckBox ID="ChkSt_DelYes" runat="server" ForeColor="Red" Text="Data Dihapus" AutoPostBack="True" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="6">
                                                <asp:GridView ID="GrvData" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <%-- 0 --%>
                                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                            <ItemStyle CssClass="myDisplayNone" />
                                                        </asp:BoundField>
                                                        <%-- 1 --%>
                                                        <asp:BoundField DataField="RcvPOScanQty" DataFormatString="{0:n0}" HeaderText="Qty Receiving">
                                                            <HeaderStyle Width="50px" />
                                                            <ItemStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <%-- 2 --%>
                                                        <asp:BoundField DataField="vRcvPOScanNote" HeaderText="Note.">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <%-- 3 --%>
                                                        <asp:BoundField DataField="vRcvPOScanUser" HeaderText="Receiving By">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <%-- 4 --%>
                                                        <asp:BoundField DataField="vRcvPOScanTime" HeaderText="Receiving Time">
                                                            <HeaderStyle Width="100px" />
                                                            <ItemStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <%-- 5 --%>
                                                        <asp:ButtonField CommandName="vDelItem" DataTextField="vDelItem" Text="Button" HeaderText="">
                                                            <HeaderStyle Width="45px" CssClass="myDisplayNone" />
                                                            <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" CssClass="myDisplayNone" />
                                                        </asp:ButtonField>
                                                        <%-- 6 --%>
                                                        <asp:BoundField DataField="vRcvPOScanDeleted" HeaderText="Deleted">
                                                            <HeaderStyle Width="45px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <%-- 7 --%>
                                                        <asp:BoundField DataField="vRcvPOScanDeletedUser" HeaderText="Deleted By">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <%-- 8 --%>
                                                        <asp:BoundField DataField="RcvPOScanDeletedNote" HeaderText="Delete Note">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <%-- 9 --%>
                                                        <asp:BoundField DataField="vRcvPOScanDeletedTime" HeaderText="Deleted Time">
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
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="6">
                                                <asp:HiddenField ID="HdfDataRowIdx" runat="server" />
                                                <asp:HiddenField ID="HdfProcess" runat="server" />
                                            </td>
                                        </tr>
                                    </table>
                                </asp:Panel>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfTransStatus" runat="server" Value="0" /></td>                            
                            <td><asp:HiddenField ID="HdfRcvType" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRcvPOType" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRcvPORefOID" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfStoKB" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfPOHOID" runat="server" /></td>
                            <td>&nbsp;</td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                &nbsp;</td>
                            <td>
                                &nbsp;</td>
                            <td>
                                &nbsp;</td>
                            <td></td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>
        
    </body>
</asp:Content>
<asp:Content ID="Content2" runat="server" contentplaceholderid="head">
    </asp:Content>

