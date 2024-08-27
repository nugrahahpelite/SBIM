<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoPutAwayDswDs.aspx.vb" Inherits="SBSto.WbfSsoPutAwayDswDs" Title="SB WMS - Penerimaan Dispatch Gudang Sama" MasterPageFile="~/SBSto.Master" %>
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
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
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
        .auto-style1 {
            width: 4px;
        }
        .auto-style2 {
            width: 268435120px;
        }
    </style>
</head>
    <body>
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnProOK" />
                <asp:PostBackTrigger ControlID="BtnStatus" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>PENERIMAAN DISPATCH GUDANG SAMA</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:1000px">
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
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
                                                <td  style="width:115px;text-align:right">No. Dispatch: </td>
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
                                                    &nbsp;</td>
                                                <td style="text-align:right">Warehouse :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">Status :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_OnDispatch" runat="server" ForeColor="#336600" Text="On Dispatch Receive" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_DispatchDone" runat="server" ForeColor="#336600" Text="Receive Confirm" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Closed" runat="server" Checked="True" ForeColor="#336600" Text="Closed" />
                                                                &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
                                                            <asp:BoundField DataField="OID" HeaderText="ID" HtmlEncode="false">
                                                                <HeaderStyle Width="55px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="DSWCompanyCode" HeaderText="Company" HtmlEncode="false">
                                                                <HeaderStyle Width="55px" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="DSWNo" DataTextField="DSWNo" HeaderText="No.Putaway" Text="Button">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vDSWDate" HeaderText="Tanggal<br />Putaway" HtmlEncode="false">
                                                                <HeaderStyle Width="80px" />
                                                                <ItemStyle HorizontalAlign="Center"/>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="DSRNo" HeaderText="No.Penerimaan" HtmlEncode="false">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="DcmDriverName" HeaderText="Driver">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="VehicleNo" HeaderText="Mobil">
                                                                <HeaderStyle Width="50px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                                <HeaderStyle Width="75px" />
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
                                <div id="DivPrOption" runat="server" style="width:100%">
                                    <asp:Panel ID="PanPrOption" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:80;display:block;width:450px;left:20%">
                                        <br />
                                        <br />
                                        <table style="width:75%;margin:auto">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="Label2" runat="server" Font-Size="17px">Pilih Report</asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="DstProReport" runat="server" style="height: 20px" Width="300px">
                                                    </asp:DropDownList>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:CheckBox ID="ChkProVarianOnly" runat="server" ForeColor="#336600" Text="Hanya yang Selisih" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:RadioButtonList ID="RdbProXls" runat="server" RepeatDirection="Horizontal">
                                                        <asp:ListItem Selected="True">Pdf</asp:ListItem>
                                                        <asp:ListItem>Xls</asp:ListItem>
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
                            <td style="width:100px">No. Penerimaan</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#E2E2E2"></asp:TextBox>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:100px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtCompany" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="300px"></asp:TextBox>
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
                                <asp:TextBox ID="TxtTransDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px" Enabled="False" BackColor="#E2E2E2"></asp:TextBox>
                                &nbsp;</td>                            
                            <td></td>
                            <td>Warehouse Asal</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtWhsAsal" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="300px"></asp:TextBox>
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
                            <td>No. Dispatch</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtDispatchNo" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>Warehouse </td>                         
                            <td style="text-align:center">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtWhs" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="300px"></asp:TextBox>
                            </td>                            
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td>Driver</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtDriver" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>Lokasi Penerimaan</td>
                            <td style="text-align:center">:</td>
                            <td rowspan="2">
                                <asp:TextBox ID="TxtStorageInfo" runat="server" autocomplete="off" BackColor="#E2E2E2" BorderColor="#999999" BorderWidth="1px" Font-Size="15px" Font-Strikeout="False" Height="45px" MaxLength="450" TabIndex="5" TextMode="MultiLine" Width="300px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Vehicle</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtVehicle" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td colspan="6">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td>  
                            <td></td>
                            <td colspan="6">
                                <asp:Label ID="Label1" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td> 
                            
                           
                        </tr>
                    </table>
                    <table>
                        <tr style="vertical-align:top">
                            <td style="width:10px"></td>
                            <td>
                                <asp:GridView ID="GrvLsPick" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                       <asp:BoundField DataField="vPCKHOID" HeaderText="ID" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle Width="45px" HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="PCKNo" DataTextField="PCKNo" HeaderText="No.Picking" Text="Button">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle ForeColor="#0066FF" />
                                        </asp:ButtonField>
                                        <asp:BoundField DataField="vPCKDate" HeaderText="Tanggal<br />Picking" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PCLRefHNo" HeaderText="No.Referensi">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SchDTypeName" HeaderText="Type">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vIsQtyConfirm" HeaderText="Confirm" HtmlEncode="false">
                                            <HeaderStyle Width="75px" />
                                            <ItemStyle Width="75px" HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PCLNo" HeaderText="No.Picklist">
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="vDelItem" DataTextField="vDelItem" HeaderText="" Text="Button">
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
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
                                <asp:Panel runat="server" ID="PanData" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:650px;height:850px" BackColor="White" Visible="false">
                                    <table style="margin-top:15px;margin-left:15px;margin-bottom:15px;width:95%">
                                        <tr>
                                            <td style="text-align:left" class="auto-style2"><asp:Label ID="LblDataTitle" runat="server" Text="Data Hasil Scan" Font-Bold="True"></asp:Label></td>
                                            <td style="text-align:right"> <asp:Button ID="BtnDataClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Font-Bold="True" ForeColor="Red" /></td>
                                         
                                        </tr>
                                        
                                        <tr>
                                            <td colspan="6" class="auto-style2">
                                                <asp:GridView ID="GrvData" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True" AllowPaging="True">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <%-- 0 --%>
                                                        <asp:BoundField DataField="vDSRSOID" HeaderText="vDSRSOID">
                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                            <ItemStyle CssClass="myDisplayNone" />
                                                        </asp:BoundField>
                                                        <%-- 1 --%>
                                                        <asp:BoundField DataField="BrgCode" HeaderText="Kode Barang" HtmlEncode="false">
                                                            <HeaderStyle Width="75px" />
                                                        </asp:BoundField>
                                                        <%-- 2 --%>
                                                        <asp:BoundField DataField="BrgName" HeaderText="Nama Barang" HtmlEncode="false">
                                                            <HeaderStyle Width="245px" />
                                                        </asp:BoundField>
                                                        <%-- 3 --%>
                                                        <asp:BoundField DataField="RcvPONo" HeaderText="No.Penerimaan" HtmlEncode="false">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <%-- 4 --%>
                                                        <asp:BoundField DataField="RcvPOHOID" HeaderText="RcvPOHOID" >
                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                            <ItemStyle CssClass="myDisplayNone" />
                                                        </asp:BoundField>
                                                        <%-- 5 --%>
                                                        <asp:BoundField DataField="DSPScanQty" DataFormatString="{0:n0}" HeaderText="Qty Dispatch">
                                                            <HeaderStyle Width="50px" />
                                                            <ItemStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <%-- 6 --%>
                                                        <asp:BoundField DataField="DSRScanQty" DataFormatString="{0:n0}" HeaderText="Qty Receive">
                                                            <HeaderStyle Width="50px" />
                                                            <ItemStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <%-- 7 --%>
                                                        <asp:ButtonField CommandName="vConfirm" DataTextField="vConfirm" Text="Button" HeaderText="">
                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                            <ItemStyle CssClass="myDisplayNone" />
                                                        </asp:ButtonField>
                                                        <%-- 8 --%>
                                                        <asp:ButtonField CommandName="vNotConfirm" DataTextField="vNotConfirm" Text="Button" HeaderText="">
                                                            <HeaderStyle CssClass="myDisplayNone" />
                                                            <ItemStyle CssClass="myDisplayNone" />
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
                                        </tr>
                                        <tr>
                                            <td colspan="6" class="auto-style2">
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
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td>&nbsp;</td>
                            <td><asp:HiddenField ID="HdfPCKHOID" runat="server" Value="0" /></td>
                            <td class="auto-style1">
                                <asp:HiddenField ID="HdfCompanyCode" runat="server" Value="0" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:HiddenField ID="HdfDSPHOID" runat="server" Value="0" />
                            </td>
                            <td class="auto-style1">
                                <asp:HiddenField ID="HdfStorageOID" runat="server" Value="0" />
                            </td>
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
