<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSOutToSOut.aspx.vb" Inherits="SBSto.WbfSsoSOutToSOut" Title="SB WMS - Moving Antar Staging Out" MasterPageFile="~/SBSto.Master" %>
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
            height: 24px;
        }
        .auto-style2 {
            width: 72px;
        }
        .auto-style3 {
            width: 999px;
        }
        .auto-style4 {
            width: 746px;
        }
    </style>
</head>
    <body>
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
                    
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td colspan="3">
                                    <asp:Label ID="LblTitle" runat="server" Font-Bold="True" Font-Size="19px" ForeColor="#333333" Text="MOVING ANTAR STAGING OUT"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td>
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
                                </td>
                                <td style="width:25px"></td>
                                <td>
                                    <asp:Button ID="BtnCancelSGO" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal Moving Antar Staging Out" Width="220px" Visible="False" />
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                                <td colspan="6">

                                    <asp:Label ID="LblXlsProses" runat="server" ForeColor="#0066FF"></asp:Label>
                                    <asp:Label ID="LblMsgXlsProsesError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                    <asp:Label ID="LblMsgConfirm" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                    <asp:Label ID="LblMsgError" runat="server" ForeColor="Red" Visible="False"></asp:Label>
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
                                                <td style="width:70px;text-align:right">Trans. ID :</td>
                                                <td class="auto-style1"><asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox></td>                                                
                                                <td style="width:115px;text-align:right">Nomor Moving : </td>
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
                                                    <%--<asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />--%>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td></td>
                                                <td style="text-align:right">Gudang :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td>Status</td>
                                                            <td>:</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Preparation" runat="server" Checked="True" ForeColor="#336600" Text="Staging Out 1 - Preparation" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_OnDelivery" runat="server" Checked="True" ForeColor="#336600" Text="On Delivery to Staging Out 2" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_OnReceiving" runat="server" Checked="True" ForeColor="#336600" Text="Staging Out 2 - On Receiving" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Done" runat="server" Checked="True" ForeColor="#336600" Text="Staging Out 2 - Receive Done" />
                                                                &nbsp;
                                                                <asp:CheckBox ID="ChkSt_Closed" runat="server" Checked="True" ForeColor="#336600" Text="Closed" />
                                                                &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
                                                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
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
                                                        <asp:BoundField DataField="SGOCompanyCode" HeaderText="Company" HtmlEncode="false">
                                                            <HeaderStyle Width="55px" />
                                                        </asp:BoundField>
                                                        <asp:ButtonField CommandName="SGONo" DataTextField="SGONo" HeaderText="No.Moving" Text="Button">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:ButtonField>
                                                        <asp:BoundField DataField="vSGODate" HeaderText="Tanggal<br />Moving" HtmlEncode="false">
                                                            <HeaderStyle Width="80px" />
                                                            <ItemStyle HorizontalAlign="Center"/>
                                                        </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vStgOut" HeaderText="Staging Out Asal">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                         <asp:BoundField DataField="vStgOut_InfoComplete" HeaderText="Staging Out Asal Tujuan Info">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vStgOut_Dest" HeaderText="Staging Out Tujuan">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                         <asp:BoundField DataField="vStgOut_Dest_InfoComplete" HeaderText="Staging Out Tujuan Info">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="SGOCancelNote" HeaderText="Cancel Note">
                                                            <HeaderStyle Width="100px" />
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
                    <table runat="server" id="tbTrans" class="auto-style3" >
                        <tr>
                            <td style="width:10px"></td>
                            <td style="width:85px">No. Moving</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#E2E2E2"></asp:TextBox>
                                
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:100px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td class="auto-style2">
                                <asp:TextBox ID="TxtCompany" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td style="width:25px"></td>
                            <td>ID Transaksi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="73px" ReadOnly="True"></asp:TextBox>
       
                            </td>
                           
                            <td>
                                 <asp:TextBox ID="TxtListFind" runat="server" BorderColor="#999999" BorderWidth="1px" MaxLength="85" TabIndex="5" Width="115px" autocomplete="off" Visible="false"></asp:TextBox>
                                </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Tanggal</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px" Enabled="False"></asp:TextBox>
                                &nbsp;</td>                            
                            <td></td>
                            <td>Warehouse</td>
                            <td style="text-align:center">:</td>
                            <td class="auto-style2">
                                <asp:TextBox ID="TxtTransWhsName" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>Status</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="190px" ReadOnly="True"></asp:TextBox>
                                
                            </td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Storage OID</td>                         
                            <td style="text-align:center" class="auto-style1">:</td>
                            <td class="auto-style1">
                                <asp:TextBox ID="TxtStorageOID" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px"></asp:TextBox>
                                <br />
                            </td>
                            <td class="auto-style1"></td>
                            <td style="vertical-align:top">Stage Out Asal</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtvStgOut_InfoComplete" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="245px" Height="45px" TextMode="MultiLine"></asp:TextBox>
                            </td>
                            <td></td>
                            <td style="vertical-align:top">Stage Out Tujuan</td>
                            <td style="text-align:center">:</td>
                            <td class="auto-style2">
                                <asp:TextBox ID="TxtvStgOut_Dest_InfoComplete" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="245px" Height="45px" TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                    
                       
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="10">
                                
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
                                        <asp:BoundField DataField="vPCKHOID" HeaderText="ID" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle Width="45px" HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:ButtonField CommandName="PCKNo" DataTextField="PCKNo" HeaderText="No.Picking" Text="Button">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle ForeColor="#0066FF" />
                                        </asp:ButtonField>
                                        <asp:BoundField DataField="vPCKDate" HeaderText="Tanggal<br />Picking" HtmlEncode="false">
                                            <HeaderStyle Width="75px" CssClass="myDisplayNone" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PCLNo" HeaderText="No.Picklist">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PCLRefHNo" HeaderText="No.Referensi">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SchDTypeName" HeaderText="Type">
                                            <HeaderStyle Width="70px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vIsQtyConfirm" HeaderText="Confirm" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle Width="45px" HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="vIsQtyConfirm_Dest" HeaderText="Confirm<br />Dest" HtmlEncode="false">
                                            <HeaderStyle Width="45px" />
                                            <ItemStyle Width="45px" HorizontalAlign="Center" />
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
                                <asp:Panel runat="server" ID="PanBRG" BorderColor="Gray" BorderStyle="Solid" style="display:block;" BackColor="White" Visible="false" Height="632px" Width="752px">
                                     <table runat="server" id="TbBRG" class="auto-style4" >
                                        <tr>
                                            <td colspan="6" style="text-align:left"><asp:Label ID="LblDataTitle" runat="server" Text="Data Hasil Scan" Font-Bold="True"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td colspan="6" style="background-color:#808080;height:4px">
                                                <asp:Label ID="Label3" runat="server" Height="24px" Visible="False" Width="75%"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                
                                
                             <td>
                                <asp:TextBox ID="TxtPCKOID" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="16px" ReadOnly="True"></asp:TextBox>
       
                            </td>
                                
                                <td style="width:500px"></td>
                                <td style="width:200px"></td>
                                <td></td>
                            </tr>
                                        <tr>
                                            <td colspan="6">
                                                <asp:GridView ID="GrvBRG" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="20" ShowHeaderWhenEmpty="True">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                         <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="vSGOOID" HeaderText="vSGOOID">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:BoundField DataField="BrgCode" HeaderText="Kode Barang" HtmlEncode="false">
                                                <HeaderStyle Width="75px" />
                                            </asp:BoundField>
                                            <%-- 2 --%>
                                            <asp:BoundField DataField="BrgName" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
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
                                            <asp:BoundField DataField="vSumPCKScanQty" DataFormatString="{0:n0}" HeaderText="Qty Pick">
                                                <HeaderStyle Width="50px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 6 --%>
                                            <asp:BoundField DataField="SGOScanQty" DataFormatString="{0:n0}" HeaderText="Qty Move">
                                                <HeaderStyle Width="50px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 7 --%>
                                            <asp:BoundField DataField="SGOScanQty_Dest" DataFormatString="{0:n0}" HeaderText="Qty Receive">
                                                <HeaderStyle Width="50px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 8 --%>
                                            <asp:ButtonField CommandName="vConfirm" DataTextField="vConfirm" Text="Button" HeaderText="">
                                                <HeaderStyle Width="45px" />
                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <%-- 9 --%>
                                            <asp:ButtonField CommandName="vNotConfirm" DataTextField="vNotConfirm" Text="Button" HeaderText="">
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
                                        </tr>
                                        <tr>
                                            <td colspan="6">
                                                <asp:HiddenField ID="HdfDataRowIdx" runat="server" />
                                                <asp:HiddenField ID="HdfProcess" runat="server" />
                                                 <asp:HiddenField ID="HdfRefNo2" runat="server" />
                                                 <asp:HiddenField ID="HdfPCKNo2" runat="server" />
                                                 <asp:HiddenField ID="HdfSGONo2" runat="server" />
                                                  <asp:HiddenField ID="HdfSGOAOID" runat="server" />
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
                            <td><asp:HiddenField ID="HdfTransStatus2" runat="server" Value="0" /></td>  
                         
                            <td><asp:HiddenField ID="HdfTransOID2" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HiddenField6" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRcvPOHOID" runat="server" Value="0" /></td>
                             <td><asp:HiddenField ID="HdfPCKNo" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfPCKOID" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfStoKB" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HiddenField3" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HiddenField4" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfPCKHOID" runat="server" />
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HdfProcessDataKey" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfBrgCode" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfLsPickRowIdx" runat="server" />
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HdfCompanyCode" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfWarehouseOID" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfStorageOID" runat="server" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfStorageOID_Dest" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:HiddenField ID="HdfTransOID" runat="server" />
                            </td>
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
