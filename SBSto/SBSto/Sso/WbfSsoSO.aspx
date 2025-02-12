﻿<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSO.aspx.vb" Inherits="SBSto.WbfSsoSO" Title="SB WMS - Stock Opname" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>

<head>
    <title></title>
    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtSODate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtSODate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowInProgress(vriProcess) {
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
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
        function fsShowFindProgress() {
            document.getElementById("<%= BtnFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblFindProgress.ClientID%>").innerText = "Sedang Proses...";
        }
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = "Sedang Proses...";
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
                <asp:PostBackTrigger ControlID="BtnSimpan" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>SB STOCK OPNAME</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:420px">
                                    <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowInProgress('Create SO')" />
                                    <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                    <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                    <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                </td>
                                <td style="width:25px">
                                    &nbsp;</td>
                                <td style="width:125px">
                                    <asp:Button ID="BtnCancelSO" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal SO" Width="120px" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnScanOpen" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Scan Open" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnScanClosed" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Scan Close" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnCloseSO" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Close SO" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnPreview" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Preview" Width="100px" />
                                </td>
                                <td style="width:45px"></td>
                                <td style="width:130px">
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
                                </td>
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
                                                <td style="text-align:right">Company :</td>
                                                <td style="width:200px">
                                                    <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td  style="width:115px;text-align:right">Nomor SO : </td>
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
                                                <td>Trans. ID :<asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                </td>
                                                <td style="text-align:right">Warehouse :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">Status SO :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ScanOpen" runat="server" ForeColor="#336600" Text="Scan Open" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ScanClosed" runat="server" ForeColor="#336600" Text="Scan Closed" />
                                                                &nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ChkSt_Closed" runat="server" ForeColor="#336600" Text="SUDAH CLOSE" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
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
                                                            <asp:ButtonField CommandName="Select" DataTextField="SONo" Text="Button" HeaderText="Nomor SO">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vSOCutOff" HeaderText="Cut Off SO" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SubWhsName" HeaderText="Sub Warehouse" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SONote" HeaderText="Note" >
                                                                <HeaderStyle Width="190px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOCloseNote" HeaderText="Close Note" >
                                                                <HeaderStyle Width="190px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOCancelNote" HeaderText="Cancel Note" >
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
                                                            <asp:BoundField DataField="vScanOpen" HeaderText="Scan Open" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vScanClosed" HeaderText="Scan Closed">
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vClosed" HeaderText="Closed">
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
                                <div id="DivLsScan" runat="server" >
                                    <asp:Panel ID="PanLsScan" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;width:950px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblLsScanTitle" runat="server" Font-Size="17px"></asp:Label>
                                                </td>                                                
                                                <td style="width:150px;text-align:right">
                                                    <asp:Button ID="BtnLsScanClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>
                                                                <asp:Button ID="BtnLsScanDataFind" runat="server" class="myButtonFind" Font-Names="Tahoma" Font-Size="12px" Height="22px" Text="Cari" UseSubmitBehavior="False" Width="100px" />
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtLsScanDataFind" runat="server" autocomplete="off" BorderColor="#999999" BorderWidth="1px" MaxLength="85" TabIndex="5" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkLsScanSt_DelNo" runat="server" Checked="True" ForeColor="#336600" Text="Data Masuk" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkLsScanSt_DelYes" runat="server" ForeColor="Red" Text="Data Dihapus" />
                                                            </td>
                                                            <td>
                                                                <asp:HiddenField ID="HdfLsScanBrgCode" runat="server" Value="0" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <table style="width: 100%;margin:auto;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvLsScan" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" PageSize="30" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Info" HtmlEncode="false">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOScanQty" HeaderText="Qty SO" DataFormatString="{0:n0}" >
                                                                <HeaderStyle Width="80px" />
                                                                <ItemStyle HorizontalAlign="Right" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOScanNote" HeaderText="Note" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSOScanUser" HeaderText="Scan By" >
                                                                <HeaderStyle Width="115px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSOScanTime" HeaderText="Scan Time" >
                                                                <HeaderStyle Width="195px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSOScanDeleted" HeaderText="Deleted" >
                                                                <HeaderStyle Width="45px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SOScanDeletedNote" HeaderText="Delete Note" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSOScanDeletedUser" HeaderText="Deleted By" >
                                                                <HeaderStyle Width="115px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSOScanDeletedTime" HeaderText="Deleted Time" >
                                                                <HeaderStyle Width="195px" />
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
                                    <asp:Panel ID="PanConfirm" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="z-index:90;display:block;text-align:center;width:450px;left:20%">
                                        <br />
                                        <asp:Label ID="LblConfirmMessage" runat="server" Text="Anda Yakin ?" Font-Size="17px"></asp:Label>
                                        <br />
                                        <asp:Label ID="LblConfirmProgress" runat="server" Font-Size="17px"></asp:Label>
                                        <br />
                                        <table style="width:75%;margin:auto">
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
                            <td style="width:85px">Nomor SO</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtSONo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#E2E2E2"></asp:TextBox>
                                <asp:Label ID="LblMsgSONo" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:100px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="300px" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgCompany" runat="server" ForeColor="Red"></asp:Label>
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
                            <td>Cut Off SO</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtSODate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px" Enabled="False"></asp:TextBox>
                                &nbsp;<asp:DropDownList ID="DstCutOffHour" runat="server" Width="50px">
                                </asp:DropDownList>
                                &nbsp;:
                                <asp:DropDownList ID="DstCutOffMin" runat="server" Width="50px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgSODate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td></td>
                            <td>Sub Warehouse</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstSubWhs" runat="server" style="height: 20px" Width="300px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgSubWhs" runat="server" ForeColor="Red"></asp:Label>
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
                            <td>Note</td>
                            <td style="text-align:center">:</td>
                            <td rowspan="2" style="vertical-align:top">
                                <asp:TextBox ID="TxtSONote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="38px" MaxLength="245" TextMode="MultiLine" Width="300px" ReadOnly="True"></asp:TextBox>

                            </td>
                            <td></td>
                            <td>File xls</td>                         
                            <td style="text-align:center">:</td>
                            <td colspan="5">
                                <asp:FileUpload ID="FupXls" runat="server" accept=".xls" onchange="setLblFupUpload()" style="color:transparent" Width="425px" Enabled="False" />
                                <br />
                                <asp:Label ID="LblMsgFupXls" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                <asp:Label ID="LblFupXls" runat="server" ForeColor="#3333FF"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td></td>
                            <td style="vertical-align:top">Worksheet</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtXlsWorksheet" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="185px">Sheet 1</asp:TextBox>
                                <asp:Label ID="LblXlsWorksheet" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                            </td>
                            <td style="vertical-align:top"></td>
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
                                            <asp:RadioButtonList ID="RdbDetailType" runat="server" AutoPostBack="True" Font-Bold="True" ForeColor="#0066FF" RepeatDirection="Horizontal">
                                                <asp:ListItem Selected="True" Value="Det">DETAIL</asp:ListItem>
                                                <asp:ListItem Value="Sum">SUMMARY</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td style="width:45px"></td>
                                        <td>
                                            <asp:Button ID="BtnFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="Find" Width="55px" OnClientClick="fsShowFindProgress()" />
                                            <asp:Label ID="LblFindProgress" runat="server" ForeColor="#3333FF"></asp:Label>
                                        </td>
                                        <td>Kode/Nama Barang</td>
                                        <td>
                                            <asp:TextBox ID="TxtFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="190px"></asp:TextBox>
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkFindVarian" runat="server" ForeColor="#336600" Text="Hanya yang Selisih" />
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkFindScan" runat="server" ForeColor="#336600" Text="Hanya yang Sudah Scan" />
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkFindIncludeDihapus" runat="server" ForeColor="Red" Text="Tampilkan Data Dihapus" />
                                        </td>
                                    </tr>
                                </table>
                            </td>                            
                            <td style="vertical-align:central">
                                <table>
                                    <tr>
                                        <td style="width:85px"></td>
                                        <td style="background-color:yellow">
                                            <asp:CheckBox ID="ChkFindNotActive" runat="server" ForeColor="Red" Text="ITEM TIDAK AKTIF" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <asp:Panel ID="PanDetail" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:580px" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                    <asp:GridView ID="GrvDetail" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" AllowPaging="true" PageSize="25" Visible="False" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="vNo" HeaderText="No.">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:BoundField DataField="OID" HeaderText="OID">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 2 --%>
                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <%-- 3 --%>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="250px" />
                                            </asp:BoundField>
                                            <%-- 4 --%>
                                            <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" HtmlEncode="false">
                                                <HeaderStyle Width="50px" />
                                            </asp:BoundField>
                                            <%-- 5 --%>
                                            <asp:BoundField DataField="SOStockQty" HeaderText="Qty Winacc" DataFormatString="{0:n0}" >
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 6 --%>
                                            <asp:ButtonField CommandName="vSumSOScanQty" DataTextField="vSumSOScanQty" Text="Button" HeaderText="Total<br />Qty SO" datatextformatstring="{0:n0}" >
                                                <HeaderStyle Width="85px" HorizontalAlign="Center" />
                                                <ItemStyle HorizontalAlign="Right" Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <%-- 7 --%>
                                            <asp:BoundField DataField="vSOStockScanVarian" HeaderText="Selisih" DataFormatString="{0:n0}" >
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 8 --%>
                                            <asp:BoundField DataField="vSOStockNote" HeaderText="Note" HtmlEncode="false">
                                                <HeaderStyle Width="180px" />
                                            </asp:BoundField>
                                            <%-- 9 --%>
                                            <asp:TemplateField HeaderText="Note">
                                                <ItemTemplate>
                                                    <asp:TextBox runat="server" ID="TxtvSOStockNote" Width="245px" MaxLength="450" TextMode="MultiLine"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <%-- 10 --%>
                                            <asp:BoundField DataField="vSOStockNoteBy" HeaderText="Edit Note By" HtmlEncode="false">
                                                <HeaderStyle Width="180px" />
                                            </asp:BoundField>
                                            <%-- 11 --%>
                                            <asp:BoundField DataField="vSOStockNoteDatetime" HeaderText="Edit Note at" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
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
                                    <asp:GridView ID="GrvTaDetail" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="30" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="vNo" HeaderText="No.">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="OID" HeaderText="OID">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="StorageOID" HeaderText="StorageOID">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="250px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" HtmlEncode="false">
                                                <HeaderStyle Width="50px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage Info" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SOScanQty" DataFormatString="{0:n0}" HeaderText="Qty SO">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SOScanNote" HeaderText="Note">
                                                <HeaderStyle Width="145px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vSOScanUser" HeaderText="Scan By">
                                                <HeaderStyle Width="115px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vSOScanTime" HeaderText="Scan Time">
                                                <HeaderStyle Width="195px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vSOScanDeleted" HeaderText="Deleted">
                                                <HeaderStyle Width="45px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="SOScanDeletedNote" HeaderText="Delete Note">
                                                <HeaderStyle Width="145px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vSOScanDeletedUser" HeaderText="Deleted By">
                                                <HeaderStyle Width="115px" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vSOScanDeletedTime" HeaderText="Deleted Time">
                                                <HeaderStyle Width="195px" />
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
                                </asp:Panel>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfTransStatus" runat="server" Value="0" /></td>                            
                            <td><asp:HiddenField ID="HdfProcess" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfActionStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfDetailOID" runat="server" /></td>
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

