<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoSmTRB.aspx.vb" Inherits="SBSto.WbfSsoSmTRB" Title="SB WMS - Summary Barang untuk TRB" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
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
            height: 32px;
        }
        .auto-style2 {
            width: 20px;
            height: 32px;
        }
    </style>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtSmDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtSOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtSOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtSmDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtSOStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtSOEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowConfirmCancel() {
            if ('<%=Session("UserOID")%>' == "") {
                window.location = "Default.aspx";
                return;
            }
            else {
                if (document.getElementById("<%= TxtTransID.ClientID%>").value == "") {
                    return;
                }
                document.getElementById("<%= DivConfirm.ClientID%>").style.visibility = "visible";
                document.getElementById("<%= LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Cancel Surat Jalan ?? WARNING : Cancel Tidak Dapat Dibatalkan";
                document.getElementById("<%= HdfProcess.ClientID%>").value = "Cancel";
            }
        }        
        function fsShowConfirmPrepare() {
            if ('<%=Session("UserOID")%>' == "") {
                window.location = "Default.aspx";
                return;
            }
            else {
                if (document.getElementById("<%= TxtTransID.ClientID%>").value == "") {
                    return;
                }
                document.getElementById("<%= DivConfirm.ClientID%>").style.visibility = "visible";
                if (document.getElementById("<%= HdfTransStatus.ClientID%>").value == document.getElementById("<%= HdfStatusBaru.ClientID%>").value) {
                    document.getElementById("<%=LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Prepare Surat Jalan ??";
                }
                else {
                    document.getElementById("<%= LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Cancel Prepare Surat Jalan ??";
                }
                document.getElementById("<%= HdfProcess.ClientID%>").value = "Prepare";
            }
        }
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = document.getElementById("<%= HdfProcess.ClientID%>").value + " in Progress.........";
        }
        function fsShowProgressSave(vriProses) {
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressSave.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
        }
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
        }
        function fsShowProgressRefSum() {
            document.getElementById("<%= BtnRefreshSum.ClientID%>").style.display = "none";
            document.getElementById("<%= LblRefreshSum.ClientID%>").innerText = "Proses Refresh...";
        }
    </script>        
</head>
    <body>
        
        <div style="width:100%">

        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnPreview" />
                <asp:PostBackTrigger ControlID="BtnStatus" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>SUMMARY BARANG UNTUK TRB LIST</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:420px">
                                    <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowProgressSave('Simpan');" />
                                    <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                    <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" Visible="False" />
                                    <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                    <asp:Label ID="LblProgressSave" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:125px">
                                    <asp:Button ID="BtnCancelSm" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal Summary" Width="120px" Visible="False" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnSudahTRB" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Close - Sudah TRB" Width="145px" Enabled="False" Visible="False" />
                                </td>
                                <td style="width:105px">
                                    &nbsp;</td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnPreview" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Preview" Width="100px" Visible="False" />
                                </td>
                                <td style="width:45px"></td>
                                <td style="width:130px">
                                    <asp:Button ID="BtnList" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="List Transaksi" Width="125px" />
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="z-index:110;display:block;width:1650px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListFind" >
                                        <table style="width:85%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td style="width:150px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="100px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressFind();" />
                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:70px;text-align:right">Trans. ID :</td>
                                                <td style="width:110px"><asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="80px"></asp:TextBox></td>                                                
                                                <td  style="width:85px;text-align:right">No. Sales Order : </td>
                                                <td style="width:120px">
                                                    <asp:TextBox ID="TxtListSalesOrderNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                </td>
                                                <td style="width:15px;"></td>
                                                <td style="width:250px;">
                                                    <table>
                                                        <tr>                                                            
                                                            <td>Periode </td>
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
                                                <td style="width:150px;text-align:right">
                                                    <asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />
                                                </td>
                                                <td></td>
                                            </tr>
                                            <tr>
                                                <td></td>
                                                <td style="text-align:right">Warehouse Asal :</td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhsFrom" runat="server" style="height: 20px" Width="250px"></asp:DropDownList>
                                                </td>                                                
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td></td>
                                                            <td>Warehouse Tujuan :</td>
                                                            <td>
                                                                <asp:DropDownList ID="DstListWhsTo" runat="server" style="height: 20px" Width="250px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td style="text-align:right">Status Summary :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Closed" runat="server" ForeColor="#336600" Text="Close - Sudah TRB" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
&nbsp;
                                                                </td>
                                                            <td style="width:25px"></td>
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
                                                            <asp:ButtonField CommandName="Select" DataTextField="OID" Text="Button" HeaderText="ID Transaksi">
                                                                <HeaderStyle Width="55px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vSmTRBDate" HeaderText="Tanggal Summary" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SmTRBCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSubWarehouseFrom" HeaderText="Sub Warehouse Asal" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSubWarehouseTo" HeaderText="Sub Warehouse Tujuan" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SmTRBNote" HeaderText="Note" >
                                                                <HeaderStyle Width="120px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                                <HeaderStyle Width="165px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vClosed" HeaderText="Closed - Sudah TRB" >
                                                                <HeaderStyle Width="165px" />
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
                                <div id="DivListSO" runat="server" >
                                    <asp:Panel ID="PanListSO" class="myPanelGreyNS" runat="server" style="z-index:84;display:block;width:1150px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListSOFind" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListSO" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DAFTAR SALES ORDER</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListSOClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>No. SO</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtListSONo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                            <td>Customer</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtListSOCust" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                &nbsp;</td>
                                                            <td style="width:25px"></td>
                                                            <td>
                                                                &nbsp;</td>
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
                                                                <asp:Button ID="BtnListSOFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                                            </td>
                                                            <td>
                                                                &nbsp;</td>
                                                            <td>
                                                                <asp:Button ID="BtnListSOSelect" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="PILIH DATA" Width="125px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="6">
                                                                <asp:Label ID="LblMsgListPO" runat="server" ForeColor="#FF0066" Font-Size="12px"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListSO" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%" HeaderStyle-CssClass="StickyHeader">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <%-- 0 --%>
                                                            <asp:TemplateField>
                                                                <ItemTemplate>
                                                                    <asp:CheckBox ID="ChkSelect" Text="" runat="server" />
                                                                </ItemTemplate>
                                                                <HeaderStyle Width="35px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:TemplateField>
                                                            <%-- 1 --%>
                                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                                <HeaderStyle Width="70px" CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <%-- 2 --%>
                                                            <asp:BoundField DataField="SalesOrderNo" HeaderText="No. SO">
                                                                <HeaderStyle Width="115px" />
                                                            </asp:BoundField>
                                                            <%-- 3 --%>
                                                            <asp:BoundField DataField="vSalesOrderDate" HeaderText="Tanggal SO">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <%-- 4 --%>
                                                            <asp:BoundField DataField="vSUB" HeaderText="Kode Customer" HtmlEncode="false">
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <%-- 5 --%>
                                                            <asp:BoundField DataField="NAMA_CUSTOMER" HeaderText="Nama Customer" HtmlEncode="false">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <%-- 6 --%>
                                                            <asp:BoundField DataField="GDGOJL" HeaderText="Gudang" HtmlEncode="false">
                                                                <HeaderStyle Width="85px" />
                                                            </asp:BoundField>
                                                            <%-- 7 --%>
                                                            <asp:BoundField DataField="BRG" HeaderText="Kode Barang" HtmlEncode="false">
                                                                <HeaderStyle Width="80px" />
                                                            </asp:BoundField>
                                                            <%-- 8 --%>
                                                            <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang" HtmlEncode="false">
                                                                <HeaderStyle Width="245px" />
                                                            </asp:BoundField>
                                                            <%-- 9 --%>
                                                            <asp:BoundField DataField="QTY" HeaderText="Qty" DataFormatString="{0:n0}" >
                                                                <HeaderStyle Width="75px" />
                                                                <ItemStyle HorizontalAlign="Right" />
                                                            </asp:BoundField>
                                                            <%-- 10 --%>
                                                            <asp:BoundField DataField="vSalesOrderDOID" HeaderText="Sales Order Detail<br />OID" HtmlEncode="false">
                                                                <HeaderStyle Width="75px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <%-- 11 --%>
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
                                </div>
                                <div id="DivPreview" runat="server" >
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="z-index:104;display:block;width:75%;margin-top:-45px">
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
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td style="width:100px">Tanggal Summary</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtSmDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgSmDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:125px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:DropDownList ID="DstCompany" runat="server" style="height: 20px" Width="300px" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgCompany" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td style="width:25px"></td>
                            <td>ID Transaksi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="73px" ReadOnly="True"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Note</td>
                            <td style="text-align:center" class="auto-style1">:</td>
                            <td rowspan="2" style="vertical-align:top">
                                <asp:TextBox ID="TxtSmNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="38px" MaxLength="245" TextMode="MultiLine" Width="240px"></asp:TextBox>
                            </td>                            
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Sub Warehouse Asal</td>
                            <td style="text-align:center" class="auto-style1">:</td>
                            <td class="auto-style1">
                                <table>
                                    <tr>
                                        <td>
                                            <asp:DropDownList ID="DstSubWhsFrom" runat="server" style="height: 20px" Width="300px">
                                            </asp:DropDownList>
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgSubWhsFrom" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Status</td>
                            <td style="text-align:center" class="auto-style2">:</td>
                            <td class="auto-style1">
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="190px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="BtnStatus" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td></td>
                            <td>Sub Warehouse Tujuan</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:DropDownList ID="DstSubWhsTo" runat="server" style="height: 20px" Width="300px">
                                            </asp:DropDownList>
                                        </td>                                        
                                        <td>
                                            <asp:Label ID="LblMsgSubWhsTo" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td>
                                &nbsp;</td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="6" style="vertical-align:central">
                                <table>
                                    <tr>
                                        <td>
                                            <asp:RadioButtonList ID="RdbDS" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" ForeColor="#0066FF">
                                                <asp:ListItem Value="D">DETAIL SALES ORDER</asp:ListItem>
                                                <asp:ListItem Value="S" Selected="True">SUMMARY</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkSummNonZero" runat="server" Checked="True" ForeColor="#336600" Text="QTY REQUEST TRB &gt; 0" AutoPostBack="True" />
                                        </td>
                                        <td></td>
                                        <td>
                                            <asp:Button ID="BtnRefreshSum" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="REFRESH SUMMARY" Width="125px" OnClientClick="fsShowProgressRefSum();" />
                                            <asp:Label ID="LblRefreshSum" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td colspan="6" style="vertical-align:central">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td>                            
                        </tr>
                    </table>
                    <div style="width:95%; height:525px; overflow:auto; border:ridge">
                        <table>
                            <tr>
                                <td style="width:10px"></td>
                                <td>
                                    <asp:GridView ID="GrvSum" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false">
                                                <HeaderStyle Width="80px" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="245px" />
                                            </asp:BoundField>
                                            <%-- 2 --%>
                                            <asp:BoundField DataField="vQty_SO" HeaderText="Sub Wh Tujuan<br />Total Qty<br />Sales Order" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 3 --%>
                                            <asp:BoundField DataField="vQty_PCL" HeaderText="Sub Wh Tujuan<br />Total Qty<br />Picklist Gantung" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 4 --%>
                                            <asp:BoundField DataField="vQty_Stock" HeaderText="Sub Wh Tujuan<br />Qty<br />Stock" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 5 --%>
                                            <asp:BoundField DataField="vQty_RequestTRB" HeaderText="Sub Wh Tujuan<br />Qty<br />Request TRB" DataFormatString="{0:n0}" HtmlEncode="false">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 6 --%>
                                            <asp:BoundField DataField="vQty_Avail_Wh_Dest" HeaderText="Sub Wh Asal<br />Qty<br />Available" DataFormatString="{0:n0}" HtmlEncode="false">
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
                                    <asp:GridView ID="GrvDetail_SO" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" HeaderStyle-CssClass="StickyHeader" ShowHeaderWhenEmpty="True" Visible="False">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="OID" HeaderText="OID">
                                                <HeaderStyle CssClass="myDisplayNone" Width="70px" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:ButtonField CommandName="vAddItem" DataTextField="vAddItem" HeaderText="" Text="Button">
                                                <HeaderStyle Width="115px" CssClass="myDisplayNone" />
                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" HorizontalAlign="Center" CssClass="myDisplayNone" />
                                            </asp:ButtonField>
                                            <%-- 2 --%>
                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                <HeaderStyle CssClass="myDisplayNone" Width="70px" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 3 --%>
                                            <asp:BoundField DataField="SalesOrderNo" HeaderText="No. SO">
                                               <HeaderStyle Width="115px" />
                                            </asp:BoundField>
                                            <%-- 4 --%>
                                            <asp:BoundField DataField="vSalesOrderDate" HeaderText="Tanggal SO">
                                                <HeaderStyle Width="70px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <%-- 5 --%>
                                                <asp:BoundField DataField="vSUB" HeaderText="Kode Customer" HtmlEncode="false">
                                                <HeaderStyle Width="70px" />
                                            </asp:BoundField>
                                            <%-- 6 --%>
                                            <asp:BoundField DataField="NAMA_CUSTOMER" HeaderText="Nama Customer" HtmlEncode="false">
                                                <HeaderStyle Width="145px" />
                                            </asp:BoundField>
                                            <%-- 7 --%>
                                            <asp:BoundField DataField="GDGOJL" HeaderText="Gudang" HtmlEncode="false">
                                                <HeaderStyle Width="85px" />
                                            </asp:BoundField>
                                            <%-- 8 --%>
                                            <asp:BoundField DataField="BRG" HeaderText="Kode Barang" HtmlEncode="false">
                                                <HeaderStyle Width="80px" />
                                            </asp:BoundField>
                                            <%-- 9 --%>
                                            <asp:BoundField DataField="NAMA_BARANG" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="245px" />
                                            </asp:BoundField>
                                            <%-- 10 --%>
                                            <asp:BoundField DataField="QTY" DataFormatString="{0:n0}" HeaderText="Qty">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 11 --%>
                                            <asp:BoundField DataField="SourceDOID" HeaderText="Sales Order Detail&lt;br /&gt;OID" HtmlEncode="false">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <%-- 12 --%>
                                            <asp:BoundField DataField="SalesOrderHOID" HeaderText="Sales Order Header&lt;br /&gt;OID" HtmlEncode="false">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <%-- 13 --%>
                                            <asp:ButtonField CommandName="vDelItem" DataTextField="vDelItem" HeaderText="" Text="Button">
                                                <HeaderStyle Width="115px" CssClass="myDisplayNone" />
                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" HorizontalAlign="Center" CssClass="myDisplayNone" />
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
                        </table>
                    </div>
                    <table>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfTransStatus" runat="server" Value="0" /></td>                            
                            <td><asp:HiddenField ID="HdfProcess" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfActionStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfDetailOID" runat="server" /></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>
                                <asp:HiddenField ID="HdfStatusBaru" runat="server" Value="0" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfStatusPrepare" runat="server" Value="2" />
                            </td>
                            <td>
                                &nbsp;</td>
                            <td>
                                &nbsp;</td>
                            <td>
                                &nbsp;</td>
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
