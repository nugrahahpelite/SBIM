<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoRcvMsc.aspx.vb" Inherits="SBSto.WbfSsoRcvMsc" Title="SB WMS - Penerimaan Lain-lain" MasterPageFile="~/SBSto.Master" %>
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
            width: 65px;
        }
    </style>
    <script type="text/javascript">
        $(function () {
            $("#<%= TxtPenerimaanDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtPenerimaanDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
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
                                <td style="font-size:15px;height:28px" colspan="3"><strong>PENERIMAAN LAIN LAIN</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:420px">
                                    <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowProgressSave('Simpan');" />
                                    <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                    <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                    <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                    <asp:Label ID="LblProgressSave" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:125px">
                                    <asp:Button ID="BtnCancelPCL" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal" Width="120px" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnPrepare" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Prepare" Width="100px" Enabled="False" />
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
                                                <td style="width:250px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="100px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressFind();" />
                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:70px;text-align:right">Company</td>
                                                <td style="width:110px">
                                                    <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td  style="width:85px;text-align:right">No. </td>
                                                <td style="width:120px">
                                                    <asp:TextBox ID="TxtListNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                </td>
                                                <td></td>
                                                <td style="width:320px">
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
                                                <td>
                                                    <asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" ForeColor="Red" Height="30px" Text="CLOSE" Width="112px" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>Trans. ID :
                                                    <asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="80px"></asp:TextBox>
                                                </td>
                                                <td style="text-align:right">Warehouse </td>
                                                <td colspan="2">                                                    
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="3">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">Status :</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Prepared" runat="server" ForeColor="#336600" Text="Prepared" Checked="True" />
                                                                &nbsp;&nbsp;&nbsp; <asp:CheckBox ID="ChkSt_OnReceive" runat="server" ForeColor="#336600" Text="On Receive" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ReceiveDone" runat="server" ForeColor="#336600" Text="Receive Done" />
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
                                                            <asp:ButtonField CommandName="Select" DataTextField="RcvMscNo" Text="Button" HeaderText="Nomor Penerimaan">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vRcvMscDate" HeaderText="Tanggal Penerimaan" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvMscCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvMscNote" HeaderText="Note" >
                                                                <HeaderStyle Width="120px" />
                                                            </asp:BoundField>
                                                             <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" >
                                                                <HeaderStyle Width="120px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="TransStatusDescr" HeaderText="Status" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vCreation" HeaderText="Creation" >
                                                                <HeaderStyle Width="145px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vPrepared" HeaderText="Prepared" >
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
                                <div id="DivCheckBRG" runat="server" >
                                    <asp:Panel ID="PanCheckBRG" class="myPanelGreyNS" runat="server" style="z-index:84;display:block;width:825px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                        
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvCheckBRG" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%" HeaderStyle-CssClass="StickyHeader">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                       <Columns>
                                                            <asp:ButtonField CommandName="BRGCODE" DataTextField="BRGCODE" HeaderText="Kode Barang">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                                <HeaderStyle Width="245px" />
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
                                <div id="DivListBrg" runat="server" >
                                    <asp:Panel ID="PanListBrg" class="myPanelGreyNS" runat="server" style="z-index:84;display:block;width:825px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" >
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListBrg" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DAFTAR BARANG</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListBrgClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                       
                                                        <tr>
                                                            <td>No. Barang</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtListBrg" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnListBrgFind" runat="server" class="myButtonFinda" Font-Bold="True" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" />
                                                            </td>
                                                            <td></td>
                                                           
                                                        </tr>
                                                        <tr>
                                                            <td colspan="6">
                                                                <asp:Label ID="LblMsgListBrg" runat="server" ForeColor="#FF0066" Font-Size="12px"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListBrg" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%" HeaderStyle-CssClass="StickyHeader">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                       <Columns>
                                                            <asp:ButtonField CommandName="BRGCODE" DataTextField="BRGCODE" HeaderText="Kode Barang">
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
                            <td style="width:95px">No. Penerimaan</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPenerimaanNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#DDDDDD"></asp:TextBox>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="width:85px">Company</td>
                            <td style="width:20px;text-align:center">:</td>
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
                            <td style="width:25px"></td>
                            <td>ID Transaksi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransID" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="73px" ReadOnly="True"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Tanggal Penerimaan</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPenerimaanDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgPenerimaanDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td></td>
                            <td>Warehouse</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:DropDownList ID="DstWhs" runat="server" style="height: 20px" Width="250px">
                                            </asp:DropDownList>
                                        </td>
                                        <td>
                                            <asp:Label ID="LblMsgWhs" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                            <td>Status</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="190px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="BtnStatus" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>Note</td>
                            <td style="text-align:center">:</td>
                            <td rowspan="2" style="vertical-align:top">

                                <asp:TextBox ID="TxtPenerimaanNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="38px" MaxLength="245" TextMode="MultiLine" Width="300px"></asp:TextBox>

                            </td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>   
                            <td></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td rowspan="2" style="vertical-align:top">
                                <table>
                                    <tr>
                                        <td></td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                            <td></td>
                            <td style="text-align:center">&nbsp;</td>
                            <td></td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td></td>
                            <td></td>
                            <td colspan="9">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                                <asp:Label ID="LblMsgXlsProsesError" runat="server" ForeColor="#FF0066"></asp:Label>
                            </td>                            
                        </tr>
                    </table>
                    <div id="DivBrg" runat="server" style="width:177%; height:525px; overflow:auto; border:ridge">
                        <table>
                            <tr>
                                <td style="width:45px"></td>
                               
                                <td>Kode/Nama Barang
                                
                                    <asp:TextBox ID="TxtFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="190px"></asp:TextBox>
                                    <asp:Button ID="BtnFind" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" OnClientClick="fsShowFindProgress()" Text="Find" Width="55px" />
                                    <asp:Label ID="LblFindProgress" runat="server" ForeColor="#3333FF"></asp:Label>
                                </td>
                                 <td>
                                    <asp:TextBox ID="TxtBrgOID" runat="server" Font-Names="Tahoma" Enabled="false" Font-Size="12px" Visible="false" ></asp:TextBox>
                                    <asp:TextBox ID="TXTBrgRcvMscHOID" runat="server" Font-Names="Tahoma" Enabled="false" Font-Size="12px" Visible="false"></asp:TextBox>
                                     <asp:TextBox ID="TxtBrgCode" runat="server" Font-Names="Tahoma" Enabled="false" Font-Size="12px" Visible="false"></asp:TextBox>
                                     <asp:TextBox ID="TxtBrgStorageOID" runat="server" Font-Names="Tahoma" Enabled="false" Font-Size="12px" Visible="false"></asp:TextBox>
                                     <asp:TextBox ID="TextBrgCommand" runat="server" Font-Names="Tahoma" Enabled="false" Font-Size="12px" Visible="false"></asp:TextBox>
                                </td>
                                <%--<td>
                                    <asp:CheckBox ID="ChkFindVarian" runat="server" ForeColor="#336600" Text="Hanya yang Selisih" />
                                </td>
                                <td>
                                    <asp:CheckBox ID="ChkFindScan" runat="server" ForeColor="#336600" Text="Hanya yang Sudah Scan" />
                                </td>
                                <td>
                                    <asp:CheckBox ID="ChkFindIncludeDihapus" runat="server" ForeColor="Red" Text="Tampilkan Data Dihapus" />
                                </td>--%>
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td>
                                    <asp:GridView ID="GrvDetail" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <%-- 0 --%>
                                            <asp:BoundField DataField="OID" HeaderText="OID">
                                                <HeaderStyle CssClass="myDisplayNone" />
                                                <ItemStyle CssClass="myDisplayNone" />
                                            </asp:BoundField>
                                            <%-- 1 --%>
                                            <asp:ButtonField CommandName="vAddItem" DataTextField="vAddItem" Text="Button" HeaderText="">
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                            </asp:ButtonField>
                                            <%-- 2 --%>
                                            <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <%-- 3 --%>
                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                <HeaderStyle Width="300px" />
                                            </asp:BoundField>
                                            <%-- 4 --%>
                                            <asp:BoundField DataField="RcvMscQty" HeaderText="Qty" DataFormatString="{0:n0}" >
                                                <HeaderStyle Width="80px" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <%-- 5 --%>
                                            <asp:TemplateField HeaderText="Qty Penerimaan">
                                                <ItemTemplate>
                                                    <asp:TextBox runat="server" ID="TxtRcvMscQty" Width="100px" MaxLength="10"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <%-- 6 --%>
                                            <asp:ButtonField CommandName="vDelItem" DataTextField="vDelItem" Text="Button" HeaderText="">
                                                <HeaderStyle Width="115px" />
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
                            <td><asp:HiddenField ID="HdfWarehouseOID" runat="server" /></td>
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
                                <asp:HiddenField ID="HdfCompanyCode" runat="server" Value="2" />
                            </td>
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
