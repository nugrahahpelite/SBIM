<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoPickList.aspx.vb" Inherits="SBSto.WbfSsoPickList" Title="SB WMS - Pick List" MasterPageFile="~/SBSto.Master" %>
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
            width: 300px;
        }
    </style>
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
                                <td style="font-size:15px;height:28px" colspan="3"><strong>PICK LIST</strong></td>                                
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
                                    <asp:Button ID="BtnCancelPCL" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal Pick List" Width="120px" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:125px">
                                    <asp:Button ID="BtnVoidPCL" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="VOID Pick List" Width="120px" />
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
                                                <td style="width:125px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="100px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressFind();" />
                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:70px;text-align:right">Company :</td>
                                                <td style="width:60px">
                                                    <asp:DropDownList ID="DstListCompany" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td  style="width:85px;text-align:right">No. Pick List : </td>
                                                <td style="width:120px">
                                                    <asp:TextBox ID="TxtListNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                </td>
                                                <td style="width:45px;text-align:right">No. Ref</td>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td style="width:100px;text-align:right">
                                                                <asp:TextBox ID="TxtListRefNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="165px"></asp:TextBox>
                                                            </td>
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
                                                <td>Trans. ID :<asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="10" Width="66px"></asp:TextBox>
                                                </td>
                                                <td style="text-align:right">Warehouse :</td>
                                                <td>
                                                    <asp:DropDownList ID="DstListWhs" runat="server" style="height: 20px" Width="250px">
                                                    </asp:DropDownList>
                                                </td>                                                
                                                <td colspan="9">
                                                    <table>
                                                        <tr>
                                                            <td style="text-align:right">&nbsp;</td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;<asp:CheckBox ID="ChkSt_Prepared" runat="server" ForeColor="#336600" Text="Prepared" Checked="True" />
                                                                <asp:CheckBox ID="ChkSt_OnPicking" runat="server" ForeColor="#336600" Text="On Picking" Checked="True" />
                                                                &nbsp;<asp:CheckBox ID="ChkSt_PickingDone" runat="server" ForeColor="#336600" Text="Picking Done" />
                                                                &nbsp;<asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
                                                                <asp:CheckBox ID="ChkSt_Void" runat="server" ForeColor="Red" Text="VOID" />
                                                            </td>
                                                            <td>
                                                                <asp:RadioButtonList ID="RdlListPickType" runat="server" RepeatDirection="Horizontal" ForeColor="Blue">
                                                                    <asp:ListItem Selected="True" Value="0">ALL</asp:ListItem>
                                                                    <asp:ListItem Value="5">TRB</asp:ListItem>
                                                                    <asp:ListItem Value="1">INVOICE</asp:ListItem>
                                                                    <asp:ListItem Value="8">DO Titip</asp:ListItem>
                                                                    <asp:ListItem Value="7">Perintah Kirim DO Titip</asp:ListItem>
                                                                </asp:RadioButtonList>
                                                            </td>
                                                            <td>
                                                                <asp:CheckBox ID="ChkCrPrioritas" runat="server" ForeColor="#336600" Text="PRIORITAS" />
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
                                                            <asp:ButtonField CommandName="Select" DataTextField="PCLNo" Text="Button" HeaderText="Nomor Pick List">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vPCLDate" HeaderText="Tanggal<br />Pick List" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vPCLScheduleDate" HeaderText="Schedule<br />Pick List" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SchDTypeName" HeaderText="Jenis" >
                                                                <HeaderStyle Width="75px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang Asal"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="DestTypeName" HeaderText="Destinasi">
                                                                <HeaderStyle Width="60px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLRefHOID" HeaderText="PCLRefHOID"> 
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLRefHNo" HeaderText="Ref No."> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vInvoicePrio" HeaderText="Prioritas" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLRefHInfo" HeaderText="Ref Info"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCKNo" HeaderText="No. Picking"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLNote" HeaderText="Note" >
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
                                                            <asp:BoundField DataField="PCLCancelNote" HeaderText="Cancel Note" >
                                                                <HeaderStyle Width="120px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PCLVoidNote" HeaderText="Void Note" >
                                                                <HeaderStyle Width="120px" />
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
                                <div id="DivListDoc" runat="server" >
                                    <asp:Panel ID="PanListDoc" class="myPanelGreyNS" runat="server" style="display:block;width:945px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListDocFind">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListDoc" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DAFTAR INVOICE</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListDocClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>No. Invoice</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtListDocNota" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                <asp:Label ID="LblMsgListDoc" runat="server" ForeColor="#FF0066"></asp:Label>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Customer</td>
                                                            <td>:</td>
                                                            <td>                                                                
                                                                <asp:TextBox ID="TxtListDocCustomer" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="200px"></asp:TextBox>                                                                
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnListDocFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" Font-Bold="True" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListDoc" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="no_nota" DataTextField="no_nota" Text="Button" HeaderText="No. Invoice">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vtanggal" HeaderText="Tanggal Invoice">
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="kode_cust" HeaderText="Kode Customer">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CUSTOMER" HeaderText="Customer">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="ALAMAT" HeaderText="Alamat">
                                                                <HeaderStyle Width="185px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="kota" HeaderText="Kota">
                                                                <HeaderStyle Width="185px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseOID" HeaderText="WarehouseOID">
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="tanggal" HeaderText="tanggal">
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
                                <div id="DivListTRB" runat="server" >
                                    <asp:Panel ID="PanListTRB" class="myPanelGreyNS" runat="server" style="display:block;width:945px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListTRBFind">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="Label2" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DAFTAR TRB</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListTRBClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>No.TRB</td>
                                                            <td>:</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListTRBNo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td></td>
                                                            <td>
                                                                <asp:Button ID="BtnListTRBFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" Font-Bold="True" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="13">
                                                                <asp:Label ID="LblMsgListTRB" runat="server" Font-Size="12px" ForeColor="#FF0066"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListTRB" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:ButtonField CommandName="NoBukti" DataTextField="NoBukti" HeaderText="No. TRB" Text="Button">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" HorizontalAlign="Center" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vTanggal" HeaderText="Tanggal TRB">
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="GudangAsal" HeaderText="Gudang Asal">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="GudangTujuan" HeaderText="Gudang Tujuan">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseAsalOID" HeaderText="WarehouseAsalOID">
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseTujuanOID" HeaderText="WarehouseTujuanOID">
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Tanggal" HeaderText="Tanggal">
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
                                <div id="DivListPKDOT" runat="server" >
                                    <asp:Panel ID="PanListPKDOT" class="myPanelGreyNS" runat="server" style="display:block;width:945px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White" DefaultButton="BtnListTRBFind">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="Label1" runat="server" Font-Size="17px" Font-Bold="True" ForeColor="#0066FF">DAFTAR PERINTAH KIRIM DO TITIP</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListPKDOTClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>No.Perintah Kirim</td>
                                                            <td>:</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListPKDOTNo" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td></td>
                                                            <td>
                                                                <asp:Button ID="BtnListPKDOTFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="TAMPILKAN DATA" Width="125px" Font-Bold="True" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="13">
                                                                <asp:Label ID="LblMsgListPKDOT" runat="server" Font-Size="12px" ForeColor="#FF0066"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListPKDOT" runat="server" AllowPaging="True" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <%-- 0 --%>
                                                            <asp:BoundField DataField="OID" HeaderText="ID Transaksi" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <%-- 1 --%>
                                                            <asp:BoundField DataField="PKDOTCompanyCode" HeaderText="Company">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <%-- 2 --%>
                                                            <asp:ButtonField CommandName="PKDOTNo" DataTextField="PKDOTNo" Text="Button" HeaderText="Nomor Perintah Kirim">
                                                                <HeaderStyle Width="125px" />
                                                                <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <%-- 3 --%>
                                                            <asp:BoundField DataField="vPKDOTScheduleDate" HeaderText="Schedule<br />Perintah Kirim" HtmlEncode="false" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <%-- 4 --%>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Gudang Asal"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <%-- 5 --%>
                                                            <asp:BoundField DataField="vWarehouseName_Dest" HeaderText="Gudang Tujuan">
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <%-- 6 --%>
                                                            <asp:BoundField DataField="NotaHOID" HeaderText="NotaHOID"> 
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <%-- 7 --%>
                                                            <asp:BoundField DataField="NotaNo" HeaderText="No. Invoice"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <%-- 8 --%>
                                                            <asp:BoundField DataField="vCustomer" HeaderText="Customer" HtmlEncode="false"> 
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <%-- 9 --%>
                                                            <asp:BoundField DataField="PKDOTScheduleDate" HeaderText="PKDOTDate"> 
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <%-- 10 --%>
                                                            <asp:BoundField DataField="WarehouseOID" HeaderText="WarehouseOID"> 
                                                                <HeaderStyle CssClass="myDisplayNone" />
                                                                <ItemStyle CssClass="myDisplayNone" />
                                                            </asp:BoundField>
                                                            <%-- 11 --%>
                                                            <asp:BoundField DataField="WarehouseOID_Dest" HeaderText="PKDOTDate"> 
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
                                <div id="DivPrintHS" runat="server" >
                                    <asp:Panel ID="PanPrintHS" class="myPanelGreyNS" runat="server" style="z-index:100;display:block;width:450px;height:350px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblRP" runat="server" Font-Size="17px" Font-Bold="True">Print History</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnRPClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <table>
                                                        <tr>
                                                            <td>Print History</td>
                                                        </tr>
                                                        <tr style="vertical-align:top">
                                                            <td>
                                                                <asp:GridView ID="GrvPrintHS" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="100%">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:BoundField DataField="PCLPrintNo" HeaderText="Print">
                                                                            <HeaderStyle Width="65px" />
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="PCLPrintBy" HeaderText="Print by">
                                                                            <HeaderStyle Width="145px" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="vPCLPrintDatetime" HeaderText="Print at">
                                                                            <HeaderStyle HorizontalAlign="Center" Width="145px" />
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
                            <td style="width:85px">Nomor Pick List</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPCLNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" ReadOnly="True" BackColor="#DDDDDD"></asp:TextBox>
                            </td>                            
                            <td style="width:25px"></td>
                            <td style="text-align:right;width:65px">Company</td>
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
                            <td></td>
                            <td>Schedule Pick</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPCLScheduleDate" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgPCLScheduleDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td></td>
                            <td style="text-align:right;">Gudang</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <table>
                                    <tr>
                                        <td>
                                            <asp:DropDownList ID="DstWhs" runat="server" style="height: 20px" Width="300px" AutoPostBack="True">
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
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" Width="145px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="BtnStatus" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td>No. Referensi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPCLRefNo" runat="server" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="11px" ReadOnly="True" Width="145px"></asp:TextBox>
                                <asp:TextBox ID="TxtPCLRefOID" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="11px" ReadOnly="True" Width="45px"></asp:TextBox>
                                <asp:Button ID="BtnPCLRefNo" runat="server" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Visible="False" Width="40px" />
                                <asp:Label ID="LblMsgPCLRefNo" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td></td>
                            <td style="text-align:right;">Tujuan</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <table>
                                    <tr>
                                        <td class="auto-style1">
                                            <asp:DropDownList ID="DstWhsDest" runat="server" style="height: 20px" Width="300px" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </td>                                        
                                        <td>
                                            <asp:Label ID="LblMsgWhsDest" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                            <td>Picking</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtPickingNo" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="190px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td colspan="5">
                                <table style="width:100%">
                                    <tr>
                                        <td>
                                            <asp:RadioButtonList ID="RdlPickType" runat="server" AutoPostBack="True" Enabled="False" ForeColor="#3333FF" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="5">TRB</asp:ListItem>
                                                <asp:ListItem Value="8">DO Titip</asp:ListItem>
                                                <asp:ListItem Value="7">Perintah Kirim DO Titip</asp:ListItem>
                                                <asp:ListItem Value="1">INVOICE</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td>
                                            <asp:CheckBox ID="ChkDest" runat="server" ForeColor="#006600" Text="LUAR KOTA" Visible="False" Font-Bold="True" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td style="text-align:center;vertical-align:top"></td>
                            <td rowspan="3" style="vertical-align:top">
                                <table>
                                    <tr>
                                        <td>
                                            <asp:TextBox ID="TxtPCLDescr" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="45px" MaxLength="245" TextMode="MultiLine" Width="400px"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                            <td style="vertical-align:top">Picking Status</td>
                            <td style="text-align:center;vertical-align:top">:</td>
                            <td style="vertical-align:top">
                                <asp:TextBox ID="TxtPickingStatus" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="190px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="6">
                                <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label>
                                <asp:Label ID="LblMsgPCL_Pick" runat="server" ForeColor="#FF0066"></asp:Label>
                                <asp:Label ID="LblPrioritas" runat="server" Font-Bold="True" Font-Size="25px" ForeColor="#0066FF">URGENT</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox ID="ChkPCL_Rack" runat="server" Checked="True" ForeColor="#336600" Text="Rack" CssClass="myDisplayNone" />
                                &nbsp;
                                <asp:CheckBox ID="ChkPCL_Floor" runat="server" Checked="True" ForeColor="#336600" Text="Floor" CssClass="myDisplayNone" />
                                &nbsp;
                                <asp:CheckBox ID="ChkPCL_CrossDock" runat="server" ForeColor="#336600" Text="Cross Dock" Checked="True" CssClass="myDisplayNone" />
                                &nbsp;
                                <asp:CheckBox ID="ChkPCL_DOTitip" runat="server" ForeColor="#336600" Text="Cross Dock" Checked="True" CssClass="myDisplayNone" />
                            </td>
                            <td>Print</td>
                            <td style="text-align:center;">:</td>
                            <td>
                                <asp:TextBox ID="TxtPCLPrint" runat="server" BackColor="#DDDDDD" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="35px"></asp:TextBox>
                                <asp:Button ID="BtnPrint" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                    </table>
                    <div style="width:177%; height:525px; overflow:auto; border:ridge">
                        <table>
                            <tr style="vertical-align:top">
                                <td style="width:10px"></td>
                                <td colspan="2">
                                    <asp:Panel ID="PanDetail" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:300px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                        <asp:GridView ID="GrvDetail" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <Columns>
                                                <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false">
                                                    <HeaderStyle Width="80px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang">
                                                    <HeaderStyle Width="245px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="RefQty" DataFormatString="{0:n0}" HeaderText="Qty Referensi">
                                                    <HeaderStyle Width="75px" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="PCLDQty" DataFormatString="{0:n0}" HeaderText="Qty Picklist">
                                                    <HeaderStyle Width="75px" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan">
                                                    <HeaderStyle Width="55px" />
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
                                    </asp:Panel>
                                    <br />

                                    <br />
                                    <asp:Label ID="LblInv" runat="server" Font-Bold="True" Font-Size="12px" ForeColor="#0066FF">DATA INVOICE</asp:Label>
                                    <asp:Panel ID="PanInv" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:450px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                        <asp:GridView ID="GrvInv" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <Columns>
                                                <asp:BoundField DataField="KodeBarang" HeaderText="Kode Barang" HtmlEncode="false">
                                                    <HeaderStyle Width="80px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="NamaBarang" HeaderText="Nama Barang">
                                                    <HeaderStyle Width="245px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Qty" DataFormatString="{0:n0}" HeaderText="Qty">
                                                    <HeaderStyle Width="75px" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="QtyBonus" DataFormatString="{0:n0}" HeaderText="Qty Bonus">
                                                    <HeaderStyle Width="75px" />
                                                    <ItemStyle HorizontalAlign="Right" />
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
                                    </asp:Panel>
                                    <br />

                                </td>
                                <td>
                                    <asp:Panel ID="PanRes" runat="server" style="height:525px">
                                        <table>
                                            <tr>
                                                <td class="auto-style2">
                                                    <asp:Label ID="LblMsgReserved" runat="server" Font-Size="12px" ForeColor="#0066FF"></asp:Label>
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Panel ID="PanReserved" class="myPanelGreyNS" runat="server" style="z-index:10;display:block;height:450px;width:98%" Visible="True" BorderStyle="Solid" BackColor="White" ScrollBars="Vertical">
                                                        <asp:GridView ID="GrvReserved" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <Columns>
                                                                <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang" HtmlEncode="false">
                                                                    <HeaderStyle Width="80px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang">
                                                                    <HeaderStyle Width="145px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="ReservedQty" DataFormatString="{0:n0}" HeaderText="Qty<br />Picklist" HtmlEncode="false">
                                                                    <HeaderStyle Width="55px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="QtyOnPicking" DataFormatString="{0:n0}" HeaderText="Qty<br />Picking" HtmlEncode="false">
                                                                    <HeaderStyle Width="55px" />
                                                                    <ItemStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="RcvPONo" HeaderText="No. Penerimaan" HtmlEncode="false">
                                                                    <HeaderStyle Width="80px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vRcvPODate" HeaderText="Tanggal<br />Penerimaan" HtmlEncode="false">
                                                                    <HeaderStyle Width="80px" />
                                                                    <ItemStyle Width="80px" HorizontalAlign="Center" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="vStorageInfoHtml" HeaderText="Storage" HtmlEncode="false">
                                                                    <HeaderStyle Width="150px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="StorageOID" HeaderText="Storage<br />OID" HtmlEncode="false">
                                                                    <HeaderStyle Width="55px" />
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="StorageStockOID" HeaderText="Storage<br />Stock<br />OID" HtmlEncode="false">
                                                                    <HeaderStyle Width="55px" />
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
                                                    </asp:Panel>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <table>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfTransStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfPickingStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfPickingHOID" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfProcess" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfRowIdx" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfActionStatus" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfDetailRowIdx" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfDetailOID" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfIsExpedition" runat="server" Value="0" /></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td><asp:HiddenField ID="HdfCompanyCode" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfWhs" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfWhsDest" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfOnList" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfPickTypeOID" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfPCLRefOID" runat="server" Value="1" /></td>
                            <td><asp:HiddenField ID="HdfPCLStorageTypeList" runat="server" Value="" /></td>
                            <td><asp:HiddenField ID="HdfEnableVoid" runat="server" Value="0" /></td>
                            <td><asp:HiddenField ID="HdfProcessDataKey" runat="server" /></td>
                            <td><asp:HiddenField ID="HdfPrioritas" runat="server" Value="" /></td>
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
