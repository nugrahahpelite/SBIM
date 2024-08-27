<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoReceiving.aspx.vb" Inherits="SBSto.WbfSsoReceiving" Title="SB WMS : Receiving" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>

    <script type="text/javascript">
        $(function () {
            $("#<%= TxtRcvDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtRcvRefDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });

            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtRcvDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtRcvRefDate.ClientID%>").datepicker({ dateFormat: 'dd M yy' });

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
                document.getElementById("<%= LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Cancel Request Return ?? WARNING : Cancel Tidak Dapat Dibatalkan";
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
                    document.getElementById("<%=LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Prepare Request Return ??";
                }
                else {
                    document.getElementById("<%= LblConfirmMessage.ClientID%>").innerText = "Anda yakin hendak Cancel Prepare Request Return ??";
                }
                document.getElementById("<%= HdfProcess.ClientID%>").value = "Prepare";
            }
        }
        function fsDisableYesConfirmSto() {
            document.getElementById("<%= BtnConfirmYes.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnConfirmNo.ClientID%>").style.display = "none";
            document.getElementById("<%= LblConfirmProgress.ClientID%>").innerText = document.getElementById("<%= HdfProcess.ClientID%>").value + " in Progress.........";
        }
        function fsShowInProgress(vriProcess) {
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProsesSimpan.ClientID%>").innerText = "Sedang Proses " + vriProcess + ".... Silakan Menunggu";
        }
        function fsShowFindProgress() {
            document.getElementById("<%= BtnFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblFindProgress.ClientID%>").innerText = "Sedang Proses...";
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
                                <td style="font-size:15px;height:28px" colspan="3"><strong>RECEIVING</strong></td>                                
                            </tr>
                            <tr>
                                <td style="width:10px"></td>
                                <td style="width:420px">
                                    <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowInProgress('Simpan')" />
                                    <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                    <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                    <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                    <asp:Label ID="LblProsesSimpan" runat="server" ForeColor="#0066FF"></asp:Label>
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnCancelRcv" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Batal Receiving" Width="100px" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnScanOpen" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Scan Open" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnScanClosed" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Scan Closed" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:10px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnCloseRcv" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Close Receiving" Width="100px" Enabled="False" />
                                </td>
                                <td style="width:25px"></td>
                                <td style="width:105px">
                                    <asp:Button ID="BtnPreview" runat="server" class="myButtonAct" Enabled="False" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Preview" Width="100px" />
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
                            <td style="width:10px">&nbsp;</td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="display:block;width:1650px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:85%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td style="width:150px">
                                                    <asp:Button ID="BtnListFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" />
                                                </td>
                                                <td style="width:70px;text-align:right">Trans. ID :</td>
                                                <td style="width:145px"><asp:TextBox ID="TxtListTransID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox></td>                                                
                                                <td  style="width:115px;text-align:right">Nomor Retur : </td>
                                                <td style="width:200px">
                                                    <asp:TextBox ID="TxtListRTNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="214px"></asp:TextBox>
                                                </td>
                                                <td style="width:100px"></td>
                                                <td style="width:100px;text-align:right">&nbsp;</td>
                                                <td style="width:270px;text-align:right">
                                                    <asp:Button ID="BtnListClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                                <td style="text-align:right">
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td></td>
                                                <td style="text-align:right">Status Receiving :</td>
                                                <td colspan="8">
                                                    <table>
                                                        <tr>
                                                            <td>
                                                                <asp:CheckBox ID="ChkSt_Baru" runat="server" ForeColor="#336600" Text="Baru" Checked="True" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_ScanOpen" runat="server" ForeColor="#336600" Text="Scan Open" />
                                                                &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ChkSt_ScanClosed" runat="server" ForeColor="#336600" Text="Scan Closed" />
                                                                &nbsp;
                                                                <asp:CheckBox ID="ChkSt_Closed" runat="server" ForeColor="#336600" Text="SUDAH CLOSE" />
                                                                &nbsp;&nbsp;
                                                                <asp:CheckBox ID="ChkSt_Cancelled" runat="server" ForeColor="Red" Text="BATAL" />
                                                            </td>
                                                            <td style="width:25px"></td>
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
                                                            <asp:ButtonField CommandName="Select" DataTextField="RcvNo" Text="Button" HeaderText="Nomor Retur">
                                                                <HeaderStyle Width="125px" />
                                                            <ItemStyle Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vRcvDate" HeaderText="Tanggal Retur" >
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvCompanyCode" HeaderText="Company" >
                                                                <HeaderStyle Width="65px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SubWhsName" HeaderText="Sub Warehouse" >
                                                                <HeaderStyle Width="70px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvNote" HeaderText="Note" >
                                                                <HeaderStyle Width="190px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvTypeName" HeaderText="Jenis Ref" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvRefNo" HeaderText="Nomor Ref" >
                                                                <HeaderStyle Width="100px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvCloseNote" HeaderText="Close Note" >
                                                                <HeaderStyle Width="190px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvCancelNote" HeaderText="Cancel Note" >
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
                                    <asp:Panel ID="PanLsScan" class="myPanelGreyNS" runat="server" style="display:block;width:950px;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
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
                                                    <asp:GridView ID="GrvLsScan" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" PageSize="20" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField DataField="RcvScanQty" HeaderText="Qty Receiving" DataFormatString="{0:n0}" >
                                                                <HeaderStyle Width="80px" />
                                                                <ItemStyle HorizontalAlign="Right" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvScanNote" HeaderText="Note" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vRcvScanUser" HeaderText="Receiving By" >
                                                                <HeaderStyle Width="115px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vRcvScanTime" HeaderText="Receiving Time" >
                                                                <HeaderStyle Width="195px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vRcvScanDeleted" HeaderText="Deleted" >
                                                                <HeaderStyle Width="45px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="RcvScanDeletedNote" HeaderText="Delete Note" >
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vRcvScanDeletedUser" HeaderText="Deleted By" >
                                                                <HeaderStyle Width="115px" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vRcvScanDeletedTime" HeaderText="Deleted Time" >
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
                                    <asp:Panel ID="PanConfirm" class="myPanelGreyNS" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="display:block;text-align:center;width:450px;left:20%">
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
                                                                <asp:Label ID="LblConfirmNote" runat="server" Text="Note : " Visible="False"></asp:Label>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtConfirmNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Visible="False" Width="270px"></asp:TextBox>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblConfirmWarning" runat="server" ForeColor="#FF0066" Font-Size="11pt"></asp:Label>                                                    
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
                                <div id="DivListItem" runat="server" >
                                    <asp:Panel ID="PanListItem" class="myPanelGreyNS" runat="server" style="display:block;width:650px;height:580px;margin-left:100px;margin-top:50px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                    <asp:Label ID="LblListItem" runat="server" Font-Size="17px" Font-Bold="True">DAFTAR BARANG</asp:Label>
                                                </td>
                                                <td style="text-align:right">
                                                    <asp:Button ID="BtnListItemClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="CLOSE" Width="112px" Font-Bold="True" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table>
                                                        <tr>
                                                            <td>Kode/Nama Barang</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListItem" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="145px"></asp:TextBox>
                                                            </td>
                                                            <td>
                                                                <asp:Button ID="BtnListItemFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Cari" Width="112px" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:GridView ID="GrvListItem" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="15" ShowHeaderWhenEmpty="True" Width="450px" AllowPaging="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:ButtonField CommandName="BRGCODE" DataTextField="BRGCODE" Text="Button" HeaderText="Kode Barang">
                                                                <HeaderStyle Width="100px" />
                                                                <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                                                <HeaderStyle Width="350px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" HtmlEncode="false">
                                                                <HeaderStyle Width="350px" />
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
                                    <asp:Panel ID="PanPrOption" runat="server" BackColor="#45D4D5" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:450px;left:20%">
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
                                                    <asp:RadioButtonList ID="RdbProXls" runat="server" RepeatDirection="Horizontal" Visible="False">
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
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:75%;margin-top:-45px">
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
                    <table>
                        <tr>
                            <td style="width:10px"></td>
                            <td style="width:115px">Nomor Penerimaan</td>
                            <td style="width:20px;text-align:center">:</td>
                            <td>
                                <asp:TextBox ID="TxtRcvNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="185px" BackColor="#E2E2E2" ReadOnly="True"></asp:TextBox>
                                <asp:Label ID="LblMsgRcvNo" runat="server" ForeColor="Red"></asp:Label>
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
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Tanggal Penerimaan</td>
                            <td style="text-align:center" class="auto-style1">:</td>
                            <td class="auto-style1">
                                <asp:TextBox ID="TxtRcvDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgRcvDate" runat="server" ForeColor="Red"></asp:Label>
                            </td>                            
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Sub Warehouse</td>
                            <td style="text-align:center" class="auto-style1">:</td>
                            <td class="auto-style1">
                                <asp:DropDownList ID="DstSubWhs" runat="server" style="height: 20px" Width="300px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgSubWhs" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                            <td class="auto-style1"></td>
                            <td class="auto-style1">Status</td>
                            <td style="text-align:center" class="auto-style2">:</td>
                            <td class="auto-style1">
                                <asp:TextBox ID="TxtTransStatus" runat="server" BackColor="#E2E2E2" Font-Names="Tahoma" Font-Size="12px" Width="190px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="BtnStatus" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                            </td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td>Note</td>
                            <td style="text-align:center">:</td>
                            <td rowspan="2">
                                <asp:TextBox ID="TxtRcvNote" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" Height="38px" MaxLength="245" TextMode="MultiLine" Width="345px"></asp:TextBox>
                            </td>
                            <td></td>
                            <td>Jenis Referensi</td>
                            <td style="text-align:center">:</td>
                            <td>
                                <asp:DropDownList ID="DstRcvRefType" runat="server" style="height: 20px" Width="195px">
                                </asp:DropDownList>
                                <asp:Label ID="LblMsgRcvRefType" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>                                                        
                            <td>No/Tanggal Referensi</td>
                            <td style="text-align:center">:</td>
                            <td>

                                <asp:TextBox ID="TxtRcvRefNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="190px"></asp:TextBox>
                                <asp:TextBox ID="TxtRcvRefDate" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="100px"></asp:TextBox>
                                <asp:Label ID="LblMsgRcvRefNo" runat="server" ForeColor="Red"></asp:Label>

                            </td>
                        </tr>
                        <tr style="vertical-align:top">
                            <td></td>
                            <td colspan="11">
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
                                        <td>
                                            <asp:CheckBox ID="ChkFindVarian" runat="server" ForeColor="#336600" Text="Tampilkan hanya yang Selisih" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                        </tr>
                    </table>
                    <table runat="server" id="TblDOPrint" >
                        <tr>
                            <td style="width:10px"></td>
                            <td>
                                <asp:GridView ID="GrvDetail" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <%-- 0 --%>
                                        <asp:BoundField DataField="OID" HeaderText="OID">
                                            <HeaderStyle CssClass="myDisplayNone" />
                                            <ItemStyle CssClass="myDisplayNone" />
                                        </asp:BoundField>
                                        <%-- 1 --%>
                                        <asp:ButtonField CommandName="vAddItem" DataTextField="vAddItem" Text="Button" HeaderText="">
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" Font-Underline="True" ForeColor="#0033CC" />
                                        </asp:ButtonField>
                                        <%-- 2 --%>
                                        <asp:BoundField DataField="BRGCODE" HeaderText="Kode Barang">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <%-- 3 --%>
                                        <asp:BoundField DataField="BRGNAME" HeaderText="Nama Barang" HtmlEncode="false">
                                            <HeaderStyle Width="350px" />
                                        </asp:BoundField>
                                        <%-- 4 --%>
                                        <asp:BoundField DataField="BRGUNIT" HeaderText="Satuan" HtmlEncode="false">
                                            <HeaderStyle Width="60px" />
                                        </asp:BoundField>
                                        <%-- 5 --%>
                                        <asp:BoundField DataField="RcvDQty" HeaderText="Qty Receiving" DataFormatString="{0:n0}" >
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Right"/>
                                        </asp:BoundField>
                                        <%-- 6 --%>
                                        <asp:TemplateField HeaderText="Qty Receiving">
                                            <ItemTemplate>
                                                <asp:TextBox runat="server" ID="TxtRcvDQty" Width="100px" MaxLength="10"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <%-- 7 --%>
                                        <asp:ButtonField CommandName="vSumRcvScanQty" DataTextField="vSumRcvScanQty" Text="Button" HeaderText="Total<br />Qty Scan" datatextformatstring="{0:n0}" >
                                            <HeaderStyle Width="85px" HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Right" Font-Underline="True" ForeColor="#0033CC" />
                                        </asp:ButtonField>
                                        <%-- 8 --%>
                                        <asp:BoundField DataField="vSumRcvScanVarian" HeaderText="Selisih" DataFormatString="{0:n0}" >
                                            <HeaderStyle Width="100px" />
                                            <ItemStyle HorizontalAlign="Right"/>
                                        </asp:BoundField>
                                        <%-- 9 --%>
                                        <asp:BoundField DataField="vRcvDNote" HeaderText="Note" HtmlEncode="false" >
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <%-- 10 --%>
                                        <asp:TemplateField HeaderText="Note">
                                            <ItemTemplate>
                                                <asp:TextBox runat="server" ID="TxtvRcvDNote" Width="245px" MaxLength="450" TextMode="MultiLine"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <%-- 11 --%>
                                        <asp:BoundField DataField="vRcvDNoteBy" HeaderText="Edit Note By" HtmlEncode="false">
                                            <HeaderStyle Width="180px" />
                                        </asp:BoundField>
                                        <%-- 12 --%>
                                        <asp:BoundField DataField="vRcvDNoteDatetime" HeaderText="Edit Note at" HtmlEncode="false">
                                            <HeaderStyle Width="100px" />
                                        </asp:BoundField>
                                        <%-- 13 --%>
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
                                <asp:HiddenField ID="HdfStatusBaru" runat="server" Value="0" />
                            </td>
                            <td>
                                <asp:HiddenField ID="HdfStatusPrepare" runat="server" Value="2" />
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
    <style type="text/css">
        .auto-style1 {
            height: 28px;
        }
        .auto-style2 {
            width: 20px;
            height: 28px;
        }
    </style>
</asp:Content>


