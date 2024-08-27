<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoStorageMs.aspx.vb" Inherits="SBSto.WbfSsoStorageMs" MasterPageFile="~/SBSto.Master" Title="SB WMS : Master Storage Location" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Master Storage Location</title>
    <script type="text/javascript">
        function fsShowSimpanProgress() {
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= LblMsgSimpan.ClientID%>").innerText = "Sedang Proses...";
        }
        function fsGenQRProgress(){
            document.getElementById("<%= BtnGenQR.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnFind.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnPreview.ClientID%>").style.display = "none";

            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= LblGenQR.ClientID%>").innerText = "Sedang Proses...";
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
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER STORAGE LOCATION</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>ID</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td><asp:TextBox ID="TxtOID" runat="server" Width="87px" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True"></asp:TextBox></td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkActive" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Active" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Warehouse</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstWarehouse" runat="server" style="height: 20px" Width="225px" AutoPostBack="True">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgWarehouse" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Building</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstBuilding" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgBuilding" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Lantai</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstLantai" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                            <asp:Label ID="LblMsgLantai" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Zona</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstZona" runat="server" style="height: 20px" Width="225px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                            <asp:Label ID="LblMsgZona" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Storage Type</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:DropDownList ID="DstStorageType" runat="server" style="height: 20px" Width="225px" AutoPostBack="True">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgStorageType" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="2">
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsMultiLevel" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="MultiLevel" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsRack" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Rack" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsStagging" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Stagging" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsCrossDock" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="CrossDock" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsKarantina" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Karantina" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsDOTitip" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="DOTitip" Enabled="False" />
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="ChkIsDamage" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Damage" Enabled="False" />
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Sequence Number</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtSeqNo" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="85px" MaxLength="5"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgSeqNo" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Column</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtColumn" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="65px" MaxLength="3"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgColumn" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Level</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtLevel" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="65px" MaxLength="3"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgLevel" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Stagging</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:RadioButtonList ID="RdbStagging" runat="server" RepeatDirection="Horizontal">
                                                            <asp:ListItem Value="1">IN</asp:ListItem>
                                                            <asp:ListItem Value="2">OUT</asp:ListItem>
                                                        </asp:RadioButtonList>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgStagging" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Storage Number</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtStorageNo" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="85px" MaxLength="5"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgStorageNo" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>QRCodeID</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:TextBox ID="TxtQRCodeID" runat="server" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True" Width="245px" MaxLength="450"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LblMsgQRCodeID" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblMsgErrorNE" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowSimpanProgress();" />
                                                        <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                                        <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" Visible="False" />
                                                        <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                                        <asp:Label ID="LblMsgSimpan" runat="server" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:HiddenField ID="HdfActionStatus" runat="server" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>
                                        <td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>  
                                
                                <div id="DivPreview" runat="server" >
                                    <asp:Panel ID="PanPreview" runat="server" BackColor="#99ffcc" BorderColor="Gray" BorderStyle="Solid" style="display:block;width:75%;margin-top:-45px">
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
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <asp:Panel ID="PanList" class="myPanelGreyNSa" runat="server" Width="100%" style="height:500px" Visible="True">
                                    <table style="width: 90%;margin:auto;font-family: tahoma;font-size:11px">
                                        <tr>
                                            <td style="font-size:15px;height:28px" colspan="3"><strong>LIST STORAGE LOCATION</strong></td>
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
                                                                        <asp:DropDownList ID="DstListWarehouse" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td></td>
                                                                    <td>Building</td>
                                                                    <td>:</td>
                                                                    <td>
                                                                        <asp:DropDownList ID="DstListBuilding" runat="server" style="height: 20px" Width="225px">
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td>Lantai</td>
                                                                    <td>:</td>
                                                                    <td>
                                                                        <asp:DropDownList ID="DstListLantai" runat="server" style="height: 20px" Width="225px">
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td></td>
                                                                    <td>Zona</td>
                                                                    <td>:</td>
                                                                    <td>
                                                                        <asp:DropDownList ID="DstListZona" runat="server" style="height: 20px" Width="225px">
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td>Storage Type</td>
                                                                    <td>:</td>
                                                                    <td>
                                                                        <asp:DropDownList ID="DstListStorageType" runat="server" AutoPostBack="True" style="height: 20px" Width="225px">
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td></td>
                                                                    <td></td>
                                                                    <td></td>
                                                                    <td>
                                                                        <asp:CheckBox ID="ChkQRNull" runat="server" Font-Names="Tahoma" Font-Size="12px" ForeColor="#0066FF" Text="QR Null Only" />
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                        <td><asp:Button ID="BtnFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" />
                                                            <br />
                                                            <br />
                                                            <asp:Button ID="BtnPreview" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Preview" Width="112px" />
                                                        </td>
                                                        <td></td>
                                                        <td><asp:Button ID="BtnGen" class="myButtonFinda" runat="server" Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" style="visibility:hidden" />
                                                            <br />
                                                            <br />
                                                            <asp:Button ID="BtnGenQR" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="Generate QR" Width="112px" OnClientClick="fsGenQRProgress();" />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table>
                                                                <tr>
                                                                    <td style="width:75px">
                                                                        <asp:CheckBox ID="ChkListCheckAll" runat="server" AutoPostBack="True" Font-Names="Tahoma" Font-Size="12px" ForeColor="#0066FF" Text="Check All" />
                                                                    </td>
                                                                    <td></td>
                                                                    <td>
                                                                        <asp:Panel ID="PanListRackN" runat="server" Visible="false">
                                                                            <table>
                                                                                <tr>
                                                                                    <td>Storage Number</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="TxtListRackN_Start" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                    </td>
                                                                                    <td>s/d</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="TxtListRackN_End" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </asp:Panel>
                                                                        <asp:Panel ID="PanListRackY" runat="server" Visible="false">
                                                                            <table>
                                                                                <tr>
                                                                                    <td>SequenceNo</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="TxtListRackY_SeqNo" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                    </td>
                                                                                    <td></td>
                                                                                    <td>Level</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="TxtListRackY_Level" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </asp:Panel>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                        <td>
                                                            &nbsp;</td>
                                                        <td></td>
                                                        <td>
                                                            <asp:Label ID="LblGenQR" runat="server" ForeColor="#0066FF"></asp:Label>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="4">
                                                            <table>
                                                                <tr>
                                                                    <td>
                                                                        <asp:CheckBox ID="ChkListQRCodeID" runat="server" Font-Names="Tahoma" Font-Size="12px" ForeColor="#0066FF" Text="Find By QR Code ID" />
                                                                    </td>
                                                                    <td>
                                                                        <asp:TextBox ID="TxtListQRCodeID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="450" Width="345px"></asp:TextBox>
                                                                    </td>
                                                                    <td></td>
                                                                    <td>
                                                                        <asp:CheckBox ID="ChkListStorageOID" runat="server" Font-Names="Tahoma" Font-Size="12px" ForeColor="#0066FF" Text="Find By Storage OID" />
                                                                    </td>
                                                                    <td>
                                                                        <asp:TextBox ID="TxtListStorageOID" runat="server" Font-Names="Tahoma" Font-Size="12px" MaxLength="5" Width="85px"></asp:TextBox>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>                
                                        </tr>                                        
                                        <tr>
                                            <td>
                                                <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" Width="945px" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="25">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:TemplateField HeaderText="Print">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkPrint" Text="" runat="server" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="35px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="BuildingName" HeaderText="Building">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="LantaiDescription" HeaderText="Lantai">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="ZonaName" HeaderText="Zona">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:ButtonField CommandName="Select" DataTextField="StorageTypeName" HeaderText="Storage Type">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:ButtonField>
                                                        <asp:BoundField DataField="StorageSequenceNumber" HeaderText="Sequence<br />Number" HtmlEncode="false">
                                                            <HeaderStyle Width="75px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="StorageColumn" HeaderText="Column" HtmlEncode="false">
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="StorageLevel" HeaderText="Level" HtmlEncode="false">
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="StorageNumber" HeaderText="Storage<br />Number" HtmlEncode="false">
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="vStorageStagIO" HeaderText="Stagging" HtmlEncode="false">
                                                            <HeaderStyle Width="70px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="StorageQRCodeID" HeaderText="Storage<br />QR Code ID" HtmlEncode="false">
                                                            <HeaderStyle Width="100px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="OID" HeaderText="Storage<br />OID" HtmlEncode="false">
                                                            <HeaderStyle Width="65px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="Status" HeaderText="Status">
                                                            <HeaderStyle Width="85px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CreationDatetime" HeaderText="Creation Date">
                                                            <HeaderStyle Width="145px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="CreationUserName" HeaderText="Creation By">
                                                            <HeaderStyle Width="125px" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="ModificationDatetime" HeaderText="Modification Date">
                                                            <HeaderStyle Width="145px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="ModificationUserName" HeaderText="Modification By">
                                                            <HeaderStyle Width="125px" />
                                                        </asp:BoundField>
                                                        <asp:CommandField ShowSelectButton="True" SelectText="Pilih" Visible="False">
                                                            <HeaderStyle Width="54px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:CommandField>
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
                                </asp:Panel>
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
    <style type="text/css">
        .auto-style1 {
            width: 456px;
        }
    </style>
    </asp:Content>
