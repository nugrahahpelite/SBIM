<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfUserGroupMs.aspx.vb" Inherits="SBSto.WbfUserGroupMs" MasterPageFile="~/SBSto.Master" Title="SB WMS : Master User Group" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Master User Group</title>
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
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <table>
                        <tr>
                            <td style="vertical-align:top; width: 585px;">
                                <table style="font-family:tahoma;font-size:12px;">
                                    <tr>
                                        <td></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER USER GROUP</strong></td>
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
                                        <td class="auto-style2"></td>
                                        <td style="height: 24px">User Group Name</td>
                                        <td style="height: 24px">:</td>
                                        <td style="height: 24px">
                                            <table>
                                                <tr>
                                                    <td style="height: 20px"><asp:TextBox ID="TxtUGName" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" CssClass="setuppercase"></asp:TextBox></td>
                                                    <td style="height: 20px">
                                            <asp:Label ID="LblMsgUGName" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="auto-style2"></td>
                                        <td style="height: 24px">User Group Description</td>
                                        <td style="height: 24px">:</td>
                                        <td style="height: 24px">
                                            <table>
                                                <tr>
                                                    <td style="height: 20px"><asp:TextBox ID="TxtUGDescr" runat="server" Width="300px" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" CssClass="setuppercase" Height="60px" TextMode="MultiLine"></asp:TextBox></td>
                                                    <td style="height: 20px">
                                            <asp:Label ID="LblMsgUGDescr" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td><asp:Label ID="LblMsgErrorNE" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>
                                            <asp:HiddenField ID="HdfActionStatus" runat="server" />
                                        </td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" />
                                                        <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                                        <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                                        <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td>&nbsp;</td>
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
                            </td>
                            <td style="width:26px"></td>
                            <td style="vertical-align:top;height:100%">
                                <table style="width: 90%;margin-left:10px;font-family: tahoma;font-size:11px">
                                    <tr>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>User Group</td>
                                                    <td>:<br />
                                                    </td>
                                                    <td><asp:TextBox ID="TxtKriteria" runat="server" Width="245px" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox></td>                                                        
                                                    <td style="width:35px"></td>
                                                    <td><asp:Button ID="BtnFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" /></td>
                                                </tr>
                                            </table>
                                        </td>                
                                    </tr>                                        
                                    <tr>
                                        <td>
                                            <asp:GridView class="myPanelGreyLight" ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="11px" ShowHeaderWhenEmpty="True" CssClass="myGridViewHeaderFontWeightNormal" PageSize="7">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="OID" HeaderText="OID">
                                                        <HeaderStyle Width="65px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:ButtonField CommandName="Select" DataTextField="SsoUserGroupName" HeaderText="User Group Name">
                                                        <HeaderStyle Width="185px" />
                                                    </asp:ButtonField>
                                                    <asp:BoundField DataField="SsoUserGroupDescr" HeaderText="User Group Description">
                                                        <HeaderStyle Width="450px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Status" HeaderText="Status">
                                                        <HeaderStyle Width="85px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vCreationDatetime" HeaderText="Creation Date">
                                                        <HeaderStyle Width="175px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CreationUserName" HeaderText="Creation By">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="vModificationDatetime" HeaderText="Modification Date">
                                                        <HeaderStyle Width="175px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ModificationUserName" HeaderText="Modification By">
                                                        <HeaderStyle Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:CommandField ShowSelectButton="True" SelectText="Pilih" Visible="False" >
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
                                            <br />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td>
                                <asp:Panel ID="PanAccess" class="myPanelGreyNSa" runat="server" Width="100%" Height="500px" Visible="True" ScrollBars="Both">
                                    <table>
                                        <tr>
                                            <td></td>
                                            <td>
                                                <asp:GridView ID="GrvAccess" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <Columns>
                                                        <asp:BoundField DataField="TransCode" HeaderText="Trans Code">
                                                            <HeaderStyle Width="65px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="TransName" HeaderText="Trans Name">
                                                            <HeaderStyle Width="125px" />
                                                        </asp:BoundField>
                                                        <asp:TemplateField HeaderText="VIW&lt;br /&gt;View&lt;br /&gt;Data">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkVIW" runat="server" Text="" ToolTip="VIW" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="CED&lt;br /&gt;Create&lt;br /&gt;Edit&lt;br /&gt;Delete">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkCED" runat="server" Text="" ToolTip="CED" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="CNC&lt;br /&gt;Cancel&lt;br /&gt;Trans">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkCNC" runat="server" Text="" ToolTip="CNC" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="CLO&lt;br /&gt;Close&lt;br /&gt;Trans">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkCLO" runat="server" Text="" ToolTip="CLO" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="SCO&lt;br /&gt;Scan Open">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkSCO" runat="server" Text="" ToolTip="SCO" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="SCC&lt;br /&gt;Scan Close">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkSCC" runat="server" Text="" ToolTip="SCC" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="SCN&lt;br /&gt;Scan QR">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkSCN" runat="server" Text="" ToolTip="SCN" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="PRN&lt;br /&gt;Print">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkPRN" runat="server" Text="" ToolTip="PRN" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="PRP&lt;br /&gt;Prepare">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkPRP" runat="server" Text="" ToolTip="PRP" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="APP&lt;br /&gt;Approve">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkAPP" runat="server" Text="" ToolTip="APP" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="UPX&lt;br /&gt;Upload">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkUPX" runat="server" Text="" ToolTip="UPX" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="VOID&lt;br /&gt;Trans">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkVOI" runat="server" Text="" ToolTip="VOI" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="SGS&lt;br /&gt;Stagging In&lt;br /&gt;Start">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkSGS" runat="server" Text="" ToolTip="SGS" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="SGF&lt;br /&gt;Stagging In&lt;br /&gt;Finish">
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="ChkSGF" runat="server" Text="" ToolTip="SGF" />
                                                            </ItemTemplate>
                                                            <HeaderStyle Width="55px" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:TemplateField>
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
        .auto-style2 {
            height: 24px;
            width: 10px;
        }
    </style>
</asp:Content>
