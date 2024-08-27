<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfUserList.aspx.vb" Inherits="SBSto.WbfUserList" Title="SB WMS : List User" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head >
    <title></title>
    <link href="~/CssFiles/CssBnsrp.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .myGridViewHeaderFontWeightNormal {}
        .auto-style9 {
            width: 15px;
            height: 43px;
        }
        .auto-style10 {
            height: 43px;
        }
    </style>
    <script type="text/javascript">   
        function fsShowProgressFind() {
            document.getElementById("<%= BtnFind.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBackMs.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses...";
        }
    </script>
</head>
<body>
    <form id="form1">
    <asp:ScriptManager ID="ScmUserList" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpnUserMs" runat="server">
        <ContentTemplate>    
            <table style="width: 100%;margin:auto;font-family: tahoma;font-size:12px">
                <tr>
                    <td class="auto-style9"></td>
                    <td class="auto-style10">
                        <table style="font-family: tahoma;font-size:12px">
                            <tr>
                                <td style="width:85px">Nama User</td>
                                <td style="width:15px">:</td>
                                <td><asp:TextBox ID="TxtKriteria" runat="server" Width="159px" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox></td>
                                <td>
                                    <asp:Button ID="BtnFind" class="myButtonFind" runat="server" Text="  Cari   " Width="88px" Font-Names="Tahoma" Font-Size="12px" Height="30px" OnClientClick="fsShowProgressFind();"/>
                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                </td>
                                <td>&nbsp;&nbsp;&nbsp; &nbsp;</td>
                                <td><asp:Button ID="BtnBackMs" class="myButtonFind" runat="server" Text="Kembali ke Master" Width="168px" Font-Names="Tahoma" Font-Size="12px" BackColor="#CCCCCC" BorderStyle="None" Height="30px" /></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td>
                        <table style="font-family: tahoma;font-size:12px">
                            <tr>
                                <td style="width:85px">User Group</td>
                                <td style="width:15px">:</td>
                                <td>
                                    <asp:DropDownList ID="DstUserGroup" runat="server" style="height: 20px" Width="250px">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width:15px"></td>
                    <td>
                        <table style="font-family: tahoma;font-size:12px">
                            <tr><td style="width:85px"></td>
                                <td style="width:15px"></td>
                                <td class="width:245px"><asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Text="LblMsgError" Visible="False"></asp:Label></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width:15px"></td>
                    <td>
                        <asp:GridView class="myPanelGreyLight" ID="GrvUser" runat="server" AutoGenerateColumns="False" Width="1000px" CellPadding="4" ForeColor="#333333" AllowPaging="True" Font-Names="Tahoma" Font-Size="12px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" PageSize="25">
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:ButtonField CommandName="Select" DataTextField="UserName" HeaderText="User Name" >
                                    <HeaderStyle Width="185px" HorizontalAlign="Center" />
                                </asp:ButtonField>
                                <asp:BoundField DataField="UserID" HeaderText="User ID" >
                                    <HeaderStyle Width="45px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UserSSO" HeaderText="User SSO">
                                    <HeaderStyle Width="25px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UserNip" HeaderText="User NIP" >
                                    <HeaderStyle Width="45px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UserCompanyCode" HeaderText="Company Code" >
                                    <HeaderStyle Width="45px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UserWarehouseCode" HeaderText="Warehouse Code" >
                                    <HeaderStyle Width="45px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Admin" HeaderText="Administrator">
                                    <HeaderStyle Width="25px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="SsoUserGroupName" HeaderText="User Group" >
                                    <HeaderStyle Width="100px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Status" HeaderText="Status">
                                    <HeaderStyle Width="54px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="OID" HeaderText="OID" >
                                    <HeaderStyle Width="25px" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="ModificationDatetime" HeaderText="Modification Date" >
                                    <HeaderStyle Width="145px" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="ModificationUserName" HeaderText="Modification By" >
                                    <HeaderStyle Width="145px" />
                                </asp:BoundField>
                            </Columns>
                            <EditRowStyle BackColor="#999999" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="25px" />
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
        </ContentTemplate>
    </asp:UpdatePanel>
    <asp:HiddenField ID="HdfSelectUserOID" runat="server" />
    </form>
</body>
</html>

</asp:Content>