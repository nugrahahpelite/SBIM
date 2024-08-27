<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="SBSto._Default" Title="SB WMS : Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>WMS</title>    
        <link href="~/CssFiles/CssBnsrp.css" rel="stylesheet" type="text/css" />  
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScmLogin" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnLogin" runat="server">
        <ContentTemplate>
            <table class="myPanelGreyLight" style="position:absolute;top:25%;left:25%;background-color:#F2F2F2;font-family:Tahoma;font-size:12px">
                <tr><td colspan="4" style="font-size:15px">
                    <table>
                        <tr>
                            <td rowspan="2"><img id="Img1" src="~/Images/SBLogo.png" runat="server" style="height: 55px; width: 91%" /></td>                    
                            <td class="auto-style3"><p class="site-title" style="font-variant: normal; font-style: italic ; font-weight: 500; font-size: xx-large ; text-transform: capitalize;">SUMBER BERKAT</p>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style4">SB WMS</td>
                        </tr>
                    </table>
                    </td>
                </tr>
                <tr><td style="height: 25px;">User ID</td>
                    <td>:&nbsp;&nbsp;&nbsp; </td>
                    <td style="width:113px" class="auto-style2">
                        <asp:TextBox ID="TxtUserID" runat="server" Font-Names="Tahoma" Font-Size="12px"></asp:TextBox>
                    </td>
                    <td class="auto-style1" style="width: 260px">                        
                    </td>
                </tr>
                <tr><td style="height: 25px;">Password&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                    <td>:</td>
                    <td style="width: 113px" class="auto-style2">
                        <asp:TextBox ID="TxtPassword" runat="server" TextMode="Password" Font-Names="Tahoma" Font-Size="12px">bnsrps</asp:TextBox>                
                    </td>
                    <td class="auto-style1" style="width: 260px"><asp:Label ID="LblMessage" runat="server" ForeColor="Red" Text="Message" Visible="False"></asp:Label></td>
                </tr>
                <tr>
                    <td style="height: 23px"></td>
                    <td style="height: 23px"></td>
                    <td style="height: 23px" colspan="2"><asp:Label ID="LblMessageSessionEnd" runat="server" ForeColor="Red" Text="SESSION ANDA SUDAH BERAKHIR. HARAP LOGIN ULANG" Visible="False"></asp:Label></td>
                </tr>
                <tr><td>
                    &nbsp;</td>
                    <td>&nbsp;</td>
                    <td style="width: 113px" class="auto-style2">
                        <asp:Button ID="BtnLogin" CssClass="myButtonAct" runat="server" Text="Login" Width="116px" Height="40.5px" Font-Names="Tahoma" Font-Size="12px" />
                    </td>
                </tr>
            </table>
        </ContentTemplate>
        </asp:UpdatePanel>
    </div>
    </form>
</body>
</html>