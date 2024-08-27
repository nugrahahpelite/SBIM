<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfCrpViewer.aspx.vb" Inherits="SBSto.WbfCrpViewer" %>

<%@ Register assembly="CrystalDecisions.Web, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" namespace="CrystalDecisions.Web" tagprefix="CR" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table class="myPanelGreyLight">
                <tr>
                    <td>
                        <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Visible="False"></asp:Label>
                        <CR:CrystalReportViewer ID="Crv" runat="server" AutoDataBind="true" PrintMode="ActiveX" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
