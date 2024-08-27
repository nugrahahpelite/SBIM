<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfCrpViewerSmk.aspx.vb" Inherits="SBSto.WbfCrpViewerSmk" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="~/CssFiles/CssBnsrp.css" rel="stylesheet" type="text/css" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button ID="BtnBackTrans" class="myButtonFind" runat="server" Text="Kembali ke Daftar Laporan" Width="239px" Font-Names="Tahoma" Font-Size="12px" BackColor="#CCCCCC" BorderStyle="None" Height="30px" />
            <br />
            <iframe src="WbfCrpViewer.aspx" style="position:fixed;overflow:hidden;width:100%; height: 100%;"></iframe>
        </div>
    </form>
</body>
</html>
